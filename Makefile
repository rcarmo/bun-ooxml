# Resolve once before child TMPDIR; portable without the workspace Makefile.
PROJECT_TMP := $(shell bun scripts/project-tmp.ts)
ifeq ($(strip $(PROJECT_TMP)),)
$(error Project temp resolution failed; invalid explicit overrides never fall back)
endif
export PROJECT_TMP_ROOT := $(PROJECT_TMP)
export PROJECT_TMP_BASE := $(dir $(PROJECT_TMP))
export OOXML_TMP_ROOT := $(PROJECT_TMP)
export PROJECT_TEST_ROOT := $(PROJECT_TMP)/tests
export PROJECT_LOG_ROOT := $(PROJECT_TMP)/logs
export BUN_INSTALL_CACHE_DIR := $(PROJECT_TMP)/cache/bun
export XDG_CACHE_HOME := $(PROJECT_TMP)/cache/xdg
export NPM_CONFIG_CACHE := $(PROJECT_TMP)/cache/npm
export DOTNET_CLI_HOME := $(PROJECT_TMP)/cache/dotnet-home
export NUGET_PACKAGES := $(PROJECT_TMP)/cache/nuget
export PYTHONPYCACHEPREFIX := $(PROJECT_TMP)/cache/python
export OOXML_BUILD_DIR := $(PROJECT_TMP)/build
export DOTNET_CLI_TELEMETRY_OPTOUT := 1
export DOTNET_SKIP_FIRST_TIME_EXPERIENCE := 1
# Base scratch for installer; each dev-run owns an isolated runs subtree.
MAKE_RUN_ID := $(shell bun -e 'console.log(crypto.randomUUID())')
export TMPDIR := $(PROJECT_TMP)/runs/install/$(MAKE_RUN_ID)/tmp
ORACLE_OUTPUT ?= $(abspath artifacts/policy-oracles/$(MAKE_RUN_ID))
export TMP := $(TMPDIR)
export TEMP := $(TMPDIR)
DEV := bun scripts/dev-run.ts
.PHONY: prepare install lint typecheck test acceptance verify examples check parity office-oracles graphics-uno smartart-uno property-campaign clean
prepare:
	mkdir -p "$(TMPDIR)" "$(BUN_INSTALL_CACHE_DIR)" "$(OOXML_BUILD_DIR)"
install: prepare
	bun install --frozen-lockfile
lint: typecheck
typecheck: prepare
	$(DEV) typecheck tool bun node_modules/typescript/bin/tsc --noEmit
test: prepare
	bun scripts/unit-batch.ts
acceptance: prepare
	$(DEV) acceptance run scripts/acceptance.ts
verify: prepare
	$(DEV) verify run scripts/verify.ts
examples: prepare
	$(DEV) example-edit run examples/agent-edit.ts
	$(DEV) example-create run examples/create-office.ts
	$(DEV) example-review run examples/review-word.ts
check: typecheck
	$(DEV) inventory run scripts/test-inventory.ts --check
	$(DEV) reconciliation run scripts/mapping-reconciliation.ts --check
	$(DEV) outcomes run scripts/outcome-mappings.ts --check
	$(MAKE) verify
	$(DEV) coverage run scripts/coverage.ts --check
	$(DEV) shared-contracts run scripts/shared-contracts.ts
	$(MAKE) test acceptance examples verify
	$(DEV) shared-contracts-final run scripts/shared-contracts.ts
# Independent development-only validators; runtime APIs never invoke them.
office-oracles: prepare
	OOXML_ORACLE_OUTPUT="$(ORACLE_OUTPUT)/office" $(DEV) office-oracles run scripts/office-oracles.ts
graphics-uno: prepare
	$(DEV) graphics-sample run scripts/graphics-sample.ts "$(ORACLE_OUTPUT)/graphics"
	$(DEV) graphics-uno tool timeout --kill-after=5s 180s /usr/bin/python3 scripts/oracles/graphics-uno.py "$(ORACLE_OUTPUT)/graphics"
smartart-uno: prepare
	$(DEV) smartart-sample run scripts/smartart-office-sample.ts "$(ORACLE_OUTPUT)/smartart"
	$(DEV) smartart-uno tool timeout --kill-after=5s 180s /usr/bin/python3 scripts/oracles/smartart-uno.py "$(ORACLE_OUTPUT)/smartart"
property-campaign: prepare
	$(DEV) property-one run scripts/property-campaign.ts 20260926 256 32
	$(DEV) property-two run scripts/property-campaign.ts 8675309 256 32
	$(DEV) performance run scripts/performance-campaign.ts
parity: check
	$(DEV) verify-full run scripts/verify.ts --full
	$(DEV) acceptance-full run scripts/acceptance.ts --full
# Requires deliberate idle confirmation; never deletes artifacts/ or evidence.
clean:
	@test "$(CONFIRM_IDLE)" = yes || { echo 'Set CONFIRM_IDLE=yes after verifying no active jobs'; exit 1; }
	@test "$$(basename "$(PROJECT_TMP)")" = bun-ooxml
	@test ! -L "$(PROJECT_TMP)"
	rm -rf -- "$(PROJECT_TMP)/cache" "$(PROJECT_TMP)/build" "$(PROJECT_TMP)/tests" "$(PROJECT_TMP)/logs" "$(PROJECT_TMP)/runs"
