.PHONY: install lint typecheck test acceptance verify examples check parity office-oracles graphics-uno smartart-uno property-campaign clean
install:
	bun install --frozen-lockfile
lint: typecheck
typecheck:
	bun run typecheck
test:
	bun test tests/unit
acceptance:
	bun run acceptance
verify:
	bun run verify
examples:
	bun run examples/agent-edit.ts
	bun run examples/create-office.ts
	bun run examples/review-word.ts
check:
	bun run typecheck
	bun run scripts/test-inventory.ts --check
	bun run scripts/mapping-reconciliation.ts --check
	bun run scripts/outcome-mappings.ts --check
	bun run verify
	bun run scripts/coverage.ts --check
	bun run scripts/shared-contracts.ts
	bun test tests/unit
	bun run acceptance
	bun run examples/agent-edit.ts
	bun run examples/create-office.ts
	bun run examples/review-word.ts
	bun run verify
	bun run scripts/shared-contracts.ts
# Independent development-only validators, never called by runtime APIs.
office-oracles:
	bun run scripts/office-oracles.ts
# Optional local LibreOffice server oracle; requires python3-uno and LibreOffice.
graphics-uno:
	bun run scripts/graphics-sample.ts
	timeout --kill-after=5s 180s /usr/bin/python3 scripts/oracles/graphics-uno.py
smartart-uno:
	bun run scripts/smartart-office-sample.ts
	timeout --kill-after=5s 180s /usr/bin/python3 scripts/oracles/smartart-uno.py
property-campaign:
	bun run scripts/property-campaign.ts 20260926 256 32
	bun run scripts/property-campaign.ts 8675309 256 32
	bun run scripts/performance-campaign.ts
parity: check
	bun run scripts/verify.ts --full
	bun run scripts/acceptance.ts --full
clean:
	rm -rf artifacts
