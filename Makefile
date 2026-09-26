.PHONY: install lint typecheck test acceptance verify examples check parity clean
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
check:
	bun run typecheck
	bun run verify
	bun run scripts/coverage.ts --check
	bun run scripts/office-contracts.ts
	bun run scripts/shared-pack.ts
	bun run scripts/symbols.ts --check
	bun test tests/unit
	bun run acceptance
	bun run examples/agent-edit.ts
	bun run examples/create-office.ts
	bun run verify
	bun run scripts/shared-pack.ts
parity: check
	bun run scripts/verify.ts --full
	bun run scripts/acceptance.ts --full
clean:
	rm -rf artifacts
