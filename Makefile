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
	bun run examples/review-word.ts
check:
	bun run typecheck
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
parity: check
	bun run scripts/verify.ts --full
	bun run scripts/acceptance.ts --full
clean:
	rm -rf artifacts
