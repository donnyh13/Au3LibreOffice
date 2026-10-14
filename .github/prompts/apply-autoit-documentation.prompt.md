---
name: Apply AutoIt documentation improvements
description: Apply approved AutoIt documentation improvements
---

# Apply AutoIt Documentation Improvements

Improve the documentation of the selected AutoIt code according to the repository's documentation instructions.

## Critical Constraint

Modify comments/documentation only.

Do NOT modify executable AutoIt code.

Do NOT:
    - Refactor code.
    - Rename variables.
    - Change function behavior.
    - Change expressions.
    - Change control flow.
    - Change whitespace unnecessarily.
    - Fix unrelated bugs.
    - Change APIs.

If you believe executable code needs to change, report the issue separately instead of changing it.

## Documentation Principles

Prefer comments that explain:
    - Why something is done.
    - Important assumptions.
    - Non-obvious behavior.
    - Side effects.
    - API contracts.
    - Error behavior.
    - Resource ownership.
    - Initialization requirements.
    - Required calling order.
    - Workarounds.
    - Compatibility requirements.
    - Important edge cases.
    - Information that a future maintainer could not easily infer from the implementation.

Do not add comments that merely narrate obvious code.

## Existing Comments

Preserve accurate and useful comments.

Improve comments that contain useful information but are unclear.

Correct comments that are demonstrably inaccurate.

Remove comments only when they are clearly:
    - Redundant.
    - Obvious.
    - Stale.
    - Misleading.
    - Completely uninformative.

If a comment appears to contain historical knowledge that cannot be verified from the code, preserve it and flag it for human review rather than deleting it.

## Public Functions

For public functions, improve documentation where needed so that a caller can understand:
    - Purpose.
    - Parameters.
    - Return value.
    - Important failure behavior.
    - Important side effects.
    - Important restrictions.
    - Required initialization or calling order.

Do not add boilerplate documentation when the information is already obvious.

## Comment Style

Keep comments concise.

Write for a future developer, not for a beginner learning AutoIt.

Prefer:

> Why this is necessary.

over:

> Do this thing because the next line does this thing.

## Verification

After making the documentation changes:
    - Review the resulting diff.
    - Confirm that no executable code changed.
    - Identify comments whose accuracy depends on information unavailable in the source.
    - Identify any code that appears unusual but whose purpose remains uncertain.

Report those uncertain items separately.

## Change Verification

Before editing:

    - Identify the exact functions and documentation regions in scope.
    - Preserve the original source as the comparison baseline.
    - Do not apply findings marked VERIFY as factual comments without additional evidence or author approval.

After editing:

    - Inspect the complete diff.
    - Confirm that no executable statements, expressions, directives, variable names, function signatures, or control-flow structures changed.
    - Confirm that no unrelated source regions were reformatted or altered.
    - Check that each new or revised comment is supported by evidence.
    - Check that comments remain adjacent to the code they explain.
    - Check that public headers remain consistent with the implementation.
    - Report any verification that could not be completed.

If a non-documentation change appears in the diff, revert that change without discarding approved documentation changes, or stop and report the issue if it cannot be safely isolated.

Do not claim that the source is unchanged merely because the intended task was documentation-only. Verify the actual diff.
