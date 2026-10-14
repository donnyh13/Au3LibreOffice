AutoIt UDF — Copilot Instructions

## Project Context

This repository contains an AutoIt UDF (User Defined Function) intended to be maintained and reused by other AutoIt scripts.

Treat the UDF as library code rather than application code. Public functions are an API and their behavior should be considered potentially relied upon by external callers.

When working with this repository, preserve existing behavior unless the user explicitly asks for a behavioral change.

## General Rules

    - Do not modify executable code when the task is specifically about documentation unless explicitly instructed to do so.
    - Do not invent behavior, intent, requirements, or historical reasons for code.
    - If the purpose of code cannot be determined confidently, say so and flag it for human review.
    - Preserve useful existing comments.
    - Treat existing comments as potentially stale or incorrect; verify them against the code before relying on them.
    - Prefer explaining WHY something exists over describing WHAT the code literally does.
    - Do not add comments merely to increase comment density.
    - Do not comment obvious syntax or self-explanatory operations.
    - Do not use comments as a substitute for clear naming when renaming would be appropriate, unless code changes were explicitly requested.
    - Avoid verbose comments that merely restate the implementation.
    - Do not make speculative claims in comments.

## Evidence and Confidence

Every proposed comment must be supported by available evidence.

Use the following evidence sources, in order of preference:

    1. The implementation being documented.
    2. Related functions, callers, and tests in the repository.
    3. Existing project documentation and examples.
    4. Authoritative AutoIt, Windows, or LibreOffice API documentation, when needed.

Distinguish confirmed behavior from inferred intent.

The implementation may establish what an operation does without establishing why the original author chose it.

Do not turn an inference about intent into a factual comment. If the reason is uncertain and materially affects safe maintenance, classify it as VERIFY.

When external documentation is unavailable, do not claim that behavior has been independently verified.

For non-obvious claims, identify the evidence supporting the finding. Prefer a specific code location, caller, test, or API reference over a general assertion.

## What Makes a Useful Comment

A comment is valuable when it communicates information that a competent programmer could not easily determine from the nearby code.

Prioritize comments explaining:
    - Why a non-obvious implementation was chosen.
    - Important assumptions.
    - Required calling order.
    - Side effects.
    - Global state that is modified or relied upon.
    - Resource ownership and cleanup requirements.
    - Windows API behavior.
    - AutoIt-specific behavior that is not obvious.
    - Workarounds for operating-system, API, AutoIt, or library limitations.
    - Compatibility requirements.
    - Error-handling behavior.
    - Non-obvious return-value conventions.
    - Performance-sensitive decisions.
    - Edge cases.
    - Conditions under which apparently unusual code is intentional.
    - Relationships between multiple functions that are not obvious from any one function.
    - Behavior that callers must know to use the function correctly.

## Comments That Usually Add Little Value

Avoid comments such as:

> $i += 1 ; Increment i
>
> ; Loop through the array
> For $i = 0 To UBound($aItems) - 1
>
> ; Return the result
> Return $iResult

Do not add comments that simply translate AutoIt syntax into English.

Comments describing straightforward control flow, assignments, function calls, or obvious variable manipulation are generally unnecessary.

## Public UDF Functions

For public functions, documentation should generally make it possible for another developer to understand:
    - What the function is for.
    - What each parameter means when its purpose is not obvious.
    - Valid parameter values or important restrictions.
    - What the function returns.
    - Important error or failure behavior.
    - Important side effects.
    - Required initialization or calling order.
    - Important resource ownership or cleanup requirements.
    - Important examples or usage constraints when needed.

Do not add boilerplate documentation to every parameter when the parameter's meaning is completely obvious.

When documenting parameter and return types, distinguish between:

    - Types explicitly constrained or validated by the implementation.
    - Types expected by the underlying LibreOffice API.
    - Types merely observed in existing callers or examples.

Do not treat these as interchangeable.

Trace all relevant return paths, including early returns and SetError() calls. Check whether helper functions or API calls alter @error or @extended before assuming those values are preserved.

Document special return values only when their meaning is supported by the implementation or authoritative documentation. If the meaning remains ambiguous, flag it for verification.

## Private/Internal Functions

Internal helper functions do not automatically require extensive documentation.

Document an internal function when:
    - Its purpose is not obvious from its name and implementation.
    - It contains important assumptions.
    - It has non-obvious side effects.
    - It implements a workaround.
    - Its behavior is relied upon by several other functions.
    - It has unusual error or return-value behavior.
    - A future maintainer would reasonably need additional context to modify it safely.

## Existing Comments

When reviewing an existing comment, classify it as one of:
    - KEEP — Accurate and useful.
    - IMPROVE — Useful idea, but unclear, incomplete, or poorly expressed.
    - REPLACE — The comment is useful but materially misleading or inaccurate.
    - REMOVE — Redundant, obvious, obsolete, or otherwise provides no useful information.
    - VERIFY — The comment may contain important historical or behavioral information, but the code alone is insufficient to confirm it.

Never remove a comment containing potentially important historical information merely because the reason cannot be inferred from the current code. Flag it for human verification instead.

## Scope and Completeness

Work only on the explicitly selected file, function, or range unless the user authorizes a broader scope.

For multi-function or repository-wide tasks, establish the scope before beginning and process the code in reviewable units.

Do not claim that an entire file or repository has been audited unless the relevant scope was actually examined.

If context limits, inaccessible dependencies, or time constraints prevent a complete review, state what was examined and what remains unchecked.

## Uncertain Intent

When the code contains something unusual and the reason cannot be determined:
    - Do not invent an explanation.
    - Do not write a speculative comment.
    - Flag the location for human review.
    - Explain what is unusual and what information appears to be missing.

For example:

>     VERIFY: This code intentionally performs X instead of the apparently simpler Y. The reason cannot be established from the surrounding code.

## Documentation Style

Comments should be:
    - Concise.
    - Specific.
    - Technically accurate, but easy to understand for general coders.
    - Written for a future maintainer.
    - Written in the same general style as the existing UDF documentation.
    - Focused on information that is difficult to infer from the implementation.

Avoid conversational filler, generic statements, and comments that merely narrate execution.

## Safety During Documentation Work

When asked to audit or improve documentation:
    - Do not change executable statements.
    - Do not change variable names.
    - Do not refactor functions.
    - Do not change whitespace unless necessary for the documentation change.
    - Do not change behavior.
    - Do not "fix" apparent code problems unless explicitly requested.
    - If you discover a probable code defect, report it separately rather than silently fixing it.

## Review Output

When performing a documentation audit, distinguish between:
    - Confirmed documentation problem.
    - Strong documentation recommendation.
    - Possible issue requiring human verification.

Always provide the reason for a recommendation.

The goal is not maximum comment coverage.

The goal is maximum useful information for the next developer who has to understand, maintain, debug, or extend this UDF.

## Existing UDF Function Headers

This project uses a standardized #FUNCTION# documentation block before public functions.

The function header is part of the public API documentation and should be treated separately from inline implementation comments.

A typical function header contains:
    - Name
    - Description
    - Syntax
    - Parameters
    - Return values
    - Author
    - Modified
    - Remarks
    - Related
    - Link
    - Example

When reviewing a public function, evaluate both:
    1. Whether the function header accurately describes the public API.
    2. Whether inline comments adequately explain non-obvious implementation details.

Do not automatically rewrite an entire function header merely to improve wording.

## Function Header Accuracy

Check whether the header accurately reflects the implementation.

Pay particular attention to:
    - Parameter types.
    - Default values.
    - Special parameter values.
    - Return types.
    - @error.
    - @extended.
    - Side effects.
    - Object ownership.
    - State changes.
    - Resource cleanup.
    - Required initialization.
    - Cursor or selection changes.
    - Conditions under which special behavior occurs.
    - Remarks that describe behavior not obvious from the implementation.

If the implementation and header appear inconsistent, flag the discrepancy rather than silently changing either one.

Use the category HEADER-MISMATCH for these findings.

## Public API Documentation

For public functions, distinguish between:

### API documentation

Information a caller needs to correctly use the function.

This belongs primarily in:
    - Description
    - Syntax
    - Parameters
    - Return values
    - Remarks
    - Related
    - Example

### Implementation documentation

Information a maintainer needs to understand why the implementation works the way it does.

This belongs primarily in inline comments.

Do not move implementation details into the public API documentation unless callers need to know them.

## Inline Comment Standard

For inline comments, prefer this hierarchy:
    - Explain WHY.
    - Explain a non-obvious consequence.
    - Explain an important assumption.
    - Explain an unusual API/runtime behavior.
    - Explain a non-obvious algorithmic step.
    - Only explain WHAT the code does when the operation itself is sufficiently complex that the purpose is not apparent.

For example, this is weak:

> ; Get the current selection.
> $oSelection = $oDoc.getCurrentSelection()

This is more useful:

> ; Save the current selection so it can be restored after the table is converted.
> $oSelection = $oDoc.getCurrentSelection()

The second comment explains why the operation matters.

## Documentation Consistency

Maintain the existing project's terminology and capitalization where practical.

In particular, distinguish consistently between:
    - Document Object
    - Table Object
    - Cursor Object
    - ViewCursor
    - Document path
    - LibreOffice URL
    - Windows path
    - @error
    - @extended

Do not introduce a new documentation style merely because it is more common in other AutoIt projects.

## Do Not Over-Document Standardized Headers

Do not add repetitive inline comments merely because a function has a large header.

The header documents the public API.

Inline comments should primarily document implementation details that are not already covered by the header.

## Project Error Handling Convention

This UDF uses @error and @extended as a structured status system.

Generally:
    - @error = 0 indicates success.
    - @error = 1 indicates an input error.
    - @error = 2 indicates an initialization error.
    - @error = 3 indicates a processing error.
    - @error = 4 indicates a property setting error.
    - @error = 5 indicates a printer related error.
    - @error = 6 indicates a LibreOffice version related error.
    - @extended identifies the specific condition within that category.

Do not describe @extended merely as an arbitrary number.

When documenting a function, verify that its Return values section accurately corresponds to every SetError() path in the implementation.

When auditing a function:
    1. Enumerate every SetError() call.
    2. Compare its @error and @extended values against the function header.
    3. Check whether any documented error condition is unreachable or missing.
    4. Check whether any reachable error condition is undocumented.
    5. Check whether success paths correctly document their @extended values.
    6. Check whether the documented return value agrees with the actual third argument passed to SetError().

If discrepancies are found, classify them as HEADER-MISMATCH.

Do not automatically change the implementation to match the documentation.

Do not automatically change the documentation to match the implementation if the intended behavior is ambiguous.

Report the discrepancy for human review.

## State Changes and Restoration

Treat temporary changes to external or document state as potentially important behavior.

In particular, identify code that:
    - Changes the current selection.
    - Changes the ViewCursor.
    - Changes the active document.
    - Changes the active controller.
    - Changes GUI state.
    - Changes document state.
    - Changes file state.
    - Acquires or releases resources.
    - Temporarily changes an object property and later restores it.

Document the reason for saving/restoring state when that reason is not obvious.

For state-changing operations, determine whether the function header already documents the behavior.

Do not add redundant inline comments if the behavior is already adequately described in the public function's Remarks section.

However, document the implementation mechanism inline when a future maintainer might otherwise accidentally remove or bypass the restoration.

Pay particular attention to code that follows this pattern:
    1. Save state.
    2. Change state.
    3. Perform an operation.
    4. Restore state.

The relationship between steps 1 and 4 is often worth documenting even when each individual operation is obvious.
