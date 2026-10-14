---
name: Audit AutoIt UDF
description: Audit AutoIt UDF documentation without modifying code
--- 

# AutoIt UDF Documentation Audit

Perform a documentation audit of the selected AutoIt code, or if none is selected ask if the user wants to review the entire file function by function, or in another appropriate method.

## Critical Constraint

Do NOT modify the source code or comments.

This is an audit only.

Do not generate a rewritten version of the source file.

## Objective

Evaluate whether the code contains the information a future maintainer would need to safely understand, debug, modify, and extend it.

The objective is NOT to maximize the number of comments.

A well-documented function does not necessarily need many comments.

## Analyze Existing Comments

For each potentially problematic comment, determine whether it should be:
    - KEEP — Accurate and useful.
    - IMPROVE — Useful but unclear or incomplete.
    - REPLACE — Useful concept, but materially inaccurate or misleading.
    - REMOVE — Redundant, obvious, stale, or otherwise not useful.
    - VERIFY — Potentially important information that cannot be confirmed from the code.

Do not recommend removing a comment merely because the reason behind it is not obvious.

## Identify Missing Documentation

Look for places where documentation would materially help a future maintainer.

Prioritize:
    1. Public API behavior.
    2. Important parameters and restrictions.
    3. Return values and error conventions.
    4. @error and @extended.
    5. Side effects.
    6. Global state.
    7. Resource ownership and cleanup.
    8. Initialization requirements.
    9. Required calling order.
    10. Windows API behavior.
    11. DllCall() behavior.
    12. Non-obvious AutoIt behavior.
    13. Workarounds.
    14. Compatibility behavior.
    15. Architecture-specific behavior.
    16. Error recovery.
    17. Edge cases.
    18. Performance-sensitive implementation decisions.
    19. Non-obvious relationships between functions.
    20. Intentional-looking code that would otherwise appear to be an error.

## Avoid False Positives

Do NOT recommend comments for code that is self-explanatory.

Examples include:
    - Return $result
    - Incrementing a counter.
    - A straightforward loop over an array.
    - A simple assignment.
    - A simple conditional.
    - An obvious function call.
    - Straightforward string manipulation.

A comment should provide information that cannot be easily inferred from the nearby implementation.

## Detect Potentially Dangerous Missing Context

Pay particular attention to code that looks unnecessarily complicated or unusual.

For each such case, ask:

    "Would a future developer be tempted to simplify or change this code because they do not understand why it exists?"

If yes, it is a high-priority documentation candidate.

Examples include:
    - Unusual ordering.
    - Apparently redundant operations.
    - Strange constants.
    - Deliberate delays.
    - Retry loops.
    - Apparently unnecessary API calls.
    - Workarounds.
    - Special-case handling.
    - Defensive checks.
    - Unusual type conversions.
    - Strange return-value handling.
    - Code that appears to duplicate another operation.

Do not invent the reason for the code.

If the reason cannot be established, classify it as VERIFY.

## Classification

Assign each finding one of:

### HIGH

A future maintainer could reasonably break behavior by misunderstanding the undocumented behavior.

### MEDIUM

The code is maintainable now, but useful context is missing.

### LOW

Documentation could improve readability, but the missing information is unlikely to cause incorrect modifications.

### VERIFY

The code contains something important or unusual whose intent cannot be established confidently.

## Output

Produce a table with:
Location | Category | Priority | Evidence | Finding | Recommended action
Function and line or code anchor | KEEP, IMPROVE, REPLACE, REMOVE, VERIFY, or missing documentation | HIGH, MEDIUM, LOW, or VERIFY | What the implementation or other source establishes | What is wrong or missing | What should change and why

For comments that should be changed, quote only the relevant short comment.

For missing documentation, describe what information should be documented rather than automatically writing the final comment.

## Final Summary

After the table, provide:

### Documentation strengths

List the strongest aspects of the existing documentation.

### Documentation weaknesses

List recurring problems or patterns.

### High-risk undocumented behavior

List the areas most likely to cause a future maintainer to accidentally change behavior.

### VERIFY items

List things where the original author's intent is needed.

## Overall assessment

Describe the general documentation state in neutral terms.

Do not assign a numerical score.

## Future-Maintainer / Refactoring Risk

Identify code that a future maintainer might incorrectly simplify, remove, reorder, or replace because the reason for its existence is not obvious.

Examples include:
    - Sleep() calls.
    - Repeated API calls.
    - Apparently redundant checks.
    - Repeated object retrieval.
    - Unusual regular expressions.
    - Temporary state changes.
    - Save/restore operations.
    - Object invalidation checks.
    - Explicit Null assignments.
    - Special handling of unsaved documents.
    - Special handling of hidden documents.
    - Type checks that appear redundant.
    - Defensive checks.
    - Retry logic.
    - Deliberate delays.
    - Comparisons against apparently unusual values.
    - Calls whose return values are intentionally ignored.
    - Calls whose only apparent purpose is to trigger side effects.
    - Workarounds around LibreOffice behavior.
    - Workarounds around AutoIt behavior.
    - Workarounds around Windows behavior.

For each such case, determine whether the reason is:
    1. Obvious from the surrounding code.
    2. Explained by an existing comment.
    3. Explained by the function header.
    4. Explained elsewhere in the UDF.
    5. Not documented and should be documented.
    6. Unknown and requires author verification.

Do not invent the reason.

A comment should be recommended when removing or changing the unusual code without understanding its purpose could plausibly introduce a regression.

## Audit Boundaries and Finding Quality

Audit only the selected code or specified code range, and the directly relevant dependencies needed to understand its behavior.

For each finding:
    - Identify a specific line, comment, function header, or code range.
    - Explain the concrete maintenance or caller risk.
    - Cite the relevant implementation evidence.
    - State whether the recommendation is supported by code, external documentation, or inference.
    - Separate a documentation defect from a possible implementation defect.

Do not create multiple findings for the same underlying problem unless they affect distinct contracts or locations.

Do not recommend comments solely because a code pattern appears in a checklist. First determine whether the pattern is already adequately explained or whether the missing context materially affects safe maintenance.

For VERIFY findings, state exactly what cannot be established and what evidence or author input would resolve the uncertainty.

If the selected scope is too large to inspect reliably in one pass, audit it in smaller units and report the coverage of each unit.

Do not infer that an undocumented behavior is intentional merely because it is unusual.
