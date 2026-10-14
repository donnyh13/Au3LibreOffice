---
name: Audit AutoIt UDF Headers
description: Audit an AutoIt UDF function header against its implementation
---

# Audit AutoIt Function Header

Compare the selected public AutoIt function's #FUNCTION# documentation block against its actual implementation.

Do not modify the source.

## Check

### Name

Confirm that the documented function name matches the implementation.

### Description

Determine whether the description accurately describes the function's purpose.

Do not rewrite the description merely for stylistic reasons.

### Syntax

Compare:
    - Parameter order.
    - Parameter names.
    - Default values.
    - Optional parameters.

Flag discrepancies.

### Parameters

For every parameter verify:
    - Name.
    - Type.
    - Default value.
    - Accepted values.
    - Restrictions.
    - Special behavior.
    - Side effects.
    - ByRef behavior when relevant.

### Return Values

Trace every return path.

For each successful return determine:
    - Return type.
    - Meaning of the return value.
    - @error.
    - @extended.

For each failure path determine:
    - @error.
    - @extended.
    - Meaning of the error.

Compare the findings with the documented Return values section.

### Remarks

Check whether important caller-visible behavior is documented, including:
    - Side effects.
    - Automatic behavior.
    - Default behavior.
    - Special cases.
    - Required calling order.
    - Resource ownership.
    - State changes.
    - Restrictions.
    - Important implementation behavior that callers need to understand.

Do not require Remarks to document internal implementation details that callers do not need to know.

### Related

Identify obviously related public functions that are missing from the Related section, but do not require exhaustive cross-linking.

### Example

Determine whether the function appears to have behavior complex enough that its example is particularly important.

Do not create an example unless explicitly requested.

## Output

Use these categories:
    - ACCURATE
    - IMPROVEMENT
    - HEADER-MISMATCH
    - MISSING-DOCUMENTATION
    - VERIFY

For each finding provide:
    - Location.
    - Category.
    - Current documentation.
    - Actual behavior.
    - Recommended change.
    - Confidence.

Do not change the source.

## Error-Path Verification

Do not infer the complete error contract from the presence of SetError() calls alone.

For each relevant path:

    - Trace the value returned to the caller.
    - Trace the effective @error and @extended values at the point the function returns.
    - Account for intervening calls that may change those values.
    - Distinguish explicit error handling from behavior supplied by AutoIt or the underlying COM API.
    - Compare the resulting contract with the header.

If the status values cannot be established confidently from the available code, report the uncertainty rather than asserting a mismatch.
