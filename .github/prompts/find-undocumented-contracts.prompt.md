---
name: Find undocumented AutoIt UDF API contracts
description: Find undocumented AutoIt UDF API contracts and hidden assumptions
---

# Find Undocumented Contracts

Analyze the selected AutoIt UDF as if you are a developer who must maintain it without access to the original author's knowledge.

Do not modify the source.

Look specifically for behavior that effectively forms part of the UDF's API but is not clearly documented.

## Look For

### Parameters

Identify parameters with:
    - Non-obvious valid values.
    - Special sentinel values.
    - Values that change behavior.
    - Required combinations.
    - Restrictions.
    - Ownership implications.
    - ByRef behavior.

### Return Values

Identify functions where:
    - The return value has multiple meanings.
    - A special value represents failure.
    - A special value represents a valid result.
    - The return value must be interpreted together with @error.
    - The return value must be interpreted together with @extended.

### Side Effects

Identify functions that:
    - Modify global variables.
    - Modify caller-provided variables through ByRef.
    - Change process state.
    - Change GUI state.
    - Create or destroy resources.
    - Open or close handles.
    - Modify files, registry values, or environment state.
    - Modify @error or @extended.

### Calling Contracts

Identify functions that:
    - Must be called after initialization.
    - Must be called before cleanup.
    - Must be called in a particular order.
    - Depend on another function having been called.
    - Require a particular global state.

### Resource Contracts

Identify resources that:
    - Must be released.
    - Must remain valid for a particular period.
    - Are owned by the caller.
    - Are owned by the UDF.
    - Are transferred between functions.

### Error Contracts

Identify:
    - Expected failures.
    - Recoverable failures.
    - Special error values.
    - Retry behavior.
    - Silent failure.
    - Error information returned through @error or @extended.

### Compatibility Contracts

Identify behavior that appears intended to support:
    - Different Windows versions.
    - Different architectures.
    - 32-bit versus 64-bit processes.
    - Different AutoIt versions.
    - Different API versions.
    - Legacy callers.

### Hidden Assumptions

Look for assumptions that are not obvious from individual functions but emerge from interactions between functions.

### Output

For each finding provide:
    - Location.
    - Contract category.
    - Description.
    - Evidence in the code.
    - Why a caller or maintainer needs to know.
    - Whether documentation should be added.
    - Confidence.

If the reason cannot be established from the source, explicitly mark it as requiring author verification.

Do not invent undocumented contracts.

## Cross-Function Contracts

When a possible contract depends on multiple functions, trace the relevant relationship rather than treating each function independently.

Look for:

    - Initialization functions and the state they establish.
    - Functions that require an object or document created by another function.
    - Functions that change the active document, controller, selection, or cursor used by subsequent calls.
    - Functions that create, return, replace, or release COM object references.
    - Cleanup functions and the resources they affect.
    - Shared globals or cached objects whose values are established or invalidated elsewhere.
    - Error conventions that callers must follow across a sequence of operations.

Distinguish a demonstrated dependency from a plausible one. A sequence used in an example does not by itself prove that the sequence is mandatory.

When the dependency cannot be confirmed, identify the functions involved and explain what remains uncertain.
