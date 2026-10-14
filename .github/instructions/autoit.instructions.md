---
applyTo: "**/*.au3"
---

# AutoIt-Specific Instructions

## General

This repository contains AutoIt source code.

Preserve valid existing AutoIt behavior unless a code change is explicitly requested.

When documenting AutoIt code, account for AutoIt's runtime behavior rather than assuming behavior from another programming language.

## AutoIt-Specific Documentation

Pay particular attention to:
    - ByRef parameters.
    - Global variables.
    - Static variables.
    - #include dependencies.
    - Functions whose behavior depends on initialization performed elsewhere.
    - DllCall() and Windows API interactions.
    - Handles and resource ownership.
    - DLL structures and pointer-related code.
    - Windows error codes.
    - @error and @extended.
    - Functions that modify @error or @extended.
    - Functions whose return value has multiple meanings depending on @error.
    - Variant/type conversion behavior.
    - Arrays and array bounds.
    - String manipulation where edge cases matter.
    - COM objects.
    - Windows registry operations.
    - File and process handles.
    - GUI handles and GUI state.
    - Callbacks.
    - Timers.
    - Event-driven behavior.
    - Environment-dependent behavior.
    - OS-version-dependent behavior.
    - Architecture-dependent behavior, particularly 32-bit versus 64-bit behavior.
    - ANSI versus Unicode behavior when relevant.
    - API structures whose layout or size is significant.
    - Workarounds for AutoIt or Windows limitations.

## Global State

If a function reads or modifies global state, consider whether that behavior is important to document for callers or maintainers.

Do not automatically add comments simply because a global variable is used.

## @error and @extended

Treat @error and @extended as part of a function's effective API when callers depend on their values.

When auditing documentation, identify functions where:
    - @error communicates failure information.
    - @extended communicates additional result information.
    - The return value alone is insufficient to interpret the result.
    - A function resets, preserves, or intentionally modifies these values.

## DLL/API Calls

For DllCall() and Windows API code, prioritize documentation of:
    - Why a particular API is being called.
    - Important parameter conventions.
    - Structure layout assumptions.
    - Required data types.
    - Handle ownership.
    - Required cleanup.
    - Important error behavior.
    - Architecture-specific considerations.
    - Non-obvious conversions.
    - Known Windows API quirks or workarounds.

Do not add comments explaining obvious syntax such as the fact that DllCall() calls a DLL.

## Unusual AutoIt Behavior

If code depends on behavior that may surprise a developer coming from another language, consider documenting it when the behavior materially affects correctness.

Do not make claims about AutoIt behavior unless they can be established from the code, project documentation, or authoritative documentation available to Copilot.

## Includes and Dependencies

When a function depends on declarations, constants, helper functions, or initialization from another included UDF, document the dependency when it would not be obvious to a maintainer.

Do not duplicate information that is already clearly documented at the appropriate higher level.

## Do Not Over-Document

Do not add comments for:
    - Simple assignments.
    - Straightforward If statements.
    - Obvious loops.
    - Simple return statements.
    - Obvious function calls.
    - Basic array access.
    - Obvious string operations.

The presence of AutoIt syntax does not by itself justify a comment.

## Status Propagation and COM Behavior

When documenting a function that uses @error or @extended:

    - Inspect the complete expression and control-flow path surrounding each relevant status check.
    - Determine whether an intervening function call, COM operation, or helper can change the status before it is checked.
    - Do not assume that a status value produced by one operation remains available after another operation.
    - Distinguish AutoIt's status reporting from errors or exceptions exposed by the underlying COM API.
    - Do not assume that a COM method's return value, an object's validity, or the success of an operation can be established from @error alone.

When documenting COM object handling:

    - Identify whether an object is acquired, created, reused, returned to a caller, or released.
    - Distinguish releasing a local object reference from closing a document or terminating LibreOffice.
    - Document object lifetime or ownership only when supported by the implementation and relevant API semantics.
    - Treat cursor, selection, controller, active-document, and document-modification state as separate concepts; do not use their names interchangeably.

Do not add these explanations mechanically. Document them when they affect the public contract or make the implementation safer to maintain.
