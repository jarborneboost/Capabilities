---
name: extract-requirements
description: Turn a client document into a reviewed list of requirements. Only use when called with `/extract-requirements`.
---

# Requirements from client documents

## Steps

1. Read the whole document. Every paragraph, list item and table row that states an obligation becomes a requirement; split compound statements into one requirement each.
2. For each requirement, fill in the three checks below.
3. Present the result as one table, one row per requirement, in document order.

Done when every obligation in the document has a row and every row has all three checks filled in.

## The three checks

**Traceability** – Quote the original wording verbatim and give its exact location (section, page, table and row), so anyone can check the requirement against the source.

**ISTQB quality** – Assess the requirement against complete, consistent, correct, feasible, necessary, prioritised, unambiguous and verifiable. For each criterion it fails, name the criterion, what is wrong, and why it matters for testing.

| Criterion | Fails when |
|---|---|
| **Complete** | Necessary information is absent: actor, object, constraint, data format, trigger or outcome |
| **Consistent** | It contradicts another requirement in the same document |
| **Correct** | It does not accurately reflect the client's need |
| **Feasible** | It cannot realistically be implemented or tested |
| **Necessary** | It is gold-plating or clearly out of scope |
| **Prioritised** | Relative importance matters but is completely absent |
| **Unambiguous** | More than one interpretation is possible |
| **Verifiable** | A pass/fail test cannot be defined without further clarification |


**Testability** – Rate whether a tester can write a pass/fail test from the requirement as written:
- **OK** – testable as written.
- **MANAGEABLE** – the basic success case is testable; limits, edge cases or error handling are unclear.
- **BLOCKING** – no pass/fail test is possible until the client clarifies it.

## Table columns

ID · Original text · Location · ISTQB issues · Testability · Rewritten requirement

Write `<document name> - Requirements.xlsx`, one sheet, one row per requirement, in document order:

| Column | Content |
|---|---|
| ID | `REQ-001`, `REQ-002`, … numbered in document order, no gaps |
| Original text | The source wording, verbatim |
| Location | `Section <number> <heading>, page <n>`; add `, table <n>, row <n>` for table cells |
| ISTQB issues | One line per failed criterion: `<Criterion>: <what is wrong> – <why it matters for testing>`; `None` if all pass |
| Testability | `OK`, `MANAGEABLE` or `BLOCKING` |
| Testability reason | What is missing; empty for `OK` |
| Rewritten requirement | One clear, testable sentence; unknown values written as `[TBD]` |
