# Output format

Write `<document name> - Requirements.xlsx`: one sheet, one row per requirement, in document order.

## Columns

| Column | Content |
|---|---|
| ID | `REQ-001`, `REQ-002`, … numbered in document order, no gaps |
| Original text | The source wording, verbatim. Add further sources when needed, see below |
| Location | `Section <number> <heading>, page <n>`; add `, table <n>, row <n>` for table cells. One entry per source in Original text, same order |
| ISTQB issues | One line per failed criterion: `<Criterion>: <what is wrong> – <why it matters for testing>`; `None` if all pass |
| Rewritten requirement | One clear, testable sentence; unknown values written as `[TBD]` |
| Testability | `OK`, `MANAGEABLE` or `BLOCKING` |
| Testability reason | What is missing; empty for `OK` |

## Requirements built from several places

When a requirement's meaning depends on text elsewhere (definition, general constraint, exception, appendix), list every source in **Original text**, one per line, starting with `Primary:` for the obligation itself and `Also:` for each qualifier. Put the matching locations in **Location**, one per line, same order. The rewritten requirement may only use information from the cited sources. Information from a link (fetched or provided by the user) is an `Also:` source; write the URL as its location.

## Language

Write ISTQB issues, Rewritten requirement and Testability reason in the output language the user chose (English or Norwegian). Original text stays verbatim. Column headers, IDs and the Testability values stay as defined above.

## Example rows

The rows below only illustrate the format. Their content is invented; do not copy their wording, numbers or findings. Each example is shown as a column/value table; in the Excel file each column is one cell and `<br>` stands for a line break inside the cell.

### Example 1: requirement built from two places, with an issue

| Column | Value |
|---|---|
| ID | REQ-007 |
| Original text | Primary: "The system shall respond quickly to search requests."<br>Also: "Quickly means within 2 seconds for 95% of requests." |
| Location | Section 3.2 Search, page 8<br>Section 6.1 Definitions, page 21, table 4, row 3 |
| ISTQB issues | Complete: no load conditions given for the 2-second limit – testers cannot set up a representative test |
| Rewritten requirement | The system shall return search results within 2 seconds for 95% of search requests under [TBD] concurrent users. |
| Testability | MANAGEABLE |
| Testability reason | The limit is testable, but the load under which it applies is not stated. |

### Example 2: requirement with no issues

| Column | Value |
|---|---|
| ID | REQ-008 |
| Original text | "A user shall be able to reset their password by email."<br>(single source, so no `Primary:` prefix) |
| Location | Section 3.4 Accounts, page 9 |
| ISTQB issues | None |
| Rewritten requirement | A registered user shall be able to reset their password through a link sent to their registered email address. |
| Testability | OK |
| Testability reason | (empty) |

For a single-source requirement, write the quote without a prefix; use `Primary:` and `Also:` only when there are several sources.
