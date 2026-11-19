# Output format

Write `<document name> - Requirements.xlsx`: one sheet, one row per requirement, in document order.

## Columns

Headers are written in the output language, see [Language](#language).

| Column | Content |
|---|---|
| ID | `REQ-001`, `REQ-002`, … numbered in document order of the first `*` source, no gaps |
| Original text | The source wording, verbatim. One line per source, see below |
| Location | `Section <number> <heading>, page <n>`; add `, table <n>, row <n>` for table cells. One line per source in Original text, same order |
| ISTQB issues | One line per failed criterion: `<Criterion>: <what is wrong> – <why it matters for testing>`; `None` if all pass |
| Rewritten requirement | One clear, testable sentence; unknown values written as `[TBD]` |
| Testability | `OK`, `MANAGEABLE` or `BLOCKING` |
| Testability reason | What is missing; empty for `OK` |

## Sources

A requirement has one row, however many places state or qualify it. Each source is one line in **Original text**, with its location on the matching line in **Location**:

- `* ` starts a passage that states the obligation. Several `*` lines mean the same obligation is stated in several places (for example a paragraph and a table row).
- `- ` starts a qualifier: a definition, general constraint, exception, appendix, or information from a link (fetched or provided by the user; the URL is its location).

List `*` lines first, then `-` lines, each group in document order. With a single source, write the quote without a marker. The rewritten requirement may only use information from the cited sources.

### Same obligation in several places

Merge passages into one row only when they mean exactly the same. Then:
- If one passage is more specific than the other (for example it adds a value), the rewritten requirement uses the specific one and both are cited with `*`.
- If the passages differ in value or condition, they are not duplicates: keep separate rows and record the contradiction (see *Consistency pass* in [istqb-criteria.md](istqb-criteria.md)).

## Language

Write everything in the output language the user chose (English or Norwegian), except the IDs, the `*` / `-` markers, and **Original text**, which stays verbatim. This covers the column headers, the Testability values, the ISTQB criterion names, `None`, `[TBD]`, and the words in Location.

| English | Norwegian |
|---|---|
| ID | ID |
| Original text | Opprinnelig tekst |
| Location | Plassering |
| ISTQB issues | ISTQB-avvik |
| Rewritten requirement | Omskrevet krav |
| Testability | Testbarhet |
| Testability reason | Begrunnelse for testbarhet |
| OK | OK |
| MANAGEABLE | HÅNDTERBAR |
| BLOCKING | BLOKKERENDE |
| None | Ingen |
| [TBD] | [TBD] |
| Section, page, table, row | Kapittel, side, tabell, rad |
| Complete | Fullstendig |
| Consistent | Konsistent |
| Correct | Korrekt |
| Feasible | Gjennomførbar |
| Necessary | Nødvendig |
| Prioritised | Prioritert |
| Unambiguous | Entydig |
| Verifiable | Etterprøvbar |

## Example rows

The rows below only illustrate the format. Their content is invented; do not copy their wording, numbers or findings. Each example is shown as a column/value table in English; in the Excel file each column is one cell and `<br>` stands for a line break inside the cell.

### Example 1: obligation with a qualifier, with an issue

| Column | Value |
|---|---|
| ID | REQ-007 |
| Original text | * "The system shall respond quickly to search requests."<br>- "Quickly means within 2 seconds for 95% of requests." |
| Location | * Section 3.2 Search, page 8<br>- Section 6.1 Definitions, page 21, table 4, row 3 |
| ISTQB issues | Complete: no load conditions given for the 2-second limit – testers cannot set up a representative test |
| Rewritten requirement | The system shall return search results within 2 seconds for 95% of search requests under [TBD] concurrent users. |
| Testability | MANAGEABLE |
| Testability reason | The limit is testable, but the load under which it applies is not stated. |

### Example 2: same obligation stated in a paragraph and a table

| Column | Value |
|---|---|
| ID | REQ-012 |
| Original text | * "Users must confirm their email address before first login."<br>* "Email confirmation: required before first login" |
| Location | * Section 3.1 Registration, page 7<br>* Section 5 Summary of rules, page 18, table 2, row 4 |
| ISTQB issues | None |
| Rewritten requirement | A new user shall confirm their email address before they can log in for the first time. |
| Testability | OK |
| Testability reason | (empty) |

### Example 3: single source, no issues

| Column | Value |
|---|---|
| ID | REQ-008 |
| Original text | "A user shall be able to reset their password by email." |
| Location | Section 3.4 Accounts, page 9 |
| ISTQB issues | None |
| Rewritten requirement | A registered user shall be able to reset their password through a link sent to their registered email address. |
| Testability | OK |
| Testability reason | (empty) |

