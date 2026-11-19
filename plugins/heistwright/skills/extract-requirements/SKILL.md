---
name: extract-requirements
description: Turn a client document into a reviewed list of requirements in an Excel file, assessed for traceability, ISTQB quality and testability. Only use when called with `/extract-requirements`.
---

# Requirements from client documents

Requirements are rarely stated in one place: definitions, general constraints and exceptions sit elsewhere in the document. Read for those connections, otherwise requirements are rewritten wrongly and flagged for gaps the client already filled.

## Steps

1. **Collect context.** Read the whole document first. Note definitions and glossary entries, general or global constraints, cross-references ("see section 4"), exceptions, appendices and any priority scheme. Also note the document language (English or Norwegian), the page count, and every link (URL or hyperlink) with its location. Do not open or fetch any link while reading.
2. **Ask the user** in one prompt round:
   - **Output language:** English or Norwegian. Default to the document language.
   - **Links** (skip if there are none): state how many links were found and ask: *ignore all*, *try to fetch and read all myself*, or *decide per link*. For *decide per link*, ask one question per link (show the URL and its location) with the options *ignore*, *try to fetch and read it*, or *I will provide the information*.
3. **Give the time estimate.** Tell the user the expected run time before starting: about 0.25 minutes per page (40 pages ≈ 10 min, 60 pages ≈ 15 min), as a rough guide. Then post a short progress message after each of steps 5–7.
4. **Resolve links.** Fetch only links the user chose to fetch; for *I will provide the information*, ask the user for it. Treat the result as an extra source and cite it like any other (location = the URL). If a fetch fails, tell the user and ask them to provide the information or ignore the link.
5. **Extract.** Every paragraph, list item and table row that states an obligation becomes a requirement; split compound statements into one requirement each. When the same obligation is stated in several places (e.g. a paragraph and a table row), write one requirement and cite every passage as a `*` source; only merge when the meaning is identical, see [output-format.md](references/output-format.md). When a requirement depends on context from step 1 or on a link source, cite it as a `-` source.
6. **Assess.** For each requirement apply the three checks below and write a rewritten requirement.
7. **Compare.** Check all requirements against each other for contradictions and record them under *Consistent*.
8. **Write** the Excel file, then run the final check.

## The three checks

- **Traceability** – Quote the original wording verbatim and give its exact location (section, page, table and row), for every source the requirement depends on. See [output-format.md](references/output-format.md).
- **ISTQB quality** – Assess against the eight criteria. Search the rest of the document before flagging missing information. Read [istqb-criteria.md](references/istqb-criteria.md) before assessing.
- **Testability** – Rate `OK`, `MANAGEABLE` or `BLOCKING` (definitions in [istqb-criteria.md](references/istqb-criteria.md)).

## Output

Write `<document name> - Requirements.xlsx` with the columns, source format, language rule (including the translation table) and examples in [output-format.md](references/output-format.md). Read it before writing the file.

## Final check

Done when:
- every obligation in the document has a row, including obligations in tables, lists and appendices;
- no two rows state the same obligation;
- every qualifier found elsewhere in the document is cited;
- every contradiction names both requirement IDs;
- every link has a user decision and no link was opened without one;
- column headers and cell values (except IDs and original text) are in the language the user chose, using the translation table;
- IDs run without gaps in document order.
