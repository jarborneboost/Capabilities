---
name: extract-requirements
description: Turn a client document into a reviewed list of requirements in an Excel file, assessed for traceability, ISTQB quality and testability. Only use when called with `/extract-requirements`.
---

# Requirements from client documents

Requirements are rarely stated in one place: definitions, general constraints and exceptions sit elsewhere in the document. Read for those connections, otherwise requirements are rewritten wrongly and flagged for gaps the client already filled.

## Steps

1. **Collect context.** Read the whole document first. Note definitions and glossary entries, general or global constraints, cross-references ("see section 4"), exceptions, appendices and any priority scheme.
2. **Extract.** Every paragraph, list item and table row that states an obligation becomes a requirement; split compound statements into one requirement each. When a requirement depends on context from step 1, cite it as well.
3. **Assess.** For each requirement apply the three checks below and write a rewritten requirement.
4. **Compare.** Check all requirements against each other for contradictions and record them under *Consistent*.
5. **Write** the Excel file, then run the final check.

## The three checks

- **Traceability** – Quote the original wording verbatim and give its exact location (section, page, table and row), for every source the requirement depends on. See [output-format.md](references/output-format.md).
- **ISTQB quality** – Assess against the eight criteria. Search the rest of the document before flagging missing information. Read [istqb-criteria.md](references/istqb-criteria.md) before assessing.
- **Testability** – Rate `OK`, `MANAGEABLE` or `BLOCKING` (definitions in [istqb-criteria.md](references/istqb-criteria.md)).

## Output

Write `<document name> - Requirements.xlsx` with the columns, multi-source format, language rule and example in [output-format.md](references/output-format.md). Read it before writing the file.

## Final check

Done when:
- every obligation in the document has its own row, including obligations in tables, lists and appendices;
- every qualifier found elsewhere in the document is cited;
- every contradiction names both requirement IDs;
- IDs run without gaps in document order.
