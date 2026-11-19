# ISTQB criteria and testability levels

## ISTQB quality criteria

Assess every requirement against all eight criteria. For each criterion it fails, name the criterion, what is wrong, and why it matters for testing.

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

### Search the document before flagging missing information

Clients often state the missing piece elsewhere: a definition in a glossary, a general constraint in a non-functional section, an exception in an appendix. A flag raised on local text alone is a false finding that wastes the client's time.

Before you flag **Complete**, **Unambiguous** or **Verifiable**:
1. Search the whole document for the missing actor, value, definition, trigger or outcome.
2. If found, use it: cite it as an `Also:` source and do not flag the criterion.
3. If not found anywhere, flag it and say that the document does not state it.

### Consistency pass

Check **Consistent** only after all requirements are extracted. Compare requirements with each other, not just with their own section. Record each contradiction on both requirements, naming the other ID: `Consistent: contradicts REQ-014 (response time 2 s vs 5 s) – testers cannot know which limit to assert`.

## Testability

Rate whether a tester can write a pass/fail test from the requirement as written, including any `Also:` sources:

- **OK** – testable as written.
- **MANAGEABLE** – the basic success case is testable; limits, edge cases or error handling are unclear.
- **BLOCKING** – no pass/fail test is possible until the client clarifies it.
