# Formula Similarity Matching

This internship prototype retrieves similar chemical formulas from a database using a transparent two-stage method:

1. **Weighted Jaccard screening** removes candidates with insufficient component overlap.
2. **Weighted Euclidean ranking** compares the proportions of shared and non-shared components.

The legacy executable is kept at [`加权Jaccard香精匹配算法.py`](../../加权Jaccard香精匹配算法.py) so existing links remain valid. Database credentials have been removed from source code.

## Configuration

Set these environment variables before running the script:

```text
FORMULA_DB_SERVER
FORMULA_DB_NAME
FORMULA_DB_USER
FORMULA_DB_PASSWORD
```

The script also expects local Excel inputs for the uploaded formula and exclusion list. Those files are intentionally ignored by Git and are not included in this public repository.

## Responsible publication

- Employer formulas, database contents, and connection details are not published.
- The repository demonstrates the matching approach rather than reproducing a production system.
- For production use, add parameterised SQL, structured logging, tests, schema validation, and a least-privilege database account.
