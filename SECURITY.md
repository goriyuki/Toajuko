# Security policy

## Reporting a problem

Please report suspected credential exposure or a security issue privately by emailing **yijunchen2003@126.com**. Do not open a public issue containing secrets, personal data, or employer information.

## Repository rules

- Never commit passwords, access tokens, database connection strings, private IP addresses, or `.env` files.
- Use environment variables and provide only a redacted `.env.example` when configuration examples are needed.
- Use synthetic or openly licensed data in reproducible examples.
- Before publishing, inspect both the working tree and Git history for sensitive content.

If a secret is committed, removing it from the latest file is not sufficient: rotate the secret immediately, then consider a coordinated history rewrite.
