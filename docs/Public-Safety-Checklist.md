# Public Safety Checklist

Use this before pushing scripts, examples, reports, screenshots, or docs to a public repository.

## Never Commit

- Access tokens, refresh tokens, client secrets, certificates, private keys, or app passwords.
- Tenant IDs unless they are intentionally public demo values.
- Production domains, customer names, internal hostnames, internal IP addresses, or ticket numbers.
- Real user exports, mailbox exports, audit logs, transcripts, or incident evidence.
- Generated reports from customer or employer tenants.

## Safe Examples

- Use `example.com`, `example.net`, and placeholder GUIDs.
- Use fake names such as Alex Admin or Casey Lee.
- Keep CSV rows short and purpose-built.
- Make sample findings realistic without exposing real environments.

## Script Safety

- Use `-WhatIf` for scripts that create, update, remove, assign, or disable anything.
- Validate CSV headers before making changes.
- Write result logs for bulk changes.
- Prefer read-only exports for assessment workflows.
- Put required Graph scopes and Exchange permissions in each README.

## Local Output

Keep real output in ignored folders such as:

```text
private-output/
exports/
reports/
transcripts/
```

Review `git status --short` before every push.
