# Security policy

## Supported versions

The latest `0.1.x` development line is the only version currently receiving
security fixes.

## Reporting a vulnerability

Please do not publish credentials, OAuth tokens, document contents, or a
working exploit in a public issue. Once this repository has GitHub Security
Advisories enabled, use a private advisory. Until then, contact the project
maintainer privately through the contact method published with the repository.

Include the affected version, operating system, reproduction steps, and the
smallest safe proof of impact. Redact API keys, refresh tokens, documents, and
personal data from reports.

## Security boundaries

OpenWordCode is designed to run its Core on loopback. Do not expose ports
`10200` or `10101` to a network interface. Keep OAuth client secrets in local
environment configuration and never commit `.env`, credential stores, browser
profiles, or Office development certificates.

The console tool is deliberately bounded and read-only in this development
line. Word writes go through the approval and stale-target checks. Treat
`Skip all approvals` as unsafe for untrusted prompts or documents.
