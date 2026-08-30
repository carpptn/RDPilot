# Security Policy

RDPilot controls the real Windows desktop and processes screenshots, UI context,
model responses, and local diagnostic artifacts. Please handle security reports
and shared artifacts carefully.

## Supported versions

| Version | Supported |
| --- | --- |
| Latest code on `master` | Yes |
| Older commits and tags | No |

## Reporting a vulnerability

Do not open a public issue for a suspected vulnerability. Report it privately
through [GitHub Security Advisories](https://github.com/carpptn/RDPilot/security/advisories/new).

Include a concise description, reproduction steps, affected configuration, the
potential impact, and any suggested mitigation. Remove API keys, credentials,
personal data, private documents, and unrelated screenshots from the report.

Please allow time to investigate and coordinate a fix before publicly disclosing
the issue.

## Operational safety

Unexpected model decisions, incorrect clicks, prompt injection from visible
content, and disclosure of information visible on screen are inherent risks of
desktop automation. They are important safety concerns but do not necessarily
indicate a software vulnerability. Follow the repository's
[safety guidance](README.md#safety-model), supervise runs, and use a disposable VM
or non-critical environment.
