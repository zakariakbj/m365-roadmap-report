# Security policy

## Reporting a vulnerability

Please do not open a public issue for a security problem. Report it privately through
[GitHub's private vulnerability reporting](https://github.com/zakariakbj/m365-roadmap-report/security/advisories/new)
instead, with the steps to reproduce it and the version you use.

You should get an answer within a week. Once the problem is fixed, the advisory is published
with credit to you, unless you prefer to stay anonymous.

## Supported versions

Only the latest version on the `main` branch gets security fixes.

## Your configuration

`env.ini` holds the client secret of your Entra ID application: keep it out of version control
(it is git-ignored) and readable only by the account that runs the script. If a secret leaks,
delete it in the app registration and create a new one.
