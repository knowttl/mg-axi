# Security reporting

Never post credentials, access tokens, private keys, cursor tokens, mutation journals or tenant data in a public issue.
Use synthetic identifiers and sanitized output when reproducing a problem.

GitHub private vulnerability reporting is not currently enabled for this repository.
For a suspected vulnerability, open a minimal [issue](https://github.com/knowttl/mg-axi/issues/new) requesting a private reporting channel, without exploit details or sensitive material.
Wait for a maintainer to arrange that channel before sharing the technical details.
For ordinary bugs, include the package version, command, expected behavior and a synthetic reproduction.

If a credential was exposed, revoke or rotate it through your organization's identity administration process.
Local `auth logout` deletes the selected profile's cache; it does not revoke upstream sessions or application credentials.
