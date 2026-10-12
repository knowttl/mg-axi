# Headless credential convention

The [AXI credential standard](https://github.com/knowttl/artium-axi/blob/main/docs/credential-standard.md) defines the shared credential contract.
A profile names an env var or a file; the process environment supplies the value.
The profile file stores only the reference, never a secret value.
Env var names match `/^[A-Za-z_][A-Za-z0-9_]*$/`.
Credential files use absolute local paths; Windows UNC and device paths are rejected.
Values are read at acquisition time, never at profile load, so rotation needs no profile change.
Key, secret and federated assertion files on POSIX must be regular files owned by the current user with no group or other permissions, never symlinks.
Mode 0400 is accepted as well as 0600.
Configuration must not be writable by group or other.
Windows uses a protected DACL granting FullControl only to the current user, SYSTEM and Administrators, with no inherited grants.
Every write applies that DACL; ordinary reads reject reparse points, and `auth status` and `doctor` verify the DACL.
A missing or empty holder fails closed with `AUTH_REQUIRED`; exposed files fail with `CONFIG_INSECURE` naming only the path and fix.
Secret values and key material never appear in config files, output, diagnostics, error messages or cache keys.
A client secret is weaker than a certificate or federated credential and needs scheduled rotation.
