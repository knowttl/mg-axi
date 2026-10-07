# Headless credential convention

This records the shared convention mg-axi application profiles follow, proposed for the other axi tools.
A profile names an env var or a file; the process environment supplies the value.
The profile file stores only the reference, never a secret value.
Env var names match `/^[A-Za-z_][A-Za-z0-9_]*$/`.
Relative file paths resolve against the working directory at acquisition time.
Values are read at acquisition time, never at profile load, so rotation needs no profile change.
Key and secret files on POSIX must be regular files owned by the current user with mode exactly 0600, never symlinks; Windows skips the ownership and mode checks.
A missing, empty or exposed holder fails closed with `AUTH_REQUIRED` naming only the reference.
Secret values and key material never appear in config files, output, diagnostics, error messages or cache keys.
A client secret is weaker than a certificate or federated credential and needs scheduled rotation.
