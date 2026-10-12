import { execFileSync } from "node:child_process";
import { randomUUID } from "node:crypto";
import { chmodSync, closeSync, constants, existsSync, fstatSync, fsyncSync, lstatSync, mkdirSync, openSync, readFileSync, renameSync, rmSync, writeFileSync } from "node:fs";
import { dirname, isAbsolute, win32 } from "node:path";
import { AxiError } from "axi-sdk-js";

type AclAction = "restrict" | "verify" | "inspect";
export type WindowsAcl = (path: string, directory: boolean, action: AclAction) => void;

// LiteralPath and single-quoted literals keep file names out of PowerShell syntax.
export const windowsAcl: WindowsAcl = (path, directory, action) => {
  const literal = `'${path.replace(/'/g, "''")}'`;
  const script = `$ErrorActionPreference='Stop'; $p=${literal};
    $item=Get-Item -LiteralPath $p -Force;
    if (($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0) { throw 'Reparse point' };
    ${action === "inspect" ? "" : `
    $user=[Security.Principal.WindowsIdentity]::GetCurrent().User;
    $ids=@($user.Value,'S-1-5-18','S-1-5-32-544');
    ${action === "restrict" ? `
    $acl=[Security.AccessControl.${directory ? "DirectorySecurity" : "FileSecurity"}]::new();
    $acl.SetOwner($user); $acl.SetAccessRuleProtection($true,$false);
    foreach ($id in $ids) {
      $sid=[Security.Principal.SecurityIdentifier]::new($id);
      $rule=[Security.AccessControl.FileSystemAccessRule]::new($sid,'FullControl','${directory ? "ContainerInherit,ObjectInherit" : "None"}','None','Allow');
      $acl.AddAccessRule($rule);
    }; Set-Acl -LiteralPath $p -AclObject $acl;` : `
    $acl=Get-Acl -LiteralPath $p;
    if (!$acl.AreAccessRulesProtected -or $acl.GetOwner([Security.Principal.SecurityIdentifier]).Value -ne $user.Value) { throw 'Unprotected ACL' };
    $rules=@($acl.GetAccessRules($true,$true,[Security.Principal.SecurityIdentifier]));
    if ($rules.Count -ne 3) { throw 'Unexpected ACL' };
    foreach ($id in $ids) {
      $matches=@($rules | Where-Object { $_.IdentityReference.Value -eq $id -and $_.AccessControlType -eq 'Allow' -and $_.FileSystemRights -eq 'FullControl' -and !$_.IsInherited -and $_.InheritanceFlags -eq '${directory ? "ContainerInherit, ObjectInherit" : "None"}' -and $_.PropagationFlags -eq 'None' });
      if ($matches.Count -ne 1) { throw 'Unexpected grant' };
    };`}`}`;
  execFileSync("powershell.exe", ["-NoProfile", "-NonInteractive", "-EncodedCommand", Buffer.from(script, "utf16le").toString("base64")], { stdio: "pipe", windowsHide: true });
};

// All private-file permissions live here; Windows branches are injectable on Linux.
export class PrivateFiles {
  constructor(private platform = process.platform, private acl: WindowsAcl = windowsAcl, private uid = process.getuid?.()) {}

  absolute(path: string): boolean {
    return this.platform === "win32" ? /^[A-Za-z]:[\\/]/.test(path) && win32.isAbsolute(path) : isAbsolute(path);
  }

  private insecure(path: string): never {
    throw new AxiError(`Insecure file ${path}`, "CONFIG_INSECURE", [this.platform === "win32" ? `Re-run the tool's write to protect ${path}` : `chmod 600 '${path.replace(/'/g, "'\\''")}'`]);
  }

  check(path: string, secret = true, verifyAcl = false) {
    try {
      const stat = lstatSync(path);
      if (!stat.isFile() || stat.isSymbolicLink()) this.insecure(path);
      if (this.platform === "win32") {
        try { this.acl(path, false, verifyAcl ? "verify" : "inspect"); }
        catch (error) {
          if (error instanceof AxiError) throw error;
          this.insecure(path);
        }
      }
      else if (stat.uid !== this.uid || (stat.mode & (secret ? 0o077 : 0o022)) !== 0) this.insecure(path);
      return stat;
    } catch (error) {
      const code = (error as NodeJS.ErrnoException).code;
      if (code === "ENOENT") throw error;
      if (code === "EACCES" || code === "EPERM" || code === "ELOOP") return this.insecure(path);
      throw error;
    }
  }

  read(path: string, secret = true, verifyAcl = false): string {
    const checked = this.check(path, secret, verifyAcl);
    let fd: number | undefined;
    try {
      fd = openSync(path, constants.O_RDONLY | (this.platform === "win32" ? 0 : constants.O_NOFOLLOW));
      const opened = fstatSync(fd);
      if (opened.dev !== checked.dev || opened.ino !== checked.ino) this.insecure(path);
      return readFileSync(fd, "utf8").replace(/^\uFEFF/, "");
    } catch (error) {
      if (error instanceof AxiError) throw error;
      const code = (error as NodeJS.ErrnoException).code;
      if (code === "EACCES" || code === "EPERM" || code === "ELOOP") this.insecure(path);
      throw error;
    }
    finally { if (fd !== undefined) try { closeSync(fd); } catch { /* original error preserved */ } }
  }

  directory(path: string): string | undefined {
    try {
      let firstCreated: string | undefined;
      if (!existsSync(path)) {
        const parent = dirname(path);
        if (!existsSync(parent)) firstCreated = this.directory(parent);
        mkdirSync(path, { mode: 0o700 });
        firstCreated ??= path;
      }
      const stat = lstatSync(path);
      if (!stat.isDirectory() || stat.isSymbolicLink()) this.insecure(path);
      if (this.platform === "win32") {
        try { this.acl(path, true, "restrict"); }
        catch (error) {
          if (error instanceof AxiError) throw error;
          this.insecure(path);
        }
      } else {
        if (stat.uid !== this.uid) this.insecure(path);
        chmodSync(path, 0o700);
      }
      return firstCreated;
    } catch (error) {
      if (error instanceof AxiError) throw error;
      const code = (error as NodeJS.ErrnoException).code;
      if (code === "EACCES" || code === "EPERM") this.insecure(path);
      throw error;
    }
  }

  protect(path: string) {
    try {
      const stat = lstatSync(path);
      if (!stat.isFile() || stat.isSymbolicLink()) this.insecure(path);
      if (this.platform === "win32") {
        try { this.acl(path, false, "restrict"); }
        catch (error) {
          if (error instanceof AxiError) throw error;
          this.insecure(path);
        }
      } else {
        if (stat.uid !== this.uid) this.insecure(path);
        chmodSync(path, 0o600);
      }
    } catch (error) {
      if (error instanceof AxiError) throw error;
      const code = (error as NodeJS.ErrnoException).code;
      if (code === "EACCES" || code === "EPERM") this.insecure(path);
      throw error;
    }
  }

  write(path: string, content: string) {
    this.directory(dirname(path));
    const temporary = `${path}.${randomUUID()}.tmp`;
    let fd: number | undefined;
    try {
      try {
        fd = openSync(temporary, "wx", this.platform === "win32" ? undefined : 0o600);
        try {
          if (this.platform === "win32") this.acl(temporary, false, "restrict");
        } catch (error) {
          if (error instanceof AxiError) throw error;
          this.insecure(path);
        }
        writeFileSync(fd, content);
        fsyncSync(fd);
        closeSync(fd);
        fd = undefined;
      } catch (error) {
        if (error instanceof AxiError) throw error;
        const code = (error as NodeJS.ErrnoException).code;
        if (code === "EACCES" || code === "EPERM") this.insecure(path);
        throw error;
      }
      const deadline = Date.now() + 1000;
      let delay = 10;
      for (;;) {
        try { renameSync(temporary, path); break; }
        catch (error) {
          if (this.platform !== "win32" || !["EPERM", "EBUSY", "EACCES"].includes((error as NodeJS.ErrnoException).code ?? "") || Date.now() >= deadline) throw error;
          Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0, Math.min(delay, deadline - Date.now()));
          delay = Math.min(delay * 2, 100);
        }
      }
    } finally {
      if (fd !== undefined) try { closeSync(fd); } catch { /* original error preserved */ }
      rmSync(temporary, { force: true });
    }
  }
}
