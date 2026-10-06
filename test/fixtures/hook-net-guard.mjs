// Offline guard for the session-start hook: any network use fails the run.
globalThis.fetch = () => { throw new Error("hook network guard: fetch is disabled"); };
