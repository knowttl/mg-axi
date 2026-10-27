export function mailText(full: boolean) {
  let truncated = false;
  return {
    text(value: string): string {
      if (full || value.length <= 500) return value;
      truncated = true;
      return `${value.slice(0, 500)}... (truncated, ${value.length} chars total)`;
    },
    help(): string[] {
      return truncated ? ["Repeat the same command with --full for complete text; redaction remains"] : [];
    },
  };
}
