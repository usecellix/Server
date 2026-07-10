const SHEET_MENTION_PATTERN = /@\[([^\]]+)\]/g;

export function extractSheetMentions(message: string): string[] {
  const names: string[] = [];
  for (const match of message.matchAll(SHEET_MENTION_PATTERN)) {
    const name = match[1]?.trim();
    if (name && !names.includes(name)) {
      names.push(name);
    }
  }
  return names;
}

export function stripSheetMentions(message: string): string {
  return message.replace(SHEET_MENTION_PATTERN, '$1').replace(/\s{2,}/g, ' ').trim();
}
