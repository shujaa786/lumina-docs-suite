function applyBatchStyle(keyword: string, config: any): void {
  const body = DocumentApp.getActiveDocument().getBody();
  let found = body.findText(keyword);

  while (found) {
    const textElement = found.getElement().asText();
    const start = found.getStartOffset();
    const end = found.getEndOffsetInclusive();

    if (config.bold) textElement.setBold(start, end, true);
    if (config.italic) textElement.setItalic(start, end, true);
    if (config.highlight) textElement.setBackgroundColor(start, end, config.highlight);
    if (config.textColor) textElement.setForegroundColor(start, end, config.textColor);

    found = body.findText(keyword, found);
  }
}
