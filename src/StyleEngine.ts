const UNDO_PROP_KEY = 'LUMINA_LAST_STYLE_SNAPSHOT';

function getElementPath(element: GoogleAppsScript.Document.Element): number[] {
  const path: number[] = [];
  let current: GoogleAppsScript.Document.Element | null = element;

  while (current && current.getParent() && current.getParent().getType() !== DocumentApp.ElementType.DOCUMENT) {
    const parent = current.getParent();
    path.unshift(parent.getChildIndex(current));
    current = parent;
  }

  return path;
}

function getElementByPath(path: number[]): GoogleAppsScript.Document.Element | null {
  let element: GoogleAppsScript.Document.Element = DocumentApp.getActiveDocument().getBody();

  for (let i = 0; i < path.length; i++) {
    if (!element || !('getChild' in element)) {
      return null;
    }
    element = element.getChild(path[i]);
  }

  return element;
}

function countMatches(keyword: string): number {
  const body = DocumentApp.getActiveDocument().getBody();
  let count = 0;
  let found = body.findText(keyword);

  while (found) {
    count += 1;
    found = body.findText(keyword, found);
  }

  return count;
}

function applyBatchStyle(keyword: string, config: any): void {
  const body = DocumentApp.getActiveDocument().getBody();
  let found = body.findText(keyword);
  const snapshot: any[] = [];

  while (found) {
    const textElement = found.getElement().asText();
    const start = found.getStartOffset();
    const end = found.getEndOffsetInclusive();

    snapshot.push({
      path: getElementPath(textElement),
      start,
      end,
      bold: textElement.isBold(start, end) === true,
      italic: textElement.isItalic(start, end) === true,
      backgroundColor: textElement.getBackgroundColor(start, end),
      foregroundColor: textElement.getForegroundColor(start, end),
    });

    if (config.bold) textElement.setBold(start, end, true);
    if (config.italic) textElement.setItalic(start, end, true);
    if (config.highlight) textElement.setBackgroundColor(start, end, config.highlight);
    if (config.textColor) textElement.setForegroundColor(start, end, config.textColor);

    found = body.findText(keyword, found);
  }

  if (snapshot.length > 0) {
    PropertiesService.getUserProperties().setProperty(UNDO_PROP_KEY, JSON.stringify({ snapshot }));
  }
}

function undoLastBatchStyle(): string {
  const saved = PropertiesService.getUserProperties().getProperty(UNDO_PROP_KEY);
  if (!saved) {
    throw new Error('No undo data available. Please run a format action first.');
  }

  const data = JSON.parse(saved);
  let restored = 0;

  for (const item of data.snapshot) {
    const element = getElementByPath(item.path);
    if (!element || element.getType() !== DocumentApp.ElementType.TEXT) {
      continue;
    }

    const textElement = element.asText();
    textElement.setBold(item.start, item.end, item.bold === true);
    textElement.setItalic(item.start, item.end, item.italic === true);
    textElement.setBackgroundColor(item.start, item.end, item.backgroundColor || null);
    textElement.setForegroundColor(item.start, item.end, item.foregroundColor || null);
    restored += 1;
  }

  PropertiesService.getUserProperties().deleteProperty(UNDO_PROP_KEY);
  return `Undo complete: restored ${restored} range${restored === 1 ? '' : 's'}.`;
}
