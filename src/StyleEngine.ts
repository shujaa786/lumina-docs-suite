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
  let appliedCount = 0;

  while (found) {
    const textElement = found.getElement().asText();
    const start = found.getStartOffset();
    const end = found.getEndOffsetInclusive();

    snapshot.push({
      path: getElementPath(textElement),
      start,
      end,
      bold: textElement.isBold(start) === true,
      italic: textElement.isItalic(start) === true,
      backgroundColor: textElement.getBackgroundColor(start),
      foregroundColor: textElement.getForegroundColor(start),
    });

    if (config.bold) {
      textElement.setBold(start, end, true);
      appliedCount++;
    }
    if (config.italic) {
      textElement.setItalic(start, end, true);
      appliedCount++;
    }
    if (config.highlight) {
      Logger.log('Applying highlight color: ' + config.highlight + ' to text from ' + start + ' to ' + end);
      textElement.setBackgroundColor(start, end, config.highlight);
      appliedCount++;
    }
    if (config.textColor) {
      textElement.setForegroundColor(start, end, config.textColor);
      appliedCount++;
    }

    found = body.findText(keyword, found);
  }

  if (snapshot.length > 0) {
    PropertiesService.getUserProperties().setProperty(UNDO_PROP_KEY, JSON.stringify({ snapshot }));
  }
  
  Logger.log('applyBatchStyle: Applied ' + appliedCount + ' styles to ' + snapshot.length + ' matches');
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

function highlightMatch(keyword: string, index: number): void {
  const body = DocumentApp.getActiveDocument().getBody();
  let found = body.findText(keyword);
  let currentIndex = 0;

  while (found) {
    if (currentIndex === index) {
      // Build a proper Range object and set selection
      const element = found.getElement();
      const startOffset = found.getStartOffset();
      const endOffset = found.getEndOffsetInclusive();
      
      const range = DocumentApp.getActiveDocument().newRange()
        .addElement(element, startOffset, endOffset)
        .build();
      
      DocumentApp.getActiveDocument().setSelection(range);
      return;
    }
    currentIndex += 1;
    found = body.findText(keyword, found);
  }
}

function formatMatchAtIndex(keyword: string, index: number, config: any): void {
  const body = DocumentApp.getActiveDocument().getBody();
  let found = body.findText(keyword);
  let currentIndex = 0;
  const snapshot: any[] = [];

  Logger.log('formatMatchAtIndex: Looking for match at index ' + index + ' with config: ' + JSON.stringify(config));

  while (found) {
    if (currentIndex === index) {
      const textElement = found.getElement().asText();
      const start = found.getStartOffset();
      const end = found.getEndOffsetInclusive();

      Logger.log('Found match at index ' + currentIndex + ' from ' + start + ' to ' + end);

      snapshot.push({
        path: getElementPath(textElement),
        start,
        end,
        bold: textElement.isBold(start) === true,
        italic: textElement.isItalic(start) === true,
        backgroundColor: textElement.getBackgroundColor(start),
        foregroundColor: textElement.getForegroundColor(start),
      });

      if (config.bold) {
        textElement.setBold(start, end, true);
        Logger.log('Applied bold');
      }
      if (config.italic) {
        textElement.setItalic(start, end, true);
        Logger.log('Applied italic');
      }
      if (config.highlight) {
        Logger.log('Applying highlight color: ' + config.highlight + ' to text from ' + start + ' to ' + end);
        textElement.setBackgroundColor(start, end, config.highlight);
        Logger.log('Highlight applied');
      }
      if (config.textColor) {
        textElement.setForegroundColor(start, end, config.textColor);
        Logger.log('Applied text color');
      }

      if (snapshot.length > 0) {
        PropertiesService.getUserProperties().setProperty(UNDO_PROP_KEY, JSON.stringify({ snapshot }));
      }
      return;
    }
    currentIndex += 1;
    found = body.findText(keyword, found);
  }
  
  Logger.log('formatMatchAtIndex: No match found at index ' + index);
}

function replaceCurrentMatch(keyword: string, replacement: string, index: number): void {
  const body = DocumentApp.getActiveDocument().getBody();
  let found = body.findText(keyword);
  let currentIndex = 0;

  while (found) {
    if (currentIndex === index) {
      const element = found.getElement();
      const startOffset = found.getStartOffset();
      const endOffset = found.getEndOffsetInclusive();
      
      if (element.getType() === DocumentApp.ElementType.TEXT) {
        const textElement = element.asText();
        // Replace the text
        textElement.deleteText(startOffset, endOffset);
        textElement.insertText(startOffset, replacement);
        
        // Select the replaced text
        const range = DocumentApp.getActiveDocument().newRange()
          .addElement(element, startOffset, startOffset + replacement.length - 1)
          .build();
        DocumentApp.getActiveDocument().setSelection(range);
      }
      return;
    }
    currentIndex += 1;
    found = body.findText(keyword, found);
  }
}

function replaceAllMatches(keyword: string, replacement: string): number {
  const body = DocumentApp.getActiveDocument().getBody();
  let found = body.findText(keyword);
  let replacedCount = 0;

  while (found) {
    const element = found.getElement();
    const startOffset = found.getStartOffset();
    const endOffset = found.getEndOffsetInclusive();
    
    if (element.getType() === DocumentApp.ElementType.TEXT) {
      const textElement = element.asText();
      textElement.deleteText(startOffset, endOffset);
      textElement.insertText(startOffset, replacement);
      replacedCount += 1;
    }
    
    // Find next match after replacement
    found = body.findText(keyword, element);
  }

  return replacedCount;
}
