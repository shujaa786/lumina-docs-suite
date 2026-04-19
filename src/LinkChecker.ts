

function checkDocumentLinks(): any {
  const doc = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const results: any[] = [];
  const urlsToCheck: string[] = [];

  const paragraphs = body.getParagraphs();
  for (var i = 0; i < paragraphs.length; i++) {
    const text = paragraphs[i].editAsText();
    const string = paragraphs[i].getText();

    for (var j = 0; j < string.length; j++) {
      const url = text.getLinkUrl(j);
      if (url && urlsToCheck.indexOf(url) === -1) {
        urlsToCheck.push(url);
      }
    }
  }

  if (urlsToCheck.length === 0) return "No links found.";

  const requests: any[] = urlsToCheck.map(function(url) {
    return {
      url: url,
      muteHttpExceptions: true,
      method: "get"
    };
  });

  try {
    const responses = UrlFetchApp.fetchAll(requests);

    for (var k = 0; k < responses.length; k++) {
      const code = responses[k].getResponseCode();
      results.push({
        url: urlsToCheck[k],
        status: code,
        isBroken: code >= 400
      });
    }

    return results;
  } catch (e: any) {
    return "Error checking links: " + e.toString();
  }
}



