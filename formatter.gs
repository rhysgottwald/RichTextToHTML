function escapeHtml(text) {
  return text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function copyFormattedText() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var richTextCol = headers.indexOf('RichText') + 1;
  var htmlTextCol = headers.indexOf('HTMLText') + 1;

  if (richTextCol === 0 || htmlTextCol === 0) {
    throw new Error('Could not find RichText and/or HTMLText column');
  }

  var richTextRange = sheet.getRange(2, richTextCol, sheet.getLastRow() - 1);
  var htmlTextRange = sheet.getRange(2, htmlTextCol, sheet.getLastRow() - 1);
  var richTextValues = richTextRange.getRichTextValues();
  var htmlTextValues = htmlTextRange.getValues();

  for (var i = 0; i < richTextValues.length; i++) {
    if (htmlTextValues[i][0] !== '') continue; // Skip processing if 'HTMLText' cell is not empty

    var richTextValue = richTextValues[i][0];
    var styledText = '';

    if (richTextValue) {
      var styleRuns = richTextValue.getRuns();

      styleRuns.forEach(function (run) {
        var text = escapeHtml(run.getText());
        var style = run.getTextStyle();

        // Apply all formatting
        if (style.isStrikethrough()) text = '<strike>' + text + '</strike>';
        if (style.isBold()) text = '<strong>' + text + '</strong>';
        if (style.isItalic()) text = '<em>' + text + '</em>';
        if (style.isUnderline()) text = '<u>' + text + '</u>';

        styledText += text;

      });

      // Collapse formatting tags that wrap only whitespace (artifact of run
      // splitting, e.g. a space between a list marker and its text getting
      // its own bold run) so marker regexes in processLists see real whitespace.
      var prevStyledText;
      do {
        prevStyledText = styledText;
        styledText = styledText.replace(/<(strong|em|strike|u)>(\s*)<\/\1>/g, '$2');
      } while (styledText !== prevStyledText);

      // Detect for lists
      styledText = processLists(styledText);

      // Finally detect and clean up line breaks
      styledText = styledText.replace(/(<\/li>|<\/ul>|<\/ol>)\n/g, "$1").replace(/\n/g, "<br />");
    }

    // Set the value of the cell in the 'HTMLText' column to the styledText
    sheet.getRange(i + 2, htmlTextCol).setValue(styledText);
  }
}

function processLists(text) {
  var lines = text.split('\n');
  var stack = []; // {type: 'ol'|'ul' (closing tag), marker: markerType, level: indent level}
  var output = [];

  var closingTypeFor = function(mType) {
    return (mType === 'number' || mType === 'letter') ? 'ol' : 'ul';
  };

  var openTagFor = function(mType) {
    if (mType === 'letter') return '<ol type="a">';
    if (mType === 'number') return '<ol>';
    return '<ul>';
  };

  for (var j = 0; j < lines.length; j++) {
    var line = lines[j];

    // Each 2-space indent is one nesting level.
    var indentSpaces = line.match(/^ */)[0].length;
    var level = Math.floor(indentSpaces / 2);
    var rest = line.slice(indentSpaces);

    // Marker may be wrapped in a leading formatting tag (e.g. whole line bolded);
    // strip it before matching the marker, then reattach around the content.
    // Only matches our own emitted tags (cell text is HTML-escaped) so literal
    // angle-bracket text typed by a user is never mistaken for a tag.
    var tagPrefixMatch = rest.match(/^(<(?:strong|em|strike|u)>)+/);
    var tagPrefix = tagPrefixMatch ? tagPrefixMatch[0] : '';
    var unprefixed = rest.slice(tagPrefix.length);

    var isListItem = false;
    var markerContent = "";
    var markerType = "";

    var matchNumber = unprefixed.match(/^(\d+)\.\s+(.*)/);
    var matchLetter = unprefixed.match(/^([a-z])\.\s+(.*)/);
    var matchDash = unprefixed.match(/^-\s+(.*)/);
    var matchStar = unprefixed.match(/^\*\s+(.*)/);

    if (matchNumber) { isListItem = true; markerContent = tagPrefix + matchNumber[2].trim(); markerType = 'number'; }
    else if (matchLetter) { isListItem = true; markerContent = tagPrefix + matchLetter[2].trim(); markerType = 'letter'; }
    else if (matchDash) { isListItem = true; markerContent = tagPrefix + matchDash[1].trim(); markerType = 'dash'; }
    else if (matchStar) { isListItem = true; markerContent = tagPrefix + matchStar[1].trim(); markerType = 'star'; }

    if (!isListItem) {
      while (stack.length > 0) {
        output[output.length - 1] += "</" + stack.pop().type + ">";
      }
      output.push(line);
      continue;
    }

    var stripped = markerContent.replace(/<[^>]+>/g, "").trim();
    if (stripped.length > 0 && !stripped.endsWith('.')) {
      var matchTags = markerContent.match(/(<\/[^>]+>)+$/);
      if (matchTags) {
        markerContent = markerContent.slice(0, -matchTags[0].length) + "." + matchTags[0];
      } else {
        markerContent += ".";
      }
    }

    var liStr = "<li>" + markerContent + "</li>";

    // Close any deeper, now-finished sublists.
    while (stack.length > 0 && stack[stack.length - 1].level > level) {
      output[output.length - 1] += "</" + stack.pop().type + ">";
    }

    var top = stack.length > 0 ? stack[stack.length - 1] : null;

    if (top && top.level === level && top.marker === markerType) {
      output.push(liStr);
    } else {
      if (top && top.level === level) {
        // Same indent, different marker type: close it and start fresh.
        output[output.length - 1] += "</" + stack.pop().type + ">";
      }
      stack.push({ type: closingTypeFor(markerType), marker: markerType, level: level });
      output.push(openTagFor(markerType) + liStr);
    }
  }

  while (stack.length > 0) {
    if (output.length > 0) {
      output[output.length - 1] += "</" + stack.pop().type + ">";
    } else {
      stack.pop();
    }
  }

  return output.join('\n');
}
