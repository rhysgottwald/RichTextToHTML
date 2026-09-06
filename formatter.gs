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
        var text = run.getText();
        var style = run.getTextStyle();
        var linkUrl = run.getLinkUrl();

        // Apply all formatting
        if (style.isStrikethrough()) text = '<strike>' + text + '</strike>';
        if (style.isBold()) text = '<strong>' + text + '</strong>';
        if (style.isItalic()) text = '<em>' + text + '</em>';

        var color = style.getForegroundColor();
        var size = style.getFontSize();

        var spanStyles = '';
        if (color && color !== "#000000" && (!linkUrl || color !== "#1155cc")) spanStyles += 'color: ' + color + ';';
        if (size && size > 10) spanStyles += 'font-size:' + size + 'px;';
        if (spanStyles !== '') text = '<span style="' + spanStyles + '">' + text + '</span>\n';

        if (style.isUnderline() && !linkUrl) text = '<u>' + text + '</u>';
        if (linkUrl) text = '<a href="' + linkUrl + '">' + text + '</a>';

        styledText += text;

      });

      // Detect for lists
      styledText = processLists(styledText);

      // Finally detect and clean up line breaks
      styledText = styledText.replace(/(<\/li>|<\/ul>|<\/ol>)\n/g, "$1").replace(/\n/g, "<br />");
      styledText = styledText.replace(/<\/span><br \/>/g, '</span>');
    }

    // Set the value of the cell in the 'HTMLText' column to the styledText
    sheet.getRange(i + 2, htmlTextCol).setValue(styledText);
  }
}

function processLists(text) {
  var lines = text.split('\n');
  var stack = [];
  var output = [];

  for (var j = 0; j < lines.length; j++) {
    var line = lines[j];
    
    var isListItem = false;
    var markerContent = "";
    var markerType = "";

    if (line.startsWith('  - ')) {
      isListItem = true;
      markerContent = line.substring(4).trim();
      markerType = 'nested-dash';
    } else {
      var trimmed = line.trim();
      var matchNumber = trimmed.match(/^(\d+)\.\s+(.*)/);
      var matchLetter = trimmed.match(/^([a-z])\.\s+(.*)/);
      var matchNormalDash = trimmed.match(/^-\s+(.*)/);
      var matchStar = trimmed.match(/^\*\s+(.*)/);

      if (matchNumber) { isListItem = true; markerContent = matchNumber[2].trim(); markerType = 'number'; }
      else if (matchLetter) { isListItem = true; markerContent = matchLetter[2].trim(); markerType = 'letter'; }
      else if (matchNormalDash) { isListItem = true; markerContent = matchNormalDash[1].trim(); markerType = 'dash'; }
      else if (matchStar) { isListItem = true; markerContent = matchStar[1].trim(); markerType = 'star'; }
    }

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

    var initListStr = function(mType, parentType) {
      if (mType === 'nested-dash') {
        return parentType === 'ol' ? '<ol>' : '<ul>';
      }
      if (mType === 'letter') return '<ol type="a">';
      if (mType === 'number') return '<ol>';
      return '<ul>';
    };

    if (stack.length === 0) {
      var t = (markerType === 'number' || markerType === 'letter') ? 'ol' : 'ul';
      stack.push({ type: t, marker: markerType });
      output.push(initListStr(markerType, null) + liStr);
    } else {
      var top = stack[stack.length - 1];

      var isContinuing = (top.marker === markerType);
      var isStartingSublist = (markerType === 'nested-dash' && top.marker !== 'nested-dash');

      if (isContinuing) {
        output.push(liStr);
      } else if (isStartingSublist) {
        var parentType = top.type; // 'ol' or 'ul'
        stack.push({ type: parentType, marker: 'nested-dash' });
        output.push(initListStr('nested-dash', parentType) + liStr);
      } else {
        var matchedParentIndex = -1;
        for (var k = stack.length - 2; k >= 0; k--) {
          if (stack[k].marker === markerType) {
            matchedParentIndex = k;
            break;
          }
        }

        if (matchedParentIndex !== -1) {
          while (stack.length > matchedParentIndex + 1) {
             output[output.length - 1] += "</" + stack.pop().type + ">";
          }
          output.push(liStr);
        } else {
          while (stack.length > 0) {
             output[output.length - 1] += "</" + stack.pop().type + ">";
          }
          var t2 = (markerType === 'number' || markerType === 'letter') ? 'ol' : 'ul';
          stack.push({ type: t2, marker: markerType });
          output.push(initListStr(markerType, null) + liStr);
        }
      }
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
