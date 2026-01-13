function onOpen() {
    SpreadsheetApp.getUi()
      .createMenu('Markdown')
      .addItem('Convert to Formatted Text', 'convertMarkdown')
      .addToUi();
  }

  function convertMarkdown() {
    const sheet = SpreadsheetApp.getActiveSheet();
    const sourceCell = sheet.getActiveCell();
    const markdown = sourceCell.getValue();

    if (!markdown) {
      SpreadsheetApp.getUi().alert('선택한 셀에 마크다운 텍스트가 없습니다.');
      return;
    }

    const lines = markdown.split('\n');
    const startRow = sourceCell.getRow();
    const startCol = sourceCell.getColumn();

    sourceCell.clear();

    // 컬럼 너비 설정 (A=200, B~D=100)
    sheet.setColumnWidth(startCol, 200);
    sheet.setColumnWidth(startCol + 1, 100);
    sheet.setColumnWidth(startCol + 2, 100);
    sheet.setColumnWidth(startCol + 3, 100);

    let currentRow = startRow;
    let i = 0;

    while (i < lines.length) {
      const line = lines[i];

      if (line.trim() === '') {
        i++;
        continue;
      }

      if (line.trim().startsWith('|') && line.trim().endsWith('|')) {
        const tableResult = processTable(lines, i, sheet, currentRow, startCol);
        currentRow = tableResult.nextRow;
        i = tableResult.nextIndex;
        continue;
      }

      if (line.trim().match(/^-{3,}$/)) {
        currentRow++;
        i++;
        continue;
      }

      // H2 감지 - 앞에 빈 row 추가
      if (line.match(/^## /)) {
        currentRow++;
      }

      const cell = sheet.getRange(currentRow, startCol);
      processLine(line, cell);
      currentRow++;

      // H2 감지 - 뒤에 빈 row 추가
      if (line.match(/^## /)) {
        currentRow++;
      }

      i++;
    }
  }

  function processLine(line, cell) {
    let text = line;
    let fontSize = 10;
    let isHeading = false;

    const h1Match = text.match(/^# (.+)$/);
    const h2Match = text.match(/^## (.+)$/);
    const h3Match = text.match(/^### (.+)$/);

    if (h3Match) {
      text = h3Match[1];
      fontSize = 12;
      isHeading = true;
    } else if (h2Match) {
      text = h2Match[1];
      fontSize = 13;
      isHeading = true;
    } else if (h1Match) {
      text = h1Match[1];
      fontSize = 14;
      isHeading = true;
    }

    const boldRanges = [];
    const boldRegex = /\*\*(.+?)\*\*/g;
    let match;
    let offset = 0;

    while ((match = boldRegex.exec(text)) !== null) {
      const startInOriginal = match.index;
      const boldContent = match[1];
      const adjustedStart = startInOriginal - offset;

      boldRanges.push({
        start: adjustedStart,
        end: adjustedStart + boldContent.length,
        content: boldContent
      });

      offset += 4;
    }

    const cleanText = text.replace(/\*\*(.+?)\*\*/g, '$1');
    const builder = SpreadsheetApp.newRichTextValue().setText(cleanText);

    if (isHeading) {
      builder.setTextStyle(0, cleanText.length,
        SpreadsheetApp.newTextStyle().setBold(true).setFontSize(fontSize).build());
    } else {
      for (const range of boldRanges) {
        if (range.end <= cleanText.length) {
          builder.setTextStyle(range.start, range.end,
            SpreadsheetApp.newTextStyle().setBold(true).build());
        }
      }
    }

    cell.setRichTextValue(builder.build());
  }

  function processTable(lines, startIndex, sheet, startRow, startCol) {
    const tableLines = [];
    let i = startIndex;

    while (i < lines.length && lines[i].trim().startsWith('|') && lines[i].trim().endsWith('|')) {
      tableLines.push(lines[i]);
      i++;
    }

    if (tableLines.length === 0) {
      return { nextRow: startRow, nextIndex: startIndex + 1 };
    }

    let currentRow = startRow;

    for (let t = 0; t < tableLines.length; t++) {
      const tableLine = tableLines[t];

      if (tableLine.match(/^\|[\s\-:|]+\|$/)) {
        if (currentRow > startRow) {
          const cols = parseTableRow(tableLines[0]);
          for (let c = 0; c < cols.length; c++) {
            const cell = sheet.getRange(currentRow - 1, startCol + c);
            cell.setBorder(null, null, true, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID);
          }
        }
        continue;
      }

      const columns = parseTableRow(tableLine);
      const isHeader = (t === 0);

      for (let c = 0; c < columns.length; c++) {
        const cell = sheet.getRange(currentRow, startCol + c);
        const content = columns[c].trim();

        const boldRanges = [];
        const boldRegex = /\*\*(.+?)\*\*/g;
        let match;
        let offset = 0;

        while ((match = boldRegex.exec(content)) !== null) {
          const startInOriginal = match.index;
          const boldContent = match[1];
          const adjustedStart = startInOriginal - offset;

          boldRanges.push({
            start: adjustedStart,
            end: adjustedStart + boldContent.length
          });

          offset += 4;
        }

        const cleanContent = content.replace(/\*\*(.+?)\*\*/g, '$1');
        const builder = SpreadsheetApp.newRichTextValue().setText(cleanContent);

        if (isHeader) {
          builder.setTextStyle(0, cleanContent.length,
            SpreadsheetApp.newTextStyle().setBold(true).build());
          cell.setBackground('#f3f3f3');
        } else {
          for (const range of boldRanges) {
            if (range.end <= cleanContent.length) {
              builder.setTextStyle(range.start, range.end,
                SpreadsheetApp.newTextStyle().setBold(true).build());
            }
          }
        }

        cell.setRichTextValue(builder.build());
        cell.setBorder(true, true, true, true, null, null, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);
      }

      currentRow++;
    }

    currentRow++;

    return { nextRow: currentRow, nextIndex: i };
  }

  function parseTableRow(line) {
    const trimmed = line.trim();
    const withoutEdges = trimmed.slice(1, -1);
    return withoutEdges.split('|');
  }
