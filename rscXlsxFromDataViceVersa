const JSZip = require('jszip');

export async function xlsxFromData(data, templateBase64 = null) {
  try {
    if (!templateBase64) {
      templateBase64 = "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";
    } else {
      const base64Prefix =
        "data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,";
      if (!templateBase64.startsWith(base64Prefix)) {
        throw new Error(
          "O templateBase64 fornecido não é um base64 de um arquivo .xlsx válido."
        );
      }
      templateBase64 = templateBase64.substring(base64Prefix.length);
    }
    const binary = Uint8Array.from(atob(templateBase64), (char) =>
      char.charCodeAt(0)
    );
    const zip = await JSZip.loadAsync(binary);
    const sheetPath = "xl/worksheets/sheet1.xml";
    let sheetXML = await zip.file(sheetPath).async("string");
    const sheetDataRegex = /<sheetData>(.*?)<\/sheetData>/s;
    const sheetDataMatch = sheetXML.match(sheetDataRegex);
    const newRows = [];
    let currentRowNumber = 1;
    if (sheetDataMatch) {
      const originalSheetData = sheetDataMatch[1];
      const lastRowRegex = /<row r="(\d+)">/g;
      let lastRowMatch;
      let lastRowNumber = 0;
      while ((lastRowMatch = lastRowRegex.exec(originalSheetData)) !== null) {
        lastRowNumber = Math.max(lastRowNumber, parseInt(lastRowMatch[1]));
      }
      currentRowNumber = lastRowNumber + 1;
      newRows.push(originalSheetData);
    }
    const chunkSize = 1000;

    function colIndexToExcelCol(colIndex) {
      let dividend = colIndex + 1;
      let columnName = "";
      let modulo;

      while (dividend > 0) {
        modulo = (dividend - 1) % 26;
        columnName = String.fromCharCode(65 + modulo) + columnName;
        dividend = parseInt((dividend - modulo) / 26);
      }

      return columnName;
    }

    for (let i = 0; i < data.length; i += chunkSize) {
      const chunk = data.slice(i, i + chunkSize);
      const chunkRows = [];
      chunk.forEach((row) => {
        let newRow = `<row r="${currentRowNumber}">`;
        row.forEach((cellData, colIndex) => {
          const cellRef = `${colIndexToExcelCol(colIndex)}${currentRowNumber}`;
          newRow += `<c r="${cellRef}" t="str"><v>${cellData}</v></c>`;
        });
        newRow += `</row>`;
        chunkRows.push(newRow);
        currentRowNumber++;
      });
      const chunkRowsString = chunkRows.join("");
      const sheetDataEndRegex = /<\/sheetData>/;
      const sheetDataEndMatch = sheetXML.match(sheetDataEndRegex);
      if (sheetDataEndMatch) {
        const insertPosition = sheetDataEndMatch.index;
        const updatedXMLChunk =
          sheetXML.substring(0, insertPosition) +
          chunkRowsString +
          sheetXML.substring(insertPosition);

        zip.file(sheetPath, updatedXMLChunk);
        sheetXML = await zip.file(sheetPath).async("string");
      } else {
        throw new Error("Tag </sheetData> não encontrada no XML.");
      }
    }
    const base64String = await zip.generateAsync({ type: "base64" });
    const dataUriString = `data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,${base64String}`;
    return dataUriString;
  } catch (error) {
    console.error("Erro ao gerar o Excel Base64:", error);
    throw error;
  }
}

export async function dataFromXlsx(xlsxBase64, options = {}) {
  const defaultOptions = {
    hourColumns: [],
    dateColumns: [],
    dateHourColumns: [],
  };

  options = { ...defaultOptions, ...options };

  function columnLetterToIndex(letter) {
    let index = 0;
    for (let i = 0; i < letter.length; i++) {
      index = index * 26 + (letter.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
    }
    return index - 1;
  }

  function indexToColumnLetter(index) {
    let columnLetter = '';
    while (index >= 0) {
      columnLetter = String.fromCharCode(65 + (index % 26)) + columnLetter;
      index = Math.floor(index / 26) - 1;
    }
    return columnLetter;
  }

  function excelDateToJSDate(excelDate) {
    if (typeof excelDate === 'number') {
      return new Date(Math.round((excelDate - 25569) * 86400 * 1000));
    } else if (!isNaN(parseInt(excelDate))) {
      return new Date(Math.round((parseInt(excelDate) - 25569) * 86400 * 1000));
    }
    return null;
  }

  function formatDate(date) {
    if (!date) return '';
    const d = date.getDate().toString().padStart(2, '0');
    const m = (date.getMonth() + 1).toString().padStart(2, '0');
    const y = date.getFullYear();
    return `${d}/${m}/${y}`;
  }

  function formatDateTime(date) {
    if (!date) return '';
    const d = date.getDate().toString().padStart(2, '0');
    const m = (date.getMonth() + 1).toString().padStart(2, '0');
    const y = date.getFullYear();
    const h = date.getHours().toString().padStart(2, '0');
    const min = date.getMinutes().toString().padStart(2, '0');
    return `${d}/${m}/${y} ${h}:${min}`;
  }

  try {
    const base64Prefix =
      'data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,';
    if (!xlsxBase64.startsWith(base64Prefix)) {
      throw new Error(
        'O xlsxBase64 fornecido não é um base64 de um arquivo .xlsx válido.'
      );
    }
    xlsxBase64 = xlsxBase64.substring(base64Prefix.length);

    const binary = Uint8Array.from(atob(xlsxBase64), (char) =>
      char.charCodeAt(0)
    );
    const zip = await JSZip.loadAsync(binary);
    const sharedStringsPath = 'xl/sharedStrings.xml';
    const hasSharedStrings = zip.file(sharedStringsPath) !== null;
    let sharedStrings = [];

    if (hasSharedStrings) {
      const sharedStringsXML = await zip
        .file(sharedStringsPath)
        .async('string');
      sharedStrings =
        sharedStringsXML
          .match(/<t[^>]*>(.*?)<\/t>/gs)
          ?.map((str) => str.replace(/<t[^>]*>(.*?)<\/t>/, '$1')) || [];
    }
    const sheetPath = 'xl/worksheets/sheet1.xml';
    const sheetXML = await zip.file(sheetPath).async('string');
    const rows = sheetXML.match(/<row[^>]*>(.*?)<\/row>/gs) || [];
    let data = [];
    let maxCols = 0;
    rows.forEach((row) => {
      const cells =
        row.match(/<c r="([A-Z]+[0-9]+)"[^>]*>(.*?)<\/c>/gs) || [];
      let rowData = [];
      let colIndex = 0;
      cells.forEach((cell) => {
        const cellContent = cell;
        const colRef = cellContent.match(/r="([A-Z]+)[0-9]+"/);
        let currentColumnIndex = 0;
        if (colRef && colRef[1]) {
          currentColumnIndex = columnLetterToIndex(colRef[1]);
        }
        maxCols = Math.max(maxCols, currentColumnIndex + 1);
        const valueMatch = cellContent.match(/<v>(.*?)<\/v>/);
        const typeMatch = cellContent.match(/t="([^"]*)"/);
        let cellValue = '';
        if (
          hasSharedStrings &&
          typeMatch &&
          typeMatch[1] === 's' &&
          valueMatch
        ) {
          const cellValueIndex = parseInt(valueMatch[1], 10);
          cellValue = sharedStrings[cellValueIndex] || '';
        } else if (valueMatch) {
          cellValue = valueMatch[1];
        }
        while (colIndex < currentColumnIndex) {
          rowData.push('');
          colIndex++;
        }
        rowData.push(cellValue);
        colIndex++;
      });
      while (colIndex < maxCols) {
        rowData.push('');
        colIndex++;
      }
      data.push(rowData);
    });

    data = data.filter((row) => row.some((cell) => cell.trim() !== ''));
    maxCols = 0;
    data.forEach((row) => {
      maxCols = Math.max(maxCols, row.length);
    });

    if (data.length > 0) {
      let colsToRemove = [];
      for (let j = 0; j < maxCols; j++) {
        let columnIsEmpty = true;
        for (let i = 0; i < data.length; i++) {
          if (data[i][j] && data[i][j].trim() !== '') {
            columnIsEmpty = false;
            break;
          }
        }
        if (columnIsEmpty) {
          colsToRemove.push(j);
        }
      }
      for (let j = colsToRemove.length - 1; j >= 0; j--) {
        data.forEach((row) => row.splice(colsToRemove[j], 1));
      }
    }

    if (data.length > 0) {
      for (let i = 1; i < data.length; i++) {
        for (let j = 0; j < data[i].length; j++) {
          const columnLetter = indexToColumnLetter(j);
          const header = data[0] ? data[0][j] : null;

          if (
            options.dateColumns.includes(j + 1) ||
            options.dateColumns.includes(columnLetter) ||
            options.dateColumns.includes(header)
          ) {
            const jsDate = excelDateToJSDate(data[i][j]);
            data[i][j] = formatDate(jsDate);
          } else if (
            options.hourColumns.includes(j + 1) ||
            options.hourColumns.includes(columnLetter) ||
            options.hourColumns.includes(header)
          ) {
            if (!isNaN(parseFloat(data[i][j]))) {
              const decimalTime = parseFloat(data[i][j]);
              const hours = Math.floor(decimalTime * 24);
              const minutes = Math.floor(((decimalTime * 24) % 1) * 60);
              data[i][j] = `${hours.toString().padStart(2, '0')}:${minutes.toString().padStart(2, '0')}`;
            }
          } else if (
            options.dateHourColumns.includes(j + 1) ||
            options.dateHourColumns.includes(columnLetter) ||
            options.dateHourColumns.includes(header)
          ) {
            const jsDate = excelDateToJSDate(data[i][j]);
            data[i][j] = formatDateTime(jsDate);
          }
        }
      }
    }

    return data;
  } catch (error) {
    console.error('Erro ao extrair dados do Excel Base64:', error);
    throw error;
  }
}
