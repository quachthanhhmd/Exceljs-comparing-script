const ExcelJS = require("exceljs");
const path = require('path');

const {
  compareConditionalFormattingRules,
} = require("./conditional-formatting-rules");
const { getStyleDifferences } = require("./compare-style-different");
const { exportToExcel } = require("./export-xlsx-file");
const { getReadableCellStyle } = require("./readable-style");

const coalesce = (a) => {
  return a ? a : "";
};

function getCellLocation(colNumber, rowNumber) {
  let colLetter = "";
  while (colNumber > 0) {
    let remainder = (colNumber - 1) % 26;
    colLetter = String.fromCharCode(65 + remainder) + colLetter;
    colNumber = Math.floor((colNumber - 1) / 26);
  }
  return `${colLetter}${rowNumber}`;
}

async function compareExcelFiles(file1, file2) {
  const workbook1 = new ExcelJS.Workbook();
  const workbook2 = new ExcelJS.Workbook();

  await workbook1.xlsx.readFile(file1);
  await workbook2.xlsx.readFile(file2);

  const differences = [];

  workbook1.worksheets.forEach((sheet1, index) => {
    const sheet2 = workbook2.worksheets[index];
    if (!sheet2) {
      differences.push(`Sheet ${sheet1.name} is missing in second file.`);
      return;
    }

    const rowCount = Math.max(sheet1.rowCount, sheet2.rowCount);
    const colCount = Math.max(
      ...sheet1.columns.map((col, i) => i + 1),
      ...sheet2.columns.map((col, i) => i + 1)
    );

    for (let rowNumber = 1; rowNumber <= rowCount; rowNumber++) {
      const row1 = sheet1.getRow(rowNumber);
      const row2 = sheet2.getRow(rowNumber);

      for (let colNumber = 1; colNumber <= colCount; colNumber++) {
        const cell1 = row1.getCell(colNumber);
        const cell2 = row2.getCell(colNumber);
        const cellLocation = getCellLocation(colNumber, rowNumber);

        // Compare values
        if (
          coalesce(cell1.value?.toString()) !==
          coalesce(cell2.value?.toString())
        ) {
          differences.push({
            sheetName: sheet1.name,
            cell: cellLocation,
            source: coalesce(cell1.value?.toString()),
            target: coalesce(cell2.value?.toString()),
            type: "Different values",
          });
        }

        // Compare styles
        if (JSON.stringify(cell1?.style) !== JSON.stringify(cell2?.style)) {
          const { source: sourceStyle, target: targetStyle } =
            getStyleDifferences(cell1.style, cell2.style);
          differences.push({
            sheetName: sheet1.name,
            cell: cellLocation,
            source: getReadableCellStyle(sourceStyle),
            target: getReadableCellStyle(targetStyle),
            type: "Different styles",
          });
        }
      }
    }

    // Compare conditional formatting rules
    const cf1 = sheet1.conditionalFormattings || [];
    const cf2 = sheet2.conditionalFormattings || [];
    differences.push(...compareConditionalFormattingRules(cf1, cf2));
  });

  const sourceFilename = file1.split("/").pop().split(".")?.[0];
  const targetFilename = file2.split("/").pop().split(".")?.[0];

  await exportToExcel(differences, sourceFilename, targetFilename, "differences.xlsx");

  // console.log("Differences:", differences.length ? differences : "No differences found.");
}

const [,, file1Name, file2Name] = process.argv;
// Resolve file paths to current working directory
const sourcePath = path.resolve(path.join(process.cwd(), 'excel') , file1Name);
const targetPath = path.resolve(path.join(process.cwd(), 'excel'), file2Name);

console.log(sourcePath, targetPath);

compareExcelFiles(sourcePath, targetPath);
