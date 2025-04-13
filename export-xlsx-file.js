const ExcelJS = require("exceljs");

const defaultHeaderStyle = {
  type: "pattern",
  pattern: "solid",
  fgColor: {
    argb: "4a86e8",
  }, // Blue fill color
};

const defaultColumns = (source, target) => [
  {
    header: "Sheet Name",
    key: "sheetName",
    width: 20,
  },
  { header: "Cell", key: "cell", width: 20 },
  { header: source, key: "source", width: 25 },
  { header: target, key: "target", width: 25 },
  { header: "Type", key: "type", width: 25 },
  {
    header: "Description",
    key: "description",
    width: 30,
  },
  { header: "Index", key: "index", width: 10 },
];

const exportToExcel = async (data, sourceFileName, targetFileName, fileName = "differences.xlsx") => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet("Differences");

  const defaultExcelColumns = defaultColumns(sourceFileName, targetFileName);
  // Define headers
  worksheet.columns = defaultExcelColumns;
  // Add data row
  data.forEach((item) => worksheet.addRow(item));

  defaultExcelColumns.forEach((_, index) => {
    const currentCell = worksheet.getRow(1).getCell(index + 1);
    currentCell.style = {
      font: { bold: true, color: { argb: "FFFFFFFF" } }, // White text
      fill: defaultHeaderStyle,
      alignment: { horizontal: "center", vertical: "middle" },
    };
  });

  // Save to file
  await workbook.xlsx.writeFile(fileName);
  console.log(`Excel file '${fileName}' has been created.`);
};

module.exports = {
  exportToExcel,
};
