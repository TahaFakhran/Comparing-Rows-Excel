const ExcelJS = require("exceljs");

async function compareExcelRows(filePath, sheetName, row1, row2) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  const worksheet = workbook.getWorksheet(sheetName);

  const row1Values = worksheet.getRow(row1).values;
  const row2Values = worksheet.getRow(row2).values;

  const differingColumns = [];

  for (let col = 1; col <= row1Values.length; col++) {
    if (row1Values[col] !== row2Values[col]) {
      differingColumns.push({
        column: worksheet.getColumn(col).letter,
        row1Value: row1Values[col],
        row2Value: row2Values[col],
      });
    }
  }

  return differingColumns;
}

const filePath = "./ResourcesData.xlsx";
const sheetName = "Sheet1";
const row1 = 1;
const row2 = 3;

compareExcelRows(filePath, sheetName, row1, row2)
  .then((differingColumns) => {
    console.log("Differing columns:");
    differingColumns.forEach(({ column, row1Value, row2Value }) => {
      console.log(`Column ${column}:`);
      console.log(`Row 1: ${row1Value}`);
      console.log("------------");
    });
  })
  .catch((error) => {
    console.error("An error occurred:", error.message);
  });
