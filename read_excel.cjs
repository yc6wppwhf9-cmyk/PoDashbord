const xlsx = require('xlsx');
const filePath = '../Purchase_Order_Reports_Purchase_Order_Report_New-S.xlsx';

try {
  const workbook = xlsx.readFile(filePath);
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  const data = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
  
  if (data.length > 0) {
    console.log("First 15 Rows:");
    for (let i = 0; i < Math.min(15, data.length); i++) {
        console.log(`Row ${i}:`, data[i]);
    }
  } else {
    console.log("Sheet is empty.");
  }
} catch (error) {
  console.error("Error reading file:", error.message);
}
