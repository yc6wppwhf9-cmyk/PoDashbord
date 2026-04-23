const xlsx = require('xlsx');

const filePath = '../Purchase_Order_Reports_Purchase_Order_Report_New-S.xlsx';
const workbook = xlsx.readFile(filePath, { cellDates: true });
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];

const rawData = xlsx.utils.sheet_to_json(worksheet, { header: 1 });
let headerRowIndex = 0;

for (let i = 0; i < Math.min(20, rawData.length); i++) {
  if (!rawData[i]) continue;
  const rowHasPONo = rawData[i].some(cell => typeof cell === 'string' && (cell.includes('PO No') || cell.includes('PO Number')));
  if (rowHasPONo) {
    headerRowIndex = i;
    break;
  }
}

const data = xlsx.utils.sheet_to_json(worksheet, { range: headerRowIndex });

const poSet = new Set();
const openPOSet = new Set();
const dueTodaySet = new Set();
const due7DaysSet = new Set();

const today = new Date();
today.setHours(0,0,0,0);
const sevenDaysFromNow = new Date(today.getTime() + 7 * 24 * 60 * 60 * 1000);

data.forEach(row => {
    const poNo = row['PO No.'] || row['PO No'] || row['Purchase Order'] || row['PO Number'];
    if (!poNo) return;

    poSet.add(poNo);

    const status = row['PO Status'];
    const balanceQty = row['Balance Qty'] || row['PO Pending Qty'] || 0;
    
    let isOpen = false;
    if (balanceQty > 0) isOpen = true;
    if (status && typeof status === 'string' && status.toLowerCase().includes('open')) isOpen = true;
    if (status && typeof status === 'string' && !status.toLowerCase().includes('total received/cancelled') && !status.toLowerCase().includes('closed')) isOpen = true; 
    if (status && typeof status === 'string' && status.toLowerCase().includes('cancelled')) isOpen = false;
    if (status && typeof status === 'string' && status.toLowerCase().includes('total received')) isOpen = false;
    if (balanceQty > 0) isOpen = true;

    if (isOpen) {
        openPOSet.add(poNo);
    }

    const rawDueDate = row['Due Date'] || row['Delivery Date'] || row['SCHEDULE_DATE'] || row['Shedule Date'] || row['Valid Till'];
    if (rawDueDate) {
        let dueDate;
        if (rawDueDate instanceof Date) {
            dueDate = rawDueDate;
            dueDate.setHours(0,0,0,0);
        } else {
            dueDate = new Date(rawDueDate);
            dueDate.setHours(0,0,0,0);
        }

        if (!isNaN(dueDate.getTime())) {
            if (dueDate.getTime() === today.getTime()) {
                dueTodaySet.add(poNo);
            }
            if (dueDate >= today && dueDate <= sevenDaysFromNow) {
                due7DaysSet.add(poNo);
            }
        }
    }
});

console.log('Total POs:', poSet.size);
console.log('Open POs:', openPOSet.size);
console.log('Due Today:', dueTodaySet.size);
console.log('Due in 7 Days:', due7DaysSet.size);
