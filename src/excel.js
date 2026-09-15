import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { amountToWordsLine } from './words';

export async function parseExcelParties(file) {
  const workbook = new ExcelJS.Workbook();
  const buffer = await file.arrayBuffer();
  await workbook.xlsx.load(buffer);
  
  const worksheet = workbook.worksheets[0];
  if (!worksheet) throw new Error("No worksheet found in file");
  
  const parties = [];
  let headerRow = null;
  let colIndices = {
    name: -1,
    bankName: -1,
    accountNo: -1,
    location: -1,
    ifsc: -1
  };
  
  // Find headers by reading rows
  worksheet.eachRow((row, rowNumber) => {
    if (headerRow) return;
    
    const vals = row.values;
    if (!vals || vals.length === 0) return;
    
    for (let i = 1; i < vals.length; i++) {
      const v = String(vals[i] || '').trim().toLowerCase();
      if (!v) continue;
      
      if (v.includes('ifsc') && colIndices.ifsc === -1) {
        colIndices.ifsc = i;
      } else if ((v.includes('account') || v.includes('acc number') || v.includes('acc no') || v.includes('ac no') || v.includes('numb')) && colIndices.accountNo === -1) {
        colIndices.accountNo = i;
      } else if ((v.includes('location') || v.includes('branch')) && colIndices.location === -1) {
        colIndices.location = i;
      } else if (v.includes('bank') && colIndices.bankName === -1) {
        colIndices.bankName = i;
      } else if ((v.includes('parti') || v.includes('particular') || v.includes('beneficiary') || v.includes('name')) && colIndices.name === -1) {
        colIndices.name = i;
      }
    }
    
    if (colIndices.name !== -1 && colIndices.accountNo !== -1) {
      headerRow = rowNumber;
    }
  });
  
  // Default to standard layout if headers not auto-detected
  if (!headerRow) {
    colIndices = { name: 2, bankName: 3, accountNo: 4, location: 5, ifsc: 6 };
    headerRow = 1;
  }
  
  const seenAccountNos = new Set();
  const seenNameAndAccounts = new Set();

  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber <= headerRow) return;
    
    const vals = row.values;
    if (!vals) return;
    
    const name = String(vals[colIndices.name] || '').trim();
    const bankName = String(vals[colIndices.bankName] || '').trim();
    const accountNo = String(vals[colIndices.accountNo] || '').trim();
    const location = String(vals[colIndices.location] || '').trim();
    const ifsc = String(vals[colIndices.ifsc] || '').trim().toUpperCase();
    
    // Normalize account number and composite key for duplicate detection
    const normAcct = accountNo.replace(/[\s-]+/g, '').toLowerCase();
    const compositeKey = `${name.toLowerCase()}_${normAcct}`;
    
    if (name && normAcct) {
      if (!seenAccountNos.has(normAcct) && !seenNameAndAccounts.has(compositeKey)) {
        seenAccountNos.add(normAcct);
        seenNameAndAccounts.add(compositeKey);
        parties.push({
          name,
          bankName: bankName || 'UNKNOWN BANK',
          accountNo,
          location: location || 'BRANCH',
          ifsc: ifsc || ''
        });
      }
    }
  });
  
  return parties;
}

export async function exportToExcel(record, prefix = 'INT ') {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Payment Voucher');

  ws.columns = [
    { width: 8 },   // A SR.NO.
    { width: 32 },  // B PARTICULARS
    { width: 26 },  // C BANK'S NAME
    { width: 20 },  // D BANK ACCOUNT NUMBER
    { width: 16 },  // E BANK'S LOCATION
    { width: 18 },  // F IFSC CODE
    { width: 16 },  // G AMOUNT
  ];

  const headers = ["SR.NO.", "PARTICULARS", "BANK'S NAME", "BANK ACCOUNT NUMBER", "BANK'S LOCATION", "IFSC CODE", "AMOUNT"];
  const headerRow = ws.addRow(headers);
  headerRow.height = 30;
  headerRow.eachCell((cell) => {
    cell.font = { bold: true, size: 11, name: 'Calibri' };
    cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
    cell.border = {
      top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' },
      bottom: { style: 'medium' },
    };
  });

  record.parties.forEach((p, i) => {
    const row = ws.addRow([
      i + 1, 
      p.name.toUpperCase(), 
      p.bankName.toUpperCase(), 
      p.accountNo, 
      p.location.toUpperCase(), 
      p.ifsc.toUpperCase(), 
      p.amount
    ]);
    row.height = 20;
    row.eachCell((cell, colNum) => {
      cell.font = { name: 'Calibri', size: 11 };
      cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
      if (colNum === 1) { 
        cell.alignment = { horizontal: 'center', vertical: 'middle' }; 
        cell.font = { bold: true, name: 'Calibri', size: 11 }; 
      }
      else if (colNum === 7) { 
        cell.alignment = { horizontal: 'right', vertical: 'middle' }; 
        cell.numFmt = '#,##0'; 
      }
      else { 
        cell.alignment = { vertical: 'middle' }; 
      }
    });
  });

  const padRows = Math.max(0, 5 - record.parties.length);
  for (let i = 0; i < padRows; i++) {
    const row = ws.addRow(['', '', '', '', '', '', '']);
    row.height = 20;
    row.eachCell(cell => {
      cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
    });
  }

  const wordsRowIdx = ws.lastRow.number + 1;
  const wordsRow = ws.addRow(['', amountToWordsLine(record.total, prefix), '', '', '', '', record.total]);
  ws.mergeCells(`B${wordsRowIdx}:F${wordsRowIdx}`);
  wordsRow.height = 22;
  const wordsCell = ws.getCell(`B${wordsRowIdx}`);
  wordsCell.font = { bold: true, size: 11, name: 'Calibri' };
  wordsCell.alignment = { horizontal: 'center', vertical: 'middle' };
  const totalCell = ws.getCell(`G${wordsRowIdx}`);
  totalCell.font = { bold: true, name: 'Calibri', size: 11 };
  totalCell.alignment = { horizontal: 'right', vertical: 'middle' };
  totalCell.numFmt = '#,##0';
  totalCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
  
  [`A${wordsRowIdx}`, `B${wordsRowIdx}`, `G${wordsRowIdx}`].forEach(addr => {
    ws.getCell(addr).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
  });

  ws.addRow([]);

  const chequeRowIdx = ws.lastRow.number + 1;
  ws.addRow([]);
  const chqCell = ws.getCell(`A${chequeRowIdx}`);
  chqCell.value = `CHEQUE NO- ${record.chequeNo}`;
  chqCell.font = { bold: true, name: 'Calibri', size: 11 };
  
  const dateCell = ws.getCell(`C${chequeRowIdx}`);
  dateCell.value = `DATE : ${formatDateDDMMYYYY(record.date)}`;
  dateCell.font = { bold: true, name: 'Calibri', size: 11 };

  ws.addRow([]);

  const accLine1 = ws.lastRow.number + 1;
  ws.addRow([]);
  const accCell1 = ws.getCell(`A${accLine1}`);
  accCell1.value = `NAME OF BANK A/C : ${record.account.holderName.toUpperCase()}`;
  accCell1.font = { bold: true, name: 'Calibri', size: 11 };

  const accLine2 = accLine1 + 1;
  ws.addRow([]);
  const accCell2 = ws.getCell(`A${accLine2}`);
  accCell2.value = `NAME OF BANK      : ${record.account.bankName.toUpperCase()}`;
  accCell2.font = { bold: true, name: 'Calibri', size: 11 };

  const accLine3 = accLine2 + 1;
  ws.addRow([]);
  const accCell3 = ws.getCell(`A${accLine3}`);
  accCell3.value = `BANK ACCOUNT NUMBER : ${record.account.accountNo}`;
  accCell3.font = { bold: true, name: 'Calibri', size: 11 };

  ws.pageSetup = { orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0 };

  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/octet-stream' });
  const formattedDate = formatDateDDMMYYYY(record.date);
  saveAs(blob, `RTGS-NEFT PAYMENT LIST ${formattedDate}.xlsx`);
}

export async function exportMonthlyReport(records, year, month) {
  const wb = new ExcelJS.Workbook();
  
  // Sheet 1: Summary
  const wsSummary = wb.addWorksheet('Monthly Summary');
  
  wsSummary.columns = [
    { width: 8 },   // A SR.NO.
    { width: 14 },  // B DATE
    { width: 16 },  // C CHEQUE NO.
    { width: 25 },  // D PAYER BANK
    { width: 32 },  // E BENEFICIARY NAME
    { width: 25 },  // F BENEFICIARY BANK
    { width: 22 },  // G ACCOUNT NO.
    { width: 16 },  // H IFSC CODE
    { width: 16 },  // I AMOUNT
  ];

  const monthName = new Date(year, month - 1).toLocaleString('default', { month: 'long' }).toUpperCase();
  const titleRow = wsSummary.addRow([`COMBINED PAYMENTS SUMMARY - ${monthName} ${year}`]);
  wsSummary.mergeCells('A1:I1');
  titleRow.height = 35;
  const titleCell = wsSummary.getCell('A1');
  titleCell.font = { bold: true, size: 14, name: 'Calibri', color: { argb: 'FFFFFFFF' } };
  titleCell.alignment = { vertical: 'middle', horizontal: 'center' };
  titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1F4E78' } };

  wsSummary.addRow([]);

  const headers = ["SR.NO.", "DATE", "CHEQUE NO.", "PAYER BANK", "BENEFICIARY NAME", "BENEFICIARY BANK", "ACCOUNT NO.", "IFSC CODE", "AMOUNT"];
  const headerRow = wsSummary.addRow(headers);
  headerRow.height = 25;
  headerRow.eachCell((cell) => {
    cell.font = { bold: true, size: 11, name: 'Calibri' };
    cell.alignment = { vertical: 'middle', horizontal: 'center' };
    cell.border = {
      top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' },
      bottom: { style: 'medium' },
    };
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEAEAEA' } };
  });

  let globalSrNo = 1;
  let monthlyTotal = 0;

  const sortedRecords = [...records].sort((a, b) => new Date(a.date) - new Date(b.date));

  sortedRecords.forEach(record => {
    record.parties.forEach(p => {
      const row = wsSummary.addRow([
        globalSrNo++,
        formatDateDDMMYYYY(record.date),
        record.chequeNo,
        (record.account ? record.account.bankName : '').toUpperCase(),
        p.name.toUpperCase(),
        p.bankName.toUpperCase(),
        p.accountNo,
        p.ifsc.toUpperCase(),
        p.amount
      ]);
      row.height = 20;
      monthlyTotal += p.amount;

      row.eachCell((cell, colNum) => {
        cell.font = { name: 'Calibri', size: 11 };
        cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
        
        if (colNum === 1 || colNum === 2 || colNum === 3 || colNum === 8) {
          cell.alignment = { horizontal: 'center', vertical: 'middle' };
        } else if (colNum === 9) {
          cell.alignment = { horizontal: 'right', vertical: 'middle' };
          cell.numFmt = '#,##0';
        } else {
          cell.alignment = { vertical: 'middle' };
        }
      });
    });
  });

  const totalRowIdx = wsSummary.lastRow.number + 1;
  const summaryTotalRow = wsSummary.addRow(['', 'TOTAL MONTHLY OUTFLOW', '', '', '', '', '', '', monthlyTotal]);
  wsSummary.mergeCells(`B${totalRowIdx}:H${totalRowIdx}`);
  summaryTotalRow.height = 24;

  const totalLabelCell = wsSummary.getCell(`B${totalRowIdx}`);
  totalLabelCell.font = { bold: true, size: 11, name: 'Calibri' };
  totalLabelCell.alignment = { horizontal: 'center', vertical: 'middle' };

  const totalValCell = wsSummary.getCell(`I${totalRowIdx}`);
  totalValCell.font = { bold: true, name: 'Calibri', size: 11 };
  totalValCell.alignment = { horizontal: 'right', vertical: 'middle' };
  totalValCell.numFmt = '#,##0';
  totalValCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };

  [`A${totalRowIdx}`, `B${totalRowIdx}`, `I${totalRowIdx}`].forEach(addr => {
    wsSummary.getCell(addr).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
  });

  // Sheet 2: Vouchers Detail
  const wsVouchers = wb.addWorksheet('Vouchers Detail');
  wsVouchers.columns = [
    { width: 8 },   // A SR.NO.
    { width: 32 },  // B PARTICULARS
    { width: 26 },  // C BANK'S NAME
    { width: 20 },  // D BANK ACCOUNT NUMBER
    { width: 16 },  // E BANK'S LOCATION
    { width: 18 },  // F IFSC CODE
    { width: 16 },  // G AMOUNT
  ];

  for (let rIdx = 0; rIdx < sortedRecords.length; rIdx++) {
    const record = sortedRecords[rIdx];
    const prefix = record.prefix ?? 'INT ';

    const headers = ["SR.NO.", "PARTICULARS", "BANK'S NAME", "BANK ACCOUNT NUMBER", "BANK'S LOCATION", "IFSC CODE", "AMOUNT"];
    const hRow = wsVouchers.addRow(headers);
    hRow.height = 30;
    hRow.eachCell((cell) => {
      cell.font = { bold: true, size: 11, name: 'Calibri' };
      cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true };
      cell.border = {
        top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' },
        bottom: { style: 'medium' },
      };
    });

    record.parties.forEach((p, i) => {
      const row = wsVouchers.addRow([
        i + 1, 
        p.name.toUpperCase(), 
        p.bankName.toUpperCase(), 
        p.accountNo, 
        p.location.toUpperCase(), 
        p.ifsc.toUpperCase(), 
        p.amount
      ]);
      row.height = 20;
      row.eachCell((cell, colNum) => {
        cell.font = { name: 'Calibri', size: 11 };
        cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
        if (colNum === 1) { 
          cell.alignment = { horizontal: 'center', vertical: 'middle' }; 
          cell.font = { bold: true, name: 'Calibri', size: 11 }; 
        }
        else if (colNum === 7) { 
          cell.alignment = { horizontal: 'right', vertical: 'middle' }; 
          cell.numFmt = '#,##0'; 
        }
        else { 
          cell.alignment = { vertical: 'middle' }; 
        }
      });
    });

    const padRows = Math.max(0, 5 - record.parties.length);
    for (let i = 0; i < padRows; i++) {
      const row = wsVouchers.addRow(['', '', '', '', '', '', '']);
      row.height = 20;
      row.eachCell(cell => {
        cell.border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
      });
    }

    const wordsRowIdx = wsVouchers.lastRow.number + 1;
    const wordsRow = wsVouchers.addRow(['', amountToWordsLine(record.total, prefix), '', '', '', '', record.total]);
    wsVouchers.mergeCells(`B${wordsRowIdx}:F${wordsRowIdx}`);
    wordsRow.height = 22;
    const wordsCell = wsVouchers.getCell(`B${wordsRowIdx}`);
    wordsCell.font = { bold: true, size: 11, name: 'Calibri' };
    wordsCell.alignment = { horizontal: 'center', vertical: 'middle' };
    const totalCell = wsVouchers.getCell(`G${wordsRowIdx}`);
    totalCell.font = { bold: true, name: 'Calibri', size: 11 };
    totalCell.alignment = { horizontal: 'right', vertical: 'middle' };
    totalCell.numFmt = '#,##0';
    totalCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFFFF00' } };
    
    [`A${wordsRowIdx}`, `B${wordsRowIdx}`, `G${wordsRowIdx}`].forEach(addr => {
      wsVouchers.getCell(addr).border = { top: { style: 'thin' }, left: { style: 'thin' }, right: { style: 'thin' }, bottom: { style: 'thin' } };
    });

    wsVouchers.addRow([]);

    const chequeRowIdx = wsVouchers.lastRow.number + 1;
    wsVouchers.addRow([]);
    const chqCell = wsVouchers.getCell(`A${chequeRowIdx}`);
    chqCell.value = `CHEQUE NO- ${record.chequeNo}`;
    chqCell.font = { bold: true, name: 'Calibri', size: 11 };
    
    const dateCell = wsVouchers.getCell(`C${chequeRowIdx}`);
    dateCell.value = `DATE : ${formatDateDDMMYYYY(record.date)}`;
    dateCell.font = { bold: true, name: 'Calibri', size: 11 };

    wsVouchers.addRow([]);

    const accLine1 = wsVouchers.lastRow.number + 1;
    wsVouchers.addRow([]);
    const accCell1 = wsVouchers.getCell(`A${accLine1}`);
    accCell1.value = `NAME OF BANK A/C : ${record.account.holderName.toUpperCase()}`;
    accCell1.font = { bold: true, name: 'Calibri', size: 11 };

    const accLine2 = accLine1 + 1;
    wsVouchers.addRow([]);
    const accCell2 = wsVouchers.getCell(`A${accLine2}`);
    accCell2.value = `NAME OF BANK      : ${record.account.bankName.toUpperCase()}`;
    accCell2.font = { bold: true, name: 'Calibri', size: 11 };

    const accLine3 = accLine2 + 1;
    wsVouchers.addRow([]);
    const accCell3 = wsVouchers.getCell(`A${accLine3}`);
    accCell3.value = `BANK ACCOUNT NUMBER : ${record.account.accountNo}`;
    accCell3.font = { bold: true, name: 'Calibri', size: 11 };

    if (rIdx < sortedRecords.length - 1) {
      wsVouchers.addRow([]);
      wsVouchers.addRow([]);
      wsVouchers.addRow([]);
    }
  }

  const buffer = await wb.xlsx.writeBuffer();
  const blob = new Blob([buffer], { type: 'application/octet-stream' });
  saveAs(blob, `Monthly-Payments-${monthName}-${year}.xlsx`);
}

function formatDateDDMMYYYY(isoDate) {
  if (!isoDate) return '';
  const [y, m, d] = isoDate.split('-');
  return `${d}-${m}-${y}`;
}
