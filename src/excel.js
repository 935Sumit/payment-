import ExcelJS from 'exceljs';
import fileSaver from 'file-saver';
const saveAs = fileSaver?.saveAs || fileSaver;
import { amountToWordsLine } from './words.js';

export function getCellValueString(val) {
  if (val === null || val === undefined) return '';
  if (typeof val === 'object') {
    if (val.text !== undefined) return String(val.text).trim();
    if (val.result !== undefined) return String(val.result).trim();
    if (Array.isArray(val.richText)) {
      return val.richText.map(t => t.text || '').join('').trim();
    }
  }
  if (typeof val === 'number') {
    return Number.isInteger(val) ? val.toString() : val.toFixed(0);
  }
  return String(val).trim();
}

function cleanHeaderToken(str) {
  return String(str || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]/g, ' ')
    .replace(/\s+/g, ' ');
}

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
    ifsc: -1,
    srNo: -1
  };
  
  // Find header row by scoring rows up to row 15
  let maxScore = 0;
  let bestHeaderRow = null;
  let bestColIndices = { ...colIndices };

  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber > 15) return;
    
    const vals = row.values;
    if (!vals || vals.length === 0) return;
    
    const currentIndices = { name: -1, bankName: -1, accountNo: -1, location: -1, ifsc: -1, srNo: -1 };
    let score = 0;

    for (let i = 1; i < vals.length; i++) {
      const cellText = getCellValueString(vals[i]);
      if (!cellText) continue;
      const h = cleanHeaderToken(cellText);

      // Check Sr / Serial Number first (to strictly prevent matching as account number)
      const isSrNo = /\b(sr|serial|sl|sno|srno|s no|sl no|seq|sr num)\b/.test(h);
      if (isSrNo && currentIndices.srNo === -1) {
        currentIndices.srNo = i;
        score += 1;
        continue;
      }

      // Check IFSC
      const isIfsc = /\b(ifsc|ifs)\b/.test(h);
      if (isIfsc && currentIndices.ifsc === -1) {
        currentIndices.ifsc = i;
        score += 3;
        continue;
      }

      // Check Bank Account Number
      const isAccount = !isSrNo && !/\b(phone|mobile|contact|aadhaar|aadhar|pan|cheque|chq|voucher|bill|invoice|id|emp id)\b/.test(h) && (
        /\b(bank\s*)?(a\s*c|acc|acct|account)\s*(no|num|numb|number|code)?\b/.test(h) ||
        /\b(a\s*c\s*no|acc\s*no|ac\s*no|acct\s*no)\b/.test(h) ||
        /\b(bank\s*account|bank\s*a\s*c)\b/.test(h)
      );
      if (isAccount && currentIndices.accountNo === -1) {
        currentIndices.accountNo = i;
        score += 4;
        continue;
      }

      // Check Location / Branch
      const isLocation = /\b(location|branch|city|place)\b/.test(h);
      if (isLocation && currentIndices.location === -1) {
        currentIndices.location = i;
        score += 2;
        continue;
      }

      // Check Bank Name
      const isBankName = !isAccount && !isLocation && !isIfsc && /\b(bank|b\s*k)\b/.test(h);
      if (isBankName && currentIndices.bankName === -1) {
        currentIndices.bankName = i;
        score += 3;
        continue;
      }

      // Check Party / Beneficiary / Employee Name
      const isName = !isBankName && !isLocation && !isIfsc && !isAccount && !isSrNo && (
        /\b(particular|particulars|beneficiary|payee|employee|emp\s*name|staff|worker|party|holder|client|vendor|person|name)\b/.test(h)
      );
      if (isName && currentIndices.name === -1) {
        currentIndices.name = i;
        score += 3;
        continue;
      }
    }

    if (score > maxScore && (currentIndices.name !== -1 || currentIndices.accountNo !== -1 || currentIndices.ifsc !== -1)) {
      maxScore = score;
      bestHeaderRow = rowNumber;
      bestColIndices = { ...currentIndices };
    }
  });

  if (bestHeaderRow && maxScore >= 3) {
    headerRow = bestHeaderRow;
    colIndices = bestColIndices;
  } else {
    // Default fallback indices if no clear headers detected
    colIndices = { name: 2, bankName: 3, accountNo: 4, location: 5, ifsc: 6, srNo: 1 };
    headerRow = 1;
  }

  // Data inspection heuristic fallback if crucial columns were not identified
  if (colIndices.accountNo === -1 || colIndices.ifsc === -1 || colIndices.name === -1) {
    const ifscRegex = /^[A-Z]{4}0[A-Z0-9]{6}$/i;
    const colStats = {};

    worksheet.eachRow((row, rowNumber) => {
      if (rowNumber <= headerRow || rowNumber > headerRow + 10) return;
      const vals = row.values;
      if (!vals) return;
      for (let i = 1; i < vals.length; i++) {
        const valStr = getCellValueString(vals[i]).trim();
        if (!valStr) continue;
        if (!colStats[i]) colStats[i] = { ifscMatches: 0, digitsLenGt8: 0, textOnly: 0, bankKeywords: 0 };
        if (ifscRegex.test(valStr)) colStats[i].ifscMatches++;
        if (/^\d{8,20}$/.test(valStr.replace(/[\s-]/g, ''))) colStats[i].digitsLenGt8++;
        if (/[a-zA-Z]{3,}/.test(valStr) && !/^\d+$/.test(valStr)) colStats[i].textOnly++;
        if (/\b(sbi|hdfc|icici|axis|punjab|canara|union|bank|kotak|baroda|indusind|idbi|yes|indian)\b/i.test(valStr)) colStats[i].bankKeywords++;
      }
    });

    Object.entries(colStats).forEach(([colIdxStr, stats]) => {
      const idx = parseInt(colIdxStr, 10);
      if (idx === colIndices.srNo) return;
      if (colIndices.ifsc === -1 && stats.ifscMatches >= 2) colIndices.ifsc = idx;
      if (colIndices.accountNo === -1 && stats.digitsLenGt8 >= 2) colIndices.accountNo = idx;
      if (colIndices.bankName === -1 && stats.bankKeywords >= 2) colIndices.bankName = idx;
    });
  }

  // Safety fallback if still missing: ensure accountNo is never assigned to srNo
  if (colIndices.accountNo === colIndices.srNo && colIndices.accountNo !== -1) {
    colIndices.accountNo = -1;
  }

  const seenAccountsInSheet = new Set();
  const seenNameAndAccountsInSheet = new Set();
  let inSheetDuplicatesCount = 0;
  let totalRowsRead = 0;

  worksheet.eachRow((row, rowNumber) => {
    if (rowNumber <= headerRow) return;
    
    const vals = row.values;
    if (!vals || vals.length === 0) return;
    
    const name = colIndices.name !== -1 ? getCellValueString(vals[colIndices.name]) : '';
    const bankName = colIndices.bankName !== -1 ? getCellValueString(vals[colIndices.bankName]) : '';
    const rawAccountNo = colIndices.accountNo !== -1 ? getCellValueString(vals[colIndices.accountNo]) : '';
    const location = colIndices.location !== -1 ? getCellValueString(vals[colIndices.location]) : '';
    const ifsc = colIndices.ifsc !== -1 ? getCellValueString(vals[colIndices.ifsc]).toUpperCase() : '';
    
    // Normalize account number and composite key for duplicate detection
    const normAcct = rawAccountNo.replace(/[\s-]+/g, '').toLowerCase();
    const compositeKey = `${name.toLowerCase()}_${normAcct}`;
    
    if (!name && !normAcct) return; // Skip empty rows
    totalRowsRead++;

    // Check if this row is an internal duplicate within the Excel file itself
    const isDupInSheet = (normAcct && seenAccountsInSheet.has(normAcct)) || 
                         (normAcct && seenNameAndAccountsInSheet.has(compositeKey)) ||
                         (!normAcct && name && seenNameAndAccountsInSheet.has(compositeKey));

    if (isDupInSheet) {
      inSheetDuplicatesCount++;
      return; // Ignore duplicate row in Excel, continue to next
    }

    if (normAcct) seenAccountsInSheet.add(normAcct);
    seenNameAndAccountsInSheet.add(compositeKey);

    parties.push({
      name: name || 'UNKNOWN PAYEE',
      bankName: bankName || 'UNKNOWN BANK',
      accountNo: rawAccountNo,
      location: location || 'BRANCH',
      ifsc: ifsc || ''
    });
  });
  
  parties.inSheetDuplicatesCount = inSheetDuplicatesCount;
  parties.totalRowsRead = totalRowsRead;
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
