export function numberToWordsIndian(num) {
  num = Math.round(Number(num) || 0);
  if (num === 0) return 'ZERO';
  const ones = [
    '', 'ONE', 'TWO', 'THREE', 'FOUR', 'FIVE', 'SIX', 'SEVEN', 'EIGHT', 'NINE', 'TEN',
    'ELEVEN', 'TWELVE', 'THIRTEEN', 'FOURTEEN', 'FIFTEEN', 'SIXTEEN', 'SEVENTEEN', 'EIGHTEEN', 'NINETEEN'
  ];
  const tens = ['', '', 'TWENTY', 'THIRTY', 'FORTY', 'FIFTY', 'SIXTY', 'SEVENTY', 'EIGHTY', 'NINETY'];

  function twoDigit(n) {
    if (n < 20) return ones[n];
    const t = Math.floor(n / 10), o = n % 10;
    return tens[t] + (o ? ' ' + ones[o] : '');
  }
  
  function threeDigit(n) {
    const h = Math.floor(n / 100), rest = n % 100;
    let str = '';
    if (h) str += ones[h] + ' HUNDRED' + (rest ? ' ' : '');
    if (rest) str += twoDigit(rest);
    return str;
  }

  let crore = Math.floor(num / 10000000);
  let lakh = Math.floor((num % 10000000) / 100000);
  let thousand = Math.floor((num % 100000) / 1000);
  let hundred = num % 1000;

  let parts = [];
  if (crore) parts.push(threeDigit(crore) + ' CRORE');
  if (lakh) parts.push(threeDigit(lakh) + ' LAKH');
  if (thousand) parts.push(threeDigit(thousand) + ' THOUSAND');
  if (hundred) parts.push(threeDigit(hundred));

  return parts.join(' ').replace(/\s+/g, ' ').trim();
}

export function amountToWordsLine(num, prefix = 'INT ') {
  const words = numberToWordsIndian(num);
  return `${prefix}${words} ONLY`.replace(/\s+/g, ' ').trim();
}
