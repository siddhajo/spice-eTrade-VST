/**
 * amount-words.js — Replaces AMT.PRG
 * Converts numbers to words in Indian numbering system
 * e.g., 1,23,456.78 → "One Lakh Twenty Three Thousand Four Hundred Fifty Six and Seventy Eight Paise Only"
 */

const ones = ['','One','Two','Three','Four','Five','Six','Seven','Eight','Nine',
  'Ten','Eleven','Twelve','Thirteen','Fourteen','Fifteen','Sixteen','Seventeen','Eighteen','Nineteen'];
const tens = ['','','Twenty','Thirty','Forty','Fifty','Sixty','Seventy','Eighty','Ninety'];

function twoDigit(n) {
  if (n === 0) return '';
  if (n < 20) return ones[n];
  return tens[Math.floor(n / 10)] + (n % 10 ? ' ' + ones[n % 10] : '');
}

function threeDigit(n) {
  if (n === 0) return '';
  const h = Math.floor(n / 100);
  const r = n % 100;
  let s = '';
  if (h) s = ones[h] + ' Hundred';
  if (r) s += (s ? ' ' : '') + twoDigit(r);
  return s;
}

function amountToWords(amount) {
  if (amount === 0) return 'Zero';

  const isNeg = amount < 0;
  amount = Math.abs(amount);

  const rupees = Math.floor(amount);
  const paise = Math.round((amount - rupees) * 100);

  if (rupees === 0 && paise === 0) return 'Zero';

  // Indian system: Crore, Lakh, Thousand, Hundred
  const crore = Math.floor(rupees / 10000000);
  const lakh = Math.floor((rupees % 10000000) / 100000);
  const thousand = Math.floor((rupees % 100000) / 1000);
  const hundred = rupees % 1000;

  let words = '';
  if (crore) words += twoDigit(crore) + ' Crore ';
  if (lakh) words += twoDigit(lakh) + ' Lakh ';
  if (thousand) words += twoDigit(thousand) + ' Thousand ';
  if (hundred) words += threeDigit(hundred);

  words = words.trim();

  // Return words WITHOUT "Only" — callers append "Only" themselves.
  // This avoids "Only Only" when the caller interpolates ` Only` at the end.
  if (words && paise) {
    words = 'Rupees ' + words + ' and ' + twoDigit(paise) + ' Paise';
  } else if (words) {
    words = 'Rupees ' + words;
  } else if (paise) {
    words = twoDigit(paise) + ' Paise';
  }

  if (isNeg) words = 'Minus ' + words;
  return words;
}

module.exports = { amountToWords };
