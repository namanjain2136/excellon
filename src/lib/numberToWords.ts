// Convert numbers to Indian/International English words

const ones = [
  "", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine",
  "ten", "eleven", "twelve", "thirteen", "fourteen", "fifteen", "sixteen",
  "seventeen", "eighteen", "nineteen",
];
const tens = ["", "", "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety"];

function twoDigits(n: number): string {
  if (n < 20) return ones[n];
  const t = Math.floor(n / 10);
  const o = n % 10;
  return tens[t] + (o ? " " + ones[o] : "");
}

function threeDigits(n: number): string {
  const h = Math.floor(n / 100);
  const r = n % 100;
  const parts: string[] = [];
  if (h) parts.push(ones[h] + " hundred");
  if (r) parts.push(twoDigits(r));
  return parts.join(" ");
}

function indianWords(num: number): string {
  if (num === 0) return "zero";
  const parts: string[] = [];
  const crore = Math.floor(num / 10000000);
  num %= 10000000;
  const lakh = Math.floor(num / 100000);
  num %= 100000;
  const thousand = Math.floor(num / 1000);
  num %= 1000;
  const rest = num;
  if (crore) parts.push(indianWords(crore) + " crore");
  if (lakh) parts.push(twoDigits(lakh) + " lakh");
  if (thousand) parts.push(twoDigits(thousand) + " thousand");
  if (rest) parts.push(threeDigits(rest));
  return parts.join(" ").trim();
}

function internationalWords(num: number): string {
  if (num === 0) return "zero";
  const parts: string[] = [];
  const billion = Math.floor(num / 1_000_000_000);
  num %= 1_000_000_000;
  const million = Math.floor(num / 1_000_000);
  num %= 1_000_000;
  const thousand = Math.floor(num / 1000);
  num %= 1000;
  const rest = num;
  if (billion) parts.push(threeDigits(billion) + " billion");
  if (million) parts.push(threeDigits(million) + " million");
  if (thousand) parts.push(threeDigits(thousand) + " thousand");
  if (rest) parts.push(threeDigits(rest));
  return parts.join(" ").trim();
}

export type NumberSystem = "indian" | "international";
export type CaseStyle = "title" | "upper" | "lower";

export interface ConvertOptions {
  system: NumberSystem;
  caseStyle: CaseStyle;
  suffix: string;
  paise: boolean;
}

function applyCase(text: string, style: CaseStyle): string {
  if (style === "upper") return text.toUpperCase();
  if (style === "lower") return text.toLowerCase();
  // title case
  return text.replace(/\w\S*/g, (w) => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase());
}

export function convertNumber(value: number, opts: ConvertOptions): string {
  if (!isFinite(value)) return "";
  const negative = value < 0;
  const abs = Math.abs(value);
  const intPart = Math.floor(abs);
  const decPart = Math.round((abs - intPart) * 100);

  const wordFn = opts.system === "indian" ? indianWords : internationalWords;
  let result = wordFn(intPart);

  if (opts.paise && decPart > 0) {
    result += " rupees and " + twoDigits(decPart) + " paise";
    if (opts.suffix) result += " " + opts.suffix.replace(/^rupees\s+only$/i, "only");
  } else if (opts.suffix) {
    result += " " + opts.suffix;
  }

  if (negative) result = "minus " + result;
  return applyCase(result.replace(/\s+/g, " ").trim(), opts.caseStyle);
}
