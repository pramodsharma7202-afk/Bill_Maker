"""
Helpers for GST calculation and converting rupee amounts into words
using the Indian numbering system (Lakh / Crore), mirroring the
"Amount in Words Helper" sheet found in the Hanuman.xlsx template.
"""

from decimal import Decimal, ROUND_HALF_UP

ONES = [
    "", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine",
    "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen",
    "Seventeen", "Eighteen", "Nineteen",
]
TENS = [
    "", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety",
]


def _two_digits(n: int) -> str:
    if n < 20:
        return ONES[n]
    return (TENS[n // 10] + (" " + ONES[n % 10] if n % 10 else "")).strip()


def _three_digits(n: int) -> str:
    if n >= 100:
        rest = _two_digits(n % 100)
        return ONES[n // 100] + " Hundred" + (" " + rest if rest else "")
    return _two_digits(n)


def number_to_indian_words(num: int) -> str:
    """Convert a non-negative integer into words, Indian numbering system."""
    if num == 0:
        return "Zero"

    crore, num = divmod(num, 10_000_000)
    lakh, num = divmod(num, 100_000)
    thousand, num = divmod(num, 1_000)
    hundred_rem = num

    parts = []
    if crore:
        parts.append(_three_digits(crore) + " Crore")
    if lakh:
        parts.append(_three_digits(lakh) + " Lakh")
    if thousand:
        parts.append(_three_digits(thousand) + " Thousand")
    if hundred_rem:
        parts.append(_three_digits(hundred_rem))

    return " ".join(parts).strip()


def amount_in_words(rupees: float) -> str:
    """Format e.g. 966000.50 -> 'Nine Lakh Sixty Six Thousand Rupees and Fifty Paise Only'."""
    amount = Decimal(str(rupees)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    paise_total = int((amount * 100).to_integral_value(rounding=ROUND_HALF_UP))
    rupees_int, paise = divmod(paise_total, 100)

    words = number_to_indian_words(rupees_int) + " Rupees"
    if paise:
        words += " and " + number_to_indian_words(paise) + " Paise"
    words += " Only"
    return words


def calculate_gst(taxable_value: float, gst_rate: float, gst_type: str):
    """
    gst_type: 'IGST' or 'CGST_SGST'
    Returns a dict with the breakup and grand total.
    """
    gst_rate = float(gst_rate)
    if gst_type == "IGST":
        igst = round(taxable_value * gst_rate / 100, 2)
        cgst = sgst = 0.0
    else:
        half = gst_rate / 2
        cgst = round(taxable_value * half / 100, 2)
        sgst = round(taxable_value * half / 100, 2)
        igst = 0.0

    grand_total = round(taxable_value + igst + cgst + sgst, 2)
    return {
        "taxable_value": taxable_value,
        "gst_rate": gst_rate,
        "gst_type": gst_type,
        "igst": igst,
        "cgst": cgst,
        "sgst": sgst,
        "grand_total": grand_total,
    }
