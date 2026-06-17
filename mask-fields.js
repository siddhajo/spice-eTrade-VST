/**
 * mask-fields.js — shared sensitive-field masking.
 *
 * One masking function used by every server-side surface that renders a
 * bank account number, IFSC code or phone number into a customer-facing
 * receipt or on-screen display. The browser (public/index.html and
 * public-mobile/app.html) carries a verbatim copy of `maskField` so the
 * UI and the generated PDFs mask identically.
 *
 * Modes (per-field, configured in Settings → Display):
 *   none    → value shown in full (default — opt-in security feature)
 *   last4   → all but the last 4 characters replaced with '*'
 *   last6   → all but the last 6 characters replaced with '*'
 *   first4  → all but the first 4 characters replaced with '*'
 *   first6  → all but the first 6 characters replaced with '*'
 *
 * '*' fills the hidden positions and the length is preserved, so the
 * reader can still gauge how long the real number is. A value at or below
 * the visible-digit count is returned unchanged — there's nothing useful
 * left to hide and masking it would only reveal its (short) length.
 *
 * IMPORTANT: this is for DISPLAY / customer-facing PDFs only. Functional
 * machine-consumed outputs — the bank NEFT/RTGS payment file, the payment
 * advice PDF, DBF exports, Tally XML, and WhatsApp send targets — must
 * keep the real numbers and never call this.
 */
function maskField(value, mode) {
  const s = value == null ? '' : String(value);
  if (!s || !mode || mode === 'none') return s;
  const show = (mode === 'last4' || mode === 'first4') ? 4
             : (mode === 'last6' || mode === 'first6') ? 6
             : 0;
  if (!show || s.length <= show) return s;
  const stars = '*'.repeat(s.length - show);
  return (mode === 'last4' || mode === 'last6')
    ? stars + s.slice(-show)
    : s.slice(0, show) + stars;
}

module.exports = { maskField };
