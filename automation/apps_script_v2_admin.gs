const SHEET_ID = '1X9tp1jU5aT2JtaJpw_HlCdxP2XUjiR2dHp8jxia3bLk';
const SHEET_NAME = 'reservations_v2';

const HEADERS = [
  'created_at',
  'booking_no',
  'status',
  'event_date',
  'name',
  'email',
  'country_code',
  'phone_number',
  'phone_full',
  'language',
  'carrier_count',
  'extra_bag_count',
  'pickup_slot',
  'drop_slot',
  'drop_type',
  'payment_method',
  'currency',
  'amount_due',
  'payment_link',
  'final_pickup_slot',
  'final_drop_slot',
  'payment_due_at',
  'invoice_link',
  'paid_at',
  'paid_amount',
  'payment_status',
  'note'
];

function setupCarryGoSheetV2() {
  const sh = getSheet_();
  sh.clear();
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sh.setFrozenRows(1);
  sh.getRange('G:I').setNumberFormat('@');
  sh.getRange('K:L').setNumberFormat('0');
  sh.getRange('R:R').setNumberFormat('0');
}

function doPost(e) {
  try {
    const body = e.postData && e.postData.contents
      ? JSON.parse(e.postData.contents)
      : {};

    if (body.action === 'mark_paid') return markPaid_(body);
    if (body.action === 'set_offer') return setOffer_(body);

    const row = buildRow_(body);
    const sh = getSheet_();
    sh.appendRow(row);
    sh.getRange(sh.getLastRow(), 7, 1, 3).setNumberFormat('@');

    return json_({ ok: true, booking_no: row[1], status: row[2] });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function buildRow_(body) {
  const now = new Date();
  const bookingNo = nextBookingNo_(body.event_date || '');

  const email = normalizeEmail_(body);
  const countryCode = String(body.phone_country || '').trim();
  const phoneNumber = String(body.phone_number || '').trim();
  const phoneFull = `${countryCode} ${phoneNumber}`.trim();

  const carrierCount = Number(body.carrier_count || 1);
  const extraBagCount = Number(body.extra_bag_count || 0);
  const pickupSlot = String(body.pickup_preference || '').trim();
  const dropSlot = String(body.return_preference || '').trim();
  const dropType = dropSlot.includes('익일') ? 'NEXT_DAY' : 'SAME_DAY';

  const paymentMethod = String(body.payment || body.payment_method || '').trim();
  const isPaypal = paymentMethod.toLowerCase().includes('paypal');
  const currency = isPaypal ? 'USD' : 'KRW';
  const amountDue = calculateAmount_(carrierCount, extraBagCount, dropType, isPaypal);
  const paymentLink = isPaypal ? '' : 'https://qr.kakaopay.com/FOzMisaMr';

  return [
    now,
    bookingNo,
    'REQUESTED',
    body.event_date || '',
    body.name || '',
    email,
    countryCode ? "'" + countryCode : '',
    phoneNumber ? "'" + phoneNumber : '',
    phoneFull ? "'" + phoneFull : '',
    body.language || '',
    carrierCount,
    extraBagCount,
    pickupSlot,
    dropSlot,
    dropType,
    paymentMethod,
    currency,
    amountDue,
    paymentLink,
    '', // final_pickup_slot
    '', // final_drop_slot
    '', // payment_due_at
    '', // invoice_link
    '', // paid_at
    '', // paid_amount
    'UNPAID',
    body.note || ''
  ];
}

function normalizeEmail_(body) {
  if (body.email) return String(body.email).trim();
  const id = String(body.email_id || '').trim();
  const domain = String(body.email_domain || '').trim();
  if (!id || !domain) return '';
  return `${id}@${domain}`;
}

function calculateAmount_(carrierCount, extraBagCount, dropType, isPaypal) {
  if (isPaypal) {
    return carrierCount * 15 + extraBagCount * 10 + (dropType === 'NEXT_DAY' ? 20 : 0);
  }
  return carrierCount * 15000 + extraBagCount * 10000 + (dropType === 'NEXT_DAY' ? 20000 : 0);
}

function nextBookingNo_(eventDate) {
  const sh = getSheet_();
  const key = dateKey_(eventDate);
  const values = sh.getDataRange().getValues();
  let max = 0;

  for (let i = 1; i < values.length; i++) {
    const no = String(values[i][1] || '');
    const m = no.match(new RegExp('CG-' + key + '-(\\d{3})'));
    if (m) max = Math.max(max, Number(m[1]));
  }

  return `CG-${key}-${String(max + 1).padStart(3, '0')}`;
}

function dateKey_(eventDate) {
  const s = String(eventDate || '');
  if (s.includes('5/29')) return '0529';
  if (s.includes('5/30')) return '0530';
  if (s.includes('5/31')) return '0531';
  return Utilities.formatDate(new Date(), 'Asia/Seoul', 'MMdd');
}

function getSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  return sh;
}


function markPaid_(body) {
  verifyAdmin_(body);
  const sh = getSheet_();
  const rowNo = findRowByBookingNo_(body.booking_no);
  const now = new Date();
  setByHeader_(sh, rowNo, 'paid_at', body.paid_at || now);
  setByHeader_(sh, rowNo, 'paid_amount', body.paid_amount || '');
  setByHeader_(sh, rowNo, 'payment_status', 'PAID');
  setByHeader_(sh, rowNo, 'status', body.status || 'PAID');
  if (body.invoice_link) setByHeader_(sh, rowNo, 'invoice_link', body.invoice_link);
  return json_({ ok: true, booking_no: body.booking_no, status: 'PAID' });
}

function setOffer_(body) {
  verifyAdmin_(body);
  const sh = getSheet_();
  const rowNo = findRowByBookingNo_(body.booking_no);
  const now = new Date();
  if (body.final_pickup_slot) setByHeader_(sh, rowNo, 'final_pickup_slot', body.final_pickup_slot);
  if (body.final_drop_slot) setByHeader_(sh, rowNo, 'final_drop_slot', body.final_drop_slot);
  if (body.payment_due_at) setByHeader_(sh, rowNo, 'payment_due_at', body.payment_due_at);
  if (body.invoice_link) setByHeader_(sh, rowNo, 'invoice_link', body.invoice_link);
  setByHeader_(sh, rowNo, 'status', body.status || 'TIME_OFFERED');
  return json_({ ok: true, booking_no: body.booking_no, status: body.status || 'TIME_OFFERED' });
}

function verifyAdmin_(body) {
  const key = PropertiesService.getScriptProperties().getProperty('ADMIN_KEY');
  if (!key) throw new Error('ADMIN_KEY is not set');
  if (body.admin_key !== key) throw new Error('Invalid admin_key');
}

function findRowByBookingNo_(bookingNo) {
  if (!bookingNo) throw new Error('booking_no is required');
  const sh = getSheet_();
  const values = sh.getDataRange().getValues();
  const bookingCol = HEADERS.indexOf('booking_no');
  for (let r = 1; r < values.length; r++) {
    if (String(values[r][bookingCol]) === String(bookingNo)) return r + 1;
  }
  throw new Error('booking_no not found: ' + bookingNo);
}

function setByHeader_(sh, rowNo, header, value) {
  const col = HEADERS.indexOf(header) + 1;
  if (col < 1) throw new Error('header not found: ' + header);
  sh.getRange(rowNo, col).setValue(value);
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
