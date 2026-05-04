const SHEET_NAME = 'reservations';

const HEADERS = [
  'created_at','booking_no','status','event_date','name','email','phone','language',
  'carrier_count','extra_bag_count','drop_type','pickup_preference','return_preference',
  'payment_method','currency','amount_due','payment_link',
  'confirmed_pickup_time','confirmed_drop_time','pickup_location','drop_location',
  'time_offer_sent_at','payment_request_sent_at','payment_due_at',
  'invoice_id','paid_at','paid_amount','payment_status','confirmation_sent_at','operator_note'
];

function setupCarryGoSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  sh.clear();
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sh.setFrozenRows(1);
  sh.getRange('G:G').setNumberFormat('@'); // phone as text
}

function doPost(e) {
  const body = e.postData && e.postData.contents ? JSON.parse(e.postData.contents) : {};
  const row = buildReservationRow(body);
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  sh.appendRow(row);
  sh.getRange(sh.getLastRow(), 7).setNumberFormat('@');
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, booking_no: row[1], status: row[2] }))
    .setMimeType(ContentService.MimeType.JSON);
}

function buildReservationRow(body) {
  const now = new Date();
  const bookingNo = nextBookingNo_(body.event_date || '');
  const method = body.payment || body.payment_method || '';
  const paypal = String(method).toLowerCase().includes('paypal');
  const currency = paypal ? 'USD' : 'KRW';
  const amount = calculateAmount_(body, paypal);
  const phone = String(body.phone || `${body.phone_country || ''} ${body.phone_number || ''}`).trim();
  const email = body.email || (body.email_id && body.email_domain ? `${body.email_id}@${body.email_domain}` : '');
  const dropType = String(body.return_preference || '').includes('익일') ? 'NEXT_DAY' : 'SAME_DAY';

  return [
    now, bookingNo, 'REQUESTED', body.event_date || '', body.name || '', email,
    phone ? "'" + phone.replace(/^'/, '') : '',
    body.language || '',
    Number(body.carrier_count || 1),
    Number(body.extra_bag_count || 0),
    dropType,
    body.pickup_preference || '',
    body.return_preference || '',
    method,
    currency,
    amount,
    currency === 'KRW' ? 'https://qr.kakaopay.com/FOzMisaMr' : '',
    '', '', '', '', '', '', '', '', '', '', 'UNPAID', '', body.note || ''
  ];
}

function calculateAmount_(body, paypal) {
  const carriers = Number(body.carrier_count || 1);
  const extra = Number(body.extra_bag_count || 0);
  const nextDay = String(body.return_preference || '').includes('익일');
  if (paypal) return carriers * 15 + extra * 10 + (nextDay ? 20 : 0);
  return carriers * 15000 + extra * 10000 + (nextDay ? 20000 : 0);
}

function nextBookingNo_(eventDate) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
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
