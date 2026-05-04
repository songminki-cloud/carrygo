/**
 * CarryGo Google Apps Script MVP
 * - 홈페이지 신청 저장
 * - 예약번호 생성
 * - PayPal 결제 완료 이메일 파싱(1차 자동화 후보)
 *
 * 사용 전 Script Properties 설정:
 * SHEET_ID = Google Sheet ID
 * KAKAOPAY_LINK = https://qr.kakaopay.com/FOzMisaMr
 */

const SHEET_NAME = 'reservations';
const HEADERS = [
  'created_at','booking_no','status','event_date','name','email','phone','language',
  'bag_count','pickup_preference','return_preference','payment_method',
  'currency','amount_due','payment_link',
  'confirmed_pickup_time','confirmed_return_time','pickup_location','return_location',
  'time_offer_sent_at','payment_request_sent_at','payment_due_at',
  'invoice_id','paid_at','paid_amount','payment_status','confirmation_sent_at',
  'operator_note'
];

function setupCarryGoSheet() {
  const ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID'));
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  sh.clear();
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sh.setFrozenRows(1);
}

function doPost(e) {
  const body = e.postData && e.postData.contents ? JSON.parse(e.postData.contents) : {};
  const row = buildReservationRow(body);
  const sh = getSheet_();
  sh.appendRow(row);
  return json_({ ok: true, booking_no: row[1], status: row[2] });
}

function buildReservationRow(body) {
  const now = new Date();
  const bookingNo = nextBookingNo_(body.event_date || 'UNKNOWN');
  const amount = calculateAmount_(body);
  const method = body.payment || body.payment_method || '';
  const currency = method.toLowerCase().includes('paypal') ? 'USD' : 'KRW';
  const kakao = PropertiesService.getScriptProperties().getProperty('KAKAOPAY_LINK') || '';
  return [
    now,
    bookingNo,
    'REQUESTED',
    body.event_date || '',
    body.name || '',
    body.email || '',
    body.phone || '',
    body.language || '',
    body.bag_count || '',
    body.pickup_preference || '',
    body.return_preference || '',
    method,
    currency,
    amount,
    currency === 'KRW' ? kakao : '',
    '', // confirmed_pickup_time
    '', // confirmed_return_time
    '', // pickup_location
    '', // return_location
    '', // time_offer_sent_at
    '', // payment_request_sent_at
    '', // payment_due_at
    '', // invoice_id
    '', // paid_at
    '', // paid_amount
    'UNPAID',
    '', // confirmation_sent_at
    body.note || ''
  ];
}

function calculateAmount_(body) {
  const carriers = Number(body.carrier_count || 1);
  const extra = Number(body.extra_bag_count || 0);
  const nextDay = body.next_day_drop === 'yes';
  const method = String(body.payment || body.payment_method || '').toLowerCase();
  const paypal = method.includes('paypal') || method.includes('international');
  // 온라인 기준: 캐리어 1개 15000원/$15, 추가 가방 10000원/$10, 익일 드랍 20000원/$20.
  if (paypal) return carriers * 15 + extra * 10 + (nextDay ? 20 : 0);
  return carriers * 15000 + extra * 10000 + (nextDay ? 20000 : 0);
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

/**
 * PayPal 결제 완료 Gmail 파싱 후보.
 * 조건: PayPal invoice memo/title에 예약번호(CG-0529-001)를 반드시 넣을 것.
 * 최초 실행 전 Gmail에서 PayPal 결제 완료 메일 실제 제목/본문을 확인해 query를 조정해야 함.
 */
function scanPayPalPaidEmails() {
  const query = 'from:(paypal.com) newer_than:14d (paid OR payment OR invoice)';
  const threads = GmailApp.search(query, 0, 50);
  const sh = getSheet_();
  const data = sh.getDataRange().getValues();
  const bookingCol = HEADERS.indexOf('booking_no');
  const statusCol = HEADERS.indexOf('status');
  const paidAtCol = HEADERS.indexOf('paid_at');
  const paidAmountCol = HEADERS.indexOf('paid_amount');
  const paymentStatusCol = HEADERS.indexOf('payment_status');

  const index = {};
  for (let r = 1; r < data.length; r++) index[String(data[r][bookingCol])] = r + 1;

  threads.forEach(thread => {
    thread.getMessages().forEach(msg => {
      const text = msg.getSubject() + '\n' + msg.getPlainBody();
      const booking = (text.match(/CG-\d{4}-\d{3}/) || [])[0];
      if (!booking || !index[booking]) return;
      const amount = extractUsdAmount_(text);
      const row = index[booking];
      sh.getRange(row, statusCol + 1).setValue('PAID');
      sh.getRange(row, paidAtCol + 1).setValue(msg.getDate());
      if (amount) sh.getRange(row, paidAmountCol + 1).setValue(amount);
      sh.getRange(row, paymentStatusCol + 1).setValue('PAID');
      thread.addLabel(getOrCreateLabel_('CarryGo/PayPalProcessed'));
    });
  });
}

function extractUsdAmount_(text) {
  const m = text.match(/\$\s?([0-9]+(?:\.[0-9]{2})?)/);
  return m ? Number(m[1]) : '';
}

function getSheet_() {
  const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  return SpreadsheetApp.openById(id).getSheetByName(SHEET_NAME);
}

function getOrCreateLabel_(name) {
  return GmailApp.getUserLabelByName(name) || GmailApp.createLabel(name);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
