/*
 * CarryGo Apps Script - Final Sheet Setup
 *
 * Safe first step for the final implementation.
 * This file creates/updates the final sheet structure without deleting existing data.
 * Do NOT run resetCarryGoSheetsFinal_() unless you intentionally want to clear data.
 */

const CARRYGO_SHEET_ID = '1X9tp1jU5aT2JtaJpw_HlCdxP2XUjiR2dHp8jxia3bLk';
const CARRYGO_TIMEZONE = 'Asia/Seoul';
const CARRYGO_AGREEMENT_VERSION = '2026-05-05-v1';

const CARRYGO_SHEETS = {
  CONCERTS: 'concerts',
  CONCERT_DATES: 'concert_dates',
  RESERVATIONS: 'reservations',
  STAFF: 'staff'
};

const CONCERTS_HEADERS = [
  'concert_id',
  'concert_code',
  'concert_title',
  'venue',
  'is_active',
  'sort_order',
  'created_at',
  'updated_at',
  'city'
];

const CONCERT_DATES_HEADERS = [
  'concert_date_id',
  'concert_id',
  'concert_date',
  'concert_time',
  'pickup_drop_guide_link',
  'next_day_pickup_guide_link',
  'location_change_guide_link',
  'is_active',
  'sort_order',
  'created_at',
  'updated_at',
  'pickup_time_options'
];

const RESERVATIONS_HEADERS = [
  'reservation_id',
  'status',
  'created_at',
  'confirmed_at',
  'picked_up_at',
  'returned_at',
  'cancelled_at',
  'refunded_at',
  'concert_id',
  'concert_date_id',
  'concert_code',
  'concert_title',
  'concert_date',
  'concert_time',
  'venue',
  'customer_country',
  'country_code',
  'customer_name',
  'customer_email',
  'phone_number',
  'phone_full',
  'payment_method',
  'base_fee',
  'currency',
  'payment_status',
  'payment_due_at',
  'paid_at',
  'paid_amount',
  'refund_amount',
  'refund_method',
  'refund_note',
  'expected_suitcase_count',
  'expected_extra_bag_count',
  'next_day_pickup_required',
  'next_day_pickup_fee_status',
  'checkin_token',
  'qr_checkin_url',
  'qr_file_url',
  'picked_up_by',
  'agreement_version',
  'note',
  'pickup_time',
  'booking_channel',
  'luggage_tag_numbers',
  'onsite_payment_method',
  'onsite_staff',
  'onsite_consent_flags',
  'actual_suitcase_count',
  'actual_extra_bag_count',
  'onsite_due_amount',
  'onsite_cash_received',
  'onsite_tag_attached',
  'onsite_photo_taken',
  'onsite_checkin_completed_at'
];

const STAFF_HEADERS = [
  'staff_id',
  'staff_name',
  'staff_code',
  'is_active',
  'created_at',
  'updated_at'
];

const FINAL_SCHEMA = [
  { name: CARRYGO_SHEETS.CONCERTS, headers: CONCERTS_HEADERS },
  { name: CARRYGO_SHEETS.CONCERT_DATES, headers: CONCERT_DATES_HEADERS },
  { name: CARRYGO_SHEETS.RESERVATIONS, headers: RESERVATIONS_HEADERS },
  { name: CARRYGO_SHEETS.STAFF, headers: STAFF_HEADERS }
];

/**
 * SAFE setup.
 * Creates final sheets and writes headers.
 * Existing data rows are preserved.
 * If headers differ, the header row is replaced but rows below remain untouched.
 */
function setupCarryGoSheetsFinal() {
  const ss = SpreadsheetApp.openById(CARRYGO_SHEET_ID);
  FINAL_SCHEMA.forEach(schema => {
    const sh = getOrCreateSheetFinal_(ss, schema.name);
    writeHeadersPreserveRows_(sh, schema.headers);
    applySheetFormatsFinal_(sh, schema.name, schema.headers);
  });
  return 'CarryGo final sheets setup complete.';
}

/**
 * Optional sample seed for development.
 * Safe: only inserts sample rows if the target sheet has no data rows.
 */
function seedCarryGoSampleDataFinal() {
  const ss = SpreadsheetApp.openById(CARRYGO_SHEET_ID);
  const now = new Date();

  const concerts = ss.getSheetByName(CARRYGO_SHEETS.CONCERTS);
  if (concerts && concerts.getLastRow() <= 1) {
    concerts.appendRow([
      'shinee_world_vii',
      'SN',
      'SHINee WORLD VII',
      'KSPO DOME',
      true,
      1,
      now,
      now
    ]);
  }

  const dates = ss.getSheetByName(CARRYGO_SHEETS.CONCERT_DATES);
  if (dates && dates.getLastRow() <= 1) {
    dates.appendRow(['shinee_20260530', 'shinee_world_vii', '2026-05-30', '19:00', '', '', '', true, 1, now, now]);
    dates.appendRow(['shinee_20260531', 'shinee_world_vii', '2026-05-31', '18:00', '', '', '', true, 2, now, now]);
    dates.appendRow(['shinee_20260601', 'shinee_world_vii', '2026-06-01', '17:00', '', '', '', true, 3, now, now]);
  }

  const staff = ss.getSheetByName(CARRYGO_SHEETS.STAFF);
  if (staff && staff.getLastRow() <= 1) {
    staff.appendRow(['JD', 'JD', '0530', true, now, now]);
  }

  return 'CarryGo sample data seed complete.';
}

/**
 * DESTRUCTIVE reset helper.
 * Intentionally not exposed as the default setup function.
 * Run only when you want to clear all final sheets and recreate headers.
 */
function resetCarryGoSheetsFinal_() {
  const ss = SpreadsheetApp.openById(CARRYGO_SHEET_ID);
  FINAL_SCHEMA.forEach(schema => {
    const sh = getOrCreateSheetFinal_(ss, schema.name);
    sh.clear();
    writeHeadersPreserveRows_(sh, schema.headers);
    applySheetFormatsFinal_(sh, schema.name, schema.headers);
  });
  return 'CarryGo final sheets reset complete.';
}

function getOrCreateSheetFinal_(ss, sheetName) {
  let sh = ss.getSheetByName(sheetName);
  if (!sh) sh = ss.insertSheet(sheetName);
  return sh;
}

function writeHeadersPreserveRows_(sh, headers) {
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.setFrozenRows(1);

  const extraCols = sh.getMaxColumns() - headers.length;
  if (extraCols > 0) {
    // Keep extra columns rather than deleting them, to avoid accidental data loss.
    // They can be manually removed after review if needed.
  }
}

function applySheetFormatsFinal_(sh, sheetName, headers) {
  sh.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#111111')
    .setFontColor('#ffffff');

  sh.autoResizeColumns(1, headers.length);

  formatColumnsAsText_(sh, headers, [
    'reservation_id',
    'concert_id',
    'concert_date_id',
    'concert_code',
    'customer_email',
    'country_code',
    'phone_number',
    'phone_full',
    'payment_method',
    'currency',
    'payment_status',
    'checkin_token',
    'qr_checkin_url',
    'qr_file_url',
    'picked_up_by',
    'agreement_version',
    'staff_id',
    'staff_code'
  ]);

  formatColumnsAsNumber_(sh, headers, [
    'sort_order',
    'base_fee',
    'paid_amount',
    'refund_amount',
    'expected_suitcase_count',
    'expected_extra_bag_count'
  ]);

  formatColumnsAsDateTime_(sh, headers, [
    'created_at',
    'updated_at',
    'confirmed_at',
    'picked_up_at',
    'returned_at',
    'cancelled_at',
    'refunded_at',
    'payment_due_at',
    'paid_at'
  ]);

  if (sheetName === CARRYGO_SHEETS.CONCERT_DATES || sheetName === CARRYGO_SHEETS.RESERVATIONS) {
    formatColumnsAsText_(sh, headers, ['concert_date', 'concert_time', 'pickup_time', 'pickup_time_options']);
  }
}

function formatColumnsAsText_(sh, headers, names) {
  names.forEach(name => {
    const idx = headers.indexOf(name);
    if (idx >= 0) sh.getRange(1, idx + 1, sh.getMaxRows(), 1).setNumberFormat('@');
  });
}

function formatColumnsAsNumber_(sh, headers, names) {
  names.forEach(name => {
    const idx = headers.indexOf(name);
    if (idx >= 0) sh.getRange(2, idx + 1, Math.max(sh.getMaxRows() - 1, 1), 1).setNumberFormat('0');
  });
}

function formatColumnsAsDateTime_(sh, headers, names) {
  names.forEach(name => {
    const idx = headers.indexOf(name);
    if (idx >= 0) sh.getRange(2, idx + 1, Math.max(sh.getMaxRows() - 1, 1), 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  });
}

// ===== CarryGo Final Public API =====

function doGet(e) {
  try {
    const params = (e && e.parameter) ? e.parameter : {};
    const mode = String(params.mode || '').trim();
    const action = String(params.action || '').trim();

    if (mode === 'checkin') {
      return renderCheckinPageFinal_(params);
    }

    if (mode === 'staff_login') {
      return renderStaffLoginPageFinal_(params);
    }

    if (mode === 'staff_logout') {
      return renderStaffLogoutPageFinal_();
    }

    if (mode === 'pickup_complete') {
      return handlePickupCompleteFinal_(params);
    }

    if (mode === 'staff_login_api') {
      return staffLoginApiFinal_(params);
    }

    if (mode === 'pickup_complete_api') {
      return pickupCompleteApiFinal_(params);
    }

    if (mode === 'onsite_checkin_api') {
      return onsiteCheckinApiFinal_(params);
    }

    if (mode === 'onsite_lookup_api') {
      return onsiteLookupApiFinal_(params);
    }

    if (mode === 'admin_normalize_luggage_tags') {
      return adminNormalizeLuggageTagsApiFinal_(params);
    }

    if (mode === 'admin_reset_checkin_tests') {
      return adminResetCheckinTestsApiFinal_(params);
    }


    if (mode === 'staff_session_api') {
      return staffSessionApiFinal_(params);
    }

    if (mode === 'admin') {
      return renderAdminPageFinal_(params);
    }

    if (mode === 'admin_list_unpaid') {
      return adminListUnpaidApiFinal_(params);
    }

    if (mode === 'admin_confirm_payment') {
      return adminConfirmPaymentApiFinal_(params);
    }

    if (mode === 'admin_cancel_expired_unpaid') {
      return adminCancelExpiredUnpaidApiFinal_(params);
    }

    if (mode === 'admin_list') {
      return adminListByStatusApiFinal_(params);
    }

    if (mode === 'admin_update_status') {
      return adminUpdateStatusApiFinal_(params);
    }

    if (mode === 'admin_refund') {
      return adminRefundApiFinal_(params);
    }

    if (mode === 'admin_create_concert') {
      return adminCreateConcertApiFinal_(params);
    }

    if (mode === 'admin_create_concert_date') {
      return adminCreateConcertDateApiFinal_(params);
    }

    if (mode === 'admin_create_concert_bundle') {
      return adminCreateConcertBundleApiFinal_(params);
    }

    if (mode === 'admin_set_active') {
      return adminSetActiveApiFinal_(params);
    }

    if (mode === 'admin_update_concert') {
      return adminUpdateConcertApiFinal_(params);
    }

    if (mode === 'admin_update_concert_date') {
      return adminUpdateConcertDateApiFinal_(params);
    }

    if (mode === 'admin_delete_concert') {
      return adminDeleteConcertApiFinal_(params);
    }

    if (mode === 'admin_delete_concert_date') {
      return adminDeleteConcertDateApiFinal_(params);
    }

    if (mode === 'admin_concerts') {
      return adminConcertsApiFinal_(params);
    }

    if (mode === 'admin_confirm_payment_page') {
      return renderAdminConfirmPaymentPageFinal_(params);
    }

    if (mode === 'admin_confirm_selected_page') {
      return renderAdminConfirmSelectedPageFinal_(params);
    }

    if (action === 'get_active_booking_data') {
      return jsonFinal_({ ok: true, concerts: getActiveBookingDataFinal() });
    }

    if (action === 'get_active_concerts') {
      return jsonFinal_({ ok: true, concerts: getActiveConcertsFinal() });
    }

    if (action === 'get_active_concert_dates') {
      return jsonFinal_({
        ok: true,
        concert_dates: getActiveConcertDatesFinal(params.concert_id)
      });
    }

    return jsonFinal_({ ok: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return jsonFinal_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function doPost(e) {
  try {
    const body = parseBodyFinal_(e);
    const action = String(body.action || '').trim();

    if (action === 'create_reservation') {
      return jsonFinal_({ ok: true, reservation: createReservationFinal(body) });
    }

    if (action === 'create_walkin_reservation') {
      return jsonFinal_({ ok: true, reservation: createWalkinReservationFinal(body) });
    }

    return jsonFinal_({ ok: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return jsonFinal_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function getActiveConcertsFinal() {
  const rows = readSheetObjectsFinal_(CARRYGO_SHEETS.CONCERTS, CONCERTS_HEADERS);
  return rows
    .filter(row => isActiveFinal_(row.is_active))
    .sort((a, b) => Number(a.sort_order || 9999) - Number(b.sort_order || 9999))
    .map(row => ({
      concert_id: row.concert_id,
      concert_code: row.concert_code,
      concert_title: row.concert_title,
      venue: row.venue,
      city: String(row.city || 'SEOUL').trim(),
      sort_order: Number(row.sort_order || 0)
    }));
}

function getActiveConcertDatesFinal(concertId) {
  if (!concertId) throw new Error('concert_id is required');
  const rows = readSheetObjectsFinal_(CARRYGO_SHEETS.CONCERT_DATES, CONCERT_DATES_HEADERS);
  return rows
    .filter(row => String(row.concert_id) === String(concertId))
    .filter(row => isActiveFinal_(row.is_active))
    .sort((a, b) => Number(a.sort_order || 9999) - Number(b.sort_order || 9999))
    .map(row => ({
      concert_date_id: row.concert_date_id,
      concert_id: row.concert_id,
      concert_date: row.concert_date,
      concert_time: row.concert_time,
      pickup_time_options: normalizePickupTimeOptionsFinal_(row.pickup_time_options).join(','),
      pickup_drop_guide_link: row.pickup_drop_guide_link,
      next_day_pickup_guide_link: row.next_day_pickup_guide_link,
      location_change_guide_link: row.location_change_guide_link,
      sort_order: Number(row.sort_order || 0)
    }));
}

function getActiveBookingDataFinal() {
  return getActiveConcertsFinal().map(concert => {
    const out = Object.assign({}, concert);
    out.concert_dates = getActiveConcertDatesFinal(concert.concert_id);
    return out;
  });
}

function createReservationFinal(body) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const now = new Date();
    const concert = findActiveConcertFinal_(body.concert_id);
    const concertDate = findActiveConcertDateFinal_(body.concert_date_id, concert.concert_id);
    const payment = resolvePaymentFinal_(body.payment_method);
    const reservationId = nextReservationIdFinal_(concert.concert_code, concertDate.concert_date);
    const paymentDueAt = new Date(now.getTime() + 6 * 60 * 60 * 1000);

    const countryCode = normalizeCountryCodeFinal_(body.country_code);
    const phoneNumber = String(body.phone_number || '').trim();
    const phoneFull = normalizePhoneFullFinal_(countryCode, phoneNumber);

    const rowObject = {
      reservation_id: reservationId,
      status: 'UNPAID',
      created_at: now,
      confirmed_at: '',
      picked_up_at: '',
      returned_at: '',
      cancelled_at: '',
      refunded_at: '',
      concert_id: concert.concert_id,
      concert_date_id: concertDate.concert_date_id,
      concert_code: concert.concert_code,
      concert_title: concert.concert_title,
      concert_date: concertDate.concert_date,
      concert_time: concertDate.concert_time,
      pickup_time: requiredPickupTimeFinal_(body.pickup_time, concertDate.pickup_time_options),
      venue: concert.venue,
      customer_country: requiredStringFinal_(body.customer_country, 'customer_country'),
      country_code: countryCode,
      customer_name: requiredStringFinal_(body.customer_name, 'customer_name'),
      customer_email: normalizeEmailFinal_(body.customer_email),
      phone_number: phoneNumber,
      phone_full: phoneFull,
      payment_method: payment.method,
      base_fee: payment.base_fee,
      currency: payment.currency,
      payment_status: 'UNPAID',
      payment_due_at: paymentDueAt,
      paid_at: '',
      paid_amount: '',
      refund_amount: '',
      refund_method: '',
      refund_note: '',
      expected_suitcase_count: normalizeCountFinal_(body.expected_suitcase_count, 1),
      expected_extra_bag_count: normalizeCountFinal_(body.expected_extra_bag_count, 0),
      next_day_pickup_required: 'NO',
      next_day_pickup_fee_status: 'NONE',
      checkin_token: '',
      qr_checkin_url: '',
      qr_file_url: '',
      picked_up_by: '',
      agreement_version: CARRYGO_AGREEMENT_VERSION,
      note: String(body.note || '').trim(),
      booking_channel: 'ONLINE',
      luggage_tag_numbers: '',
      onsite_payment_method: '',
      onsite_staff: '',
      onsite_consent_flags: '',
      actual_suitcase_count: '',
      actual_extra_bag_count: '',
      onsite_due_amount: '',
      onsite_cash_received: '',
      onsite_tag_attached: '',
      onsite_photo_taken: '',
      onsite_checkin_completed_at: ''
    };

    validateReservationFinal_(rowObject);

    const sh = getSheetFinal_(CARRYGO_SHEETS.RESERVATIONS);
    sh.appendRow(RESERVATIONS_HEADERS.map(header => rowObject[header] !== undefined ? rowObject[header] : ''));
    applySheetFormatsFinal_(sh, CARRYGO_SHEETS.RESERVATIONS, RESERVATIONS_HEADERS);

    sendPaymentInstructionEmailFinal_(rowObject);

    return {
      reservation_id: reservationId,
      status: rowObject.status,
      payment_status: rowObject.payment_status,
      payment_method: payment.method,
      base_fee: payment.base_fee,
      currency: payment.currency,
      amount_display: payment.amount_display,
      payment_due_at: formatDateTimeFinal_(paymentDueAt),
      concert_title: rowObject.concert_title,
      concert_date: rowObject.concert_date,
      concert_time: rowObject.concert_time,
      pickup_time: rowObject.pickup_time,
      venue: rowObject.venue,
      expected_suitcase_count: rowObject.expected_suitcase_count,
      expected_extra_bag_count: rowObject.expected_extra_bag_count,
      payment_instructions: buildPaymentInstructionsFinal_(payment.method, reservationId)
    };
  } finally {
    lock.releaseLock();
  }
}


function createWalkinReservationFinal(body) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const staff = findStaffByCodeFinal_(body.staff_code);
    if (!staff) throw new Error('Invalid staff_code');

    const now = new Date();
    const concert = findActiveConcertFinal_(body.concert_id);
    const concertDate = findActiveConcertDateFinal_(body.concert_date_id, concert.concert_id);
    const reservationId = nextWalkinReservationIdFinal_(concertDate.concert_date);
    const suitcaseCount = normalizeCountFinal_(body.expected_suitcase_count, 1);
    const extraBagCount = normalizeCountFinal_(body.expected_extra_bag_count, 0);
    if (suitcaseCount < 1) throw new Error('expected_suitcase_count must be at least 1');
    const paidAmount = suitcaseCount * 20000 + extraBagCount * 10000;
    const countryCode = normalizeCountryCodeFinal_(body.country_code || '+82');
    const phoneNumber = String(body.phone_number || '').trim();
    const phoneFull = normalizePhoneFullFinal_(countryCode, phoneNumber);
    const tagNumbers = nextLuggageTagNumbersFinal_(suitcaseCount + extraBagCount, concert.concert_id);
    const consentFlags = 'HARDCOPY_CONFIRMED=' + String(body.hardcopy_confirmed || 'NO').trim().toUpperCase();
    if (String(body.hardcopy_confirmed || '').trim().toUpperCase() !== 'YES') throw new Error('hardcopy confirmation is required');
    if (String(body.onsite_cash_received || '').trim().toUpperCase() !== 'YES') throw new Error('cash received confirmation is required');
    if (String(body.onsite_tag_attached || '').trim().toUpperCase() !== 'YES') throw new Error('tag attachment confirmation is required');
    if (String(body.onsite_photo_taken || '').trim().toUpperCase() !== 'YES') throw new Error('photo confirmation is required');

    const rowObject = {
      reservation_id: reservationId,
      status: 'PICKED_UP',
      created_at: now,
      confirmed_at: now,
      picked_up_at: now,
      returned_at: '',
      cancelled_at: '',
      refunded_at: '',
      concert_id: concert.concert_id,
      concert_date_id: concertDate.concert_date_id,
      concert_code: concert.concert_code,
      concert_title: concert.concert_title,
      concert_date: concertDate.concert_date,
      concert_time: concertDate.concert_time,
      venue: concert.venue,
      customer_country: String(body.customer_country || 'WALK_IN').trim(),
      country_code: countryCode,
      customer_name: requiredStringFinal_(body.customer_name, 'customer_name'),
      customer_email: String(body.customer_email || '').trim(),
      phone_number: phoneNumber,
      phone_full: phoneFull,
      payment_method: 'CASH',
      base_fee: paidAmount,
      currency: 'KRW',
      payment_status: 'PAID',
      payment_due_at: '',
      paid_at: now,
      paid_amount: paidAmount,
      refund_amount: '',
      refund_method: '',
      refund_note: '',
      expected_suitcase_count: suitcaseCount,
      expected_extra_bag_count: extraBagCount,
      next_day_pickup_required: 'NO',
      next_day_pickup_fee_status: 'NONE',
      checkin_token: '',
      qr_checkin_url: '',
      qr_file_url: '',
      picked_up_by: staff.staff_id || staff.staff_name || 'STAFF',
      agreement_version: CARRYGO_AGREEMENT_VERSION,
      note: String(body.note || '').trim(),
      pickup_time: String(body.pickup_time || 'WALK_IN').trim(),
      booking_channel: 'WALK_IN',
      luggage_tag_numbers: tagNumbers,
      onsite_payment_method: 'CASH',
      onsite_staff: staff.staff_id || staff.staff_name || '',
      onsite_consent_flags: consentFlags,
      actual_suitcase_count: suitcaseCount,
      actual_extra_bag_count: extraBagCount,
      onsite_due_amount: paidAmount,
      onsite_cash_received: String(body.onsite_cash_received || 'NO').trim().toUpperCase(),
      onsite_tag_attached: String(body.onsite_tag_attached || 'NO').trim().toUpperCase(),
      onsite_photo_taken: String(body.onsite_photo_taken || 'NO').trim().toUpperCase(),
      onsite_checkin_completed_at: now
    };

    const sh = getSheetFinal_(CARRYGO_SHEETS.RESERVATIONS);
    sh.appendRow(RESERVATIONS_HEADERS.map(header => rowObject[header] !== undefined ? rowObject[header] : ''));
    applySheetFormatsFinal_(sh, CARRYGO_SHEETS.RESERVATIONS, RESERVATIONS_HEADERS);

    return {
      reservation_id: rowObject.reservation_id,
      status: rowObject.status,
      booking_channel: rowObject.booking_channel,
      payment_status: rowObject.payment_status,
      paid_amount: rowObject.paid_amount,
      amount_display: '₩' + Number(rowObject.paid_amount || 0).toLocaleString('ko-KR'),
      concert_title: rowObject.concert_title,
      concert_date: rowObject.concert_date,
      concert_time: rowObject.concert_time,
      customer_name: rowObject.customer_name,
      phone_full: rowObject.phone_full,
      expected_suitcase_count: rowObject.expected_suitcase_count,
      expected_extra_bag_count: rowObject.expected_extra_bag_count,
      luggage_tag_numbers: rowObject.luggage_tag_numbers
    };
  } finally {
    lock.releaseLock();
  }
}

function normalizeLuggageTagNumberFinal_(value) {
  const digits = String(value || '').replace(/[^0-9]/g, '');
  if (!digits) return 0;
  return Number(digits.slice(-3));
}

function normalizeLuggageTagStringFinal_(value) {
  return String(value || '')
    .split(/[,\s]+/)
    .map(part => String(part || '').trim())
    .filter(Boolean)
    .map(part => {
      const digits = String(part || '').replace(/[^0-9]/g, '');
      return digits ? digits.slice(-3).padStart(3, '0') : part;
    })
    .join(',');
}

function nextLuggageTagNumbersFinal_(count, concertId) {
  const total = Math.max(1, Number(count || 1));
  const targetConcertId = String(concertId || '').trim();
  const rows = readSheetObjectsFinal_(CARRYGO_SHEETS.RESERVATIONS, RESERVATIONS_HEADERS);
  let max = 0;
  rows.forEach(row => {
    if (targetConcertId && String(row.concert_id || '').trim() !== targetConcertId) return;
    String(row.luggage_tag_numbers || '').split(/[,\s]+/).forEach(part => {
      const n = normalizeLuggageTagNumberFinal_(part);
      if (!isNaN(n)) max = Math.max(max, n);
    });
  });
  const tags = [];
  for (let i = 1; i <= total; i++) tags.push(String(max + i).padStart(3, '0'));
  return tags.join(',');
}

function findActiveConcertFinal_(concertId) {
  if (!concertId) throw new Error('concert_id is required');
  const rows = readSheetObjectsFinal_(CARRYGO_SHEETS.CONCERTS, CONCERTS_HEADERS);
  const row = rows.find(item => String(item.concert_id) === String(concertId) && isActiveFinal_(item.is_active));
  if (!row) throw new Error('Active concert not found: ' + concertId);
  return row;
}

function findActiveConcertDateFinal_(concertDateId, concertId) {
  if (!concertDateId) throw new Error('concert_date_id is required');
  const rows = readSheetObjectsFinal_(CARRYGO_SHEETS.CONCERT_DATES, CONCERT_DATES_HEADERS);
  const row = rows.find(item =>
    String(item.concert_date_id) === String(concertDateId) &&
    String(item.concert_id) === String(concertId) &&
    isActiveFinal_(item.is_active)
  );
  if (!row) throw new Error('Active concert date not found: ' + concertDateId);
  return row;
}

function nextReservationIdFinal_(concertCode, concertDate) {
  const dateKey = dateKeyFromConcertDateFinal_(concertDate);
  const prefix = String(concertCode || '').trim().toUpperCase() + '-' + dateKey + '-';
  const rows = readSheetObjectsFinal_(CARRYGO_SHEETS.RESERVATIONS, RESERVATIONS_HEADERS);
  let max = 0;

  rows.forEach(row => {
    const id = String(row.reservation_id || '');
    if (id.indexOf(prefix) === 0) {
      const seq = Number(id.slice(prefix.length));
      if (!isNaN(seq)) max = Math.max(max, seq);
    }
  });

  return prefix + String(max + 1).padStart(3, '0');
}


function nextWalkinReservationIdFinal_(concertDate) {
  const dateKey = dateKeyFromConcertDateFinal_(concertDate);
  const prefix = 'WK-' + dateKey + '-';
  const rows = readSheetObjectsFinal_(CARRYGO_SHEETS.RESERVATIONS, RESERVATIONS_HEADERS);
  let max = 0;
  rows.forEach(row => {
    const id = String(row.reservation_id || '');
    if (id.indexOf(prefix) === 0) {
      const seq = Number(id.slice(prefix.length));
      if (!isNaN(seq)) max = Math.max(max, seq);
    }
  });
  return prefix + String(max + 1).padStart(3, '0');
}

function dateKeyFromConcertDateFinal_(concertDate) {
  const value = String(concertDate || '').trim();
  const match = value.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (match) return match[1].slice(2) + match[2] + match[3];
  const parsed = new Date(value);
  if (!isNaN(parsed.getTime())) return Utilities.formatDate(parsed, CARRYGO_TIMEZONE, 'yyMMdd');
  throw new Error('Invalid concert_date: ' + concertDate);
}

function resolvePaymentFinal_(paymentMethod) {
  const method = String(paymentMethod || '').trim().toUpperCase();
  if (method === 'KAKAOPAY') {
    return { method: 'KAKAOPAY', base_fee: 20000, currency: 'KRW', amount_display: '₩20,000' };
  }
  if (method === 'BANK') {
    return { method: 'BANK', base_fee: 20000, currency: 'KRW', amount_display: '₩20,000' };
  }
  if (method === 'PAYPAL') {
    return { method: 'PAYPAL', base_fee: 15, currency: 'USD', amount_display: '$15' };
  }
  throw new Error('Invalid payment_method: ' + paymentMethod);
}

function buildPaymentInstructionsFinal_(paymentMethod, reservationId) {
  if (paymentMethod === 'KAKAOPAY') {
    return {
      type: 'KAKAOPAY',
      amount_display: '₩20,000',
      memo: reservationId,
      kakaopay_link: 'https://qr.kakaopay.com/FOzMisaMr'
    };
  }

  if (paymentMethod === 'PAYPAL') {
    return {
      type: 'PAYPAL',
      amount_display: '$15',
      payment_note: reservationId,
      paypal_link: ''
    };
  }

  if (paymentMethod === 'BANK') {
    return {
      type: 'BANK',
      amount_display: '₩20,000',
      transfer_memo: reservationId,
      bank_name: '신한은행',
      account_no: getScriptPropertyFinal_('BANK_ACCOUNT_NO'),
      account_holder: getScriptPropertyFinal_('BANK_ACCOUNT_HOLDER')
    };
  }

  return { type: paymentMethod };
}

function validateReservationFinal_(row) {
  if (!row.customer_name) throw new Error('customer_name is required');
  if (!row.customer_email) throw new Error('customer_email is required');
  if (!row.country_code) throw new Error('country_code is required');
  if (!row.phone_number) throw new Error('phone_number is required');
  if (!row.phone_full) throw new Error('phone_full is required');
  if (row.expected_suitcase_count < 1) throw new Error('expected_suitcase_count must be at least 1');
  if (!row.pickup_time) throw new Error('pickup_time is required');
}

function requiredPickupTimeFinal_(value, optionSource) {
  const normalized = String(value || '').trim();
  const allowed = normalizePickupTimeOptionsFinal_(optionSource);
  if (!allowed.includes(normalized)) {
    throw new Error('pickup_time must be one of: ' + allowed.join(', '));
  }
  return normalized;
}

function normalizePickupTimeOptionsFinal_(value) {
  const raw = String(value || '').trim();
  const parts = raw ? raw.split(/[\n,]+/).map(v => v.trim()).filter(Boolean) : ['10:00', '12:00', '14:00'];
  const valid = [];
  parts.forEach(item => {
    const m = item.match(/^([0-2]?\d):([0-5]\d)$/);
    if (!m) return;
    const hh = String(Number(m[1])).padStart(2, '0');
    const time = hh + ':' + m[2];
    if (!valid.includes(time)) valid.push(time);
  });
  return valid.length ? valid : ['10:00', '12:00', '14:00'];
}

function readSheetObjectsFinal_(sheetName, headers) {
  const sh = getSheetFinal_(sheetName);
  const lastRow = sh.getLastRow();
  if (lastRow <= 1) return [];

  const range = sh.getRange(2, 1, lastRow - 1, headers.length);
  const values = range.getValues();
  const displayValues = range.getDisplayValues();
  return values
    .filter(row => row.some(cell => String(cell).trim() !== ''))
    .map((row, rowIndex) => {
      const obj = {};
      headers.forEach((header, index) => {
        obj[header] = normalizeReadCellFinal_(header, row[index], displayValues[rowIndex][index]);
      });
      return obj;
    });
}

function normalizeReadCellFinal_(header, value, displayValue) {
  if (header === 'concert_date') return normalizeConcertDateCellFinal_(value, displayValue);
  if (header === 'concert_time') return normalizeConcertTimeCellFinal_(value, displayValue);
  return value;
}

function normalizeConcertDateCellFinal_(value, displayValue) {
  const display = String(displayValue || '').trim();
  let match = display.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})$/);
  if (match) return match[1] + '-' + match[2].padStart(2, '0') + '-' + match[3].padStart(2, '0');
  match = display.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})$/);
  if (match) return match[3] + '-' + match[1].padStart(2, '0') + '-' + match[2].padStart(2, '0');
  if (value instanceof Date && !isNaN(value.getTime())) return Utilities.formatDate(value, CARRYGO_TIMEZONE, 'yyyy-MM-dd');
  return display || String(value || '').trim();
}

function normalizeConcertTimeCellFinal_(value, displayValue) {
  const display = String(displayValue || '').trim();
  const match = display.match(/(\d{1,2}):(\d{2})/);
  if (match) return match[1].padStart(2, '0') + ':' + match[2];
  if (value instanceof Date && !isNaN(value.getTime())) return Utilities.formatDate(value, CARRYGO_TIMEZONE, 'HH:mm');
  return display || String(value || '').trim();
}

function getSheetFinal_(sheetName) {
  const ss = SpreadsheetApp.openById(CARRYGO_SHEET_ID);
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error('Sheet not found. Run setupCarryGoSheetsFinal first: ' + sheetName);
  return sh;
}

function parseBodyFinal_(e) {
  if (!e || !e.postData || !e.postData.contents) return {};
  return JSON.parse(e.postData.contents);
}

function jsonFinal_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function isActiveFinal_(value) {
  if (value === true) return true;
  return String(value).trim().toUpperCase() === 'TRUE';
}

function requiredStringFinal_(value, fieldName) {
  const text = String(value || '').trim();
  if (!text) throw new Error(fieldName + ' is required');
  return text;
}

function normalizeEmailFinal_(value) {
  const email = String(value || '').trim().toLowerCase();
  if (!email) throw new Error('customer_email is required');
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email)) throw new Error('Invalid customer_email: ' + email);
  return email;
}

function normalizeCountryCodeFinal_(value) {
  const text = String(value || '').trim();
  if (!text) throw new Error('country_code is required');
  const digits = text.replace(/[^0-9]/g, '');
  if (!digits) throw new Error('Invalid country_code: ' + value);
  return '+' + digits;
}

function normalizePhoneFullFinal_(countryCode, phoneNumber) {
  let digits = String(phoneNumber || '').replace(/[^0-9]/g, '');
  if (!digits) return '';

  const codeDigits = String(countryCode || '').replace(/[^0-9]/g, '');
  if (codeDigits === '82' && digits.charAt(0) === '0') digits = digits.slice(1);
  if (digits.indexOf(codeDigits) === 0) return '+' + digits;
  return '+' + codeDigits + digits;
}

function normalizeCountFinal_(value, defaultValue) {
  const n = Number(value === undefined || value === null || value === '' ? defaultValue : value);
  if (isNaN(n) || n < 0) return defaultValue;
  return Math.floor(n);
}

function getScriptPropertyFinal_(key) {
  return PropertiesService.getScriptProperties().getProperty(key) || '';
}

function formatDateTimeFinal_(date) {
  return Utilities.formatDate(date, CARRYGO_TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
}

function formatKoreanDateFinal_(value) {
  const text = String(value || '').trim();
  const match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (!match) return text;
  return match[1] + '년 ' + Number(match[2]) + '월 ' + Number(match[3]) + '일';
}

function formatKoreanTimeFinal_(value) {
  const match = String(value || '').trim().match(/^(\d{1,2}):(\d{2})/);
  if (!match) return String(value || '').trim();
  const hour = Number(match[1]);
  const minute = Number(match[2]);
  const period = hour < 12 ? '오전' : '오후';
  const displayHour = hour % 12 || 12;
  return period + ' ' + displayHour + '시' + (minute ? ' ' + minute + '분' : '');
}

function pickupStartTimeFinal_(value) {
  const match = String(value || '').trim().match(/^(\d{1,2}):(\d{2})/);
  if (!match) return '';
  const date = new Date(2000, 0, 1, Number(match[1]), Number(match[2]));
  date.setMinutes(date.getMinutes() - 30);
  return Utilities.formatDate(date, CARRYGO_TIMEZONE, 'HH:mm');
}

function formatPickupWindowFinal_(value) {
  const pickup = String(value || '').trim();
  if (!pickup) return '';
  const start = pickupStartTimeFinal_(pickup);
  return formatKoreanTimeFinal_(start) + '부터 접수 · ' + formatKoreanTimeFinal_(pickup) + ' 마감';
}

// ===== CarryGo Final Test Helpers =====

/**
 * Creates one sample reservation in the final reservations sheet.
 * Run this from Apps Script after setupCarryGoSheetsFinal() and seedCarryGoSampleDataFinal().
 */
function testCreateReservationFinal() {
  const result = createReservationFinal({
    action: 'create_reservation',
    concert_id: 'shinee_world_vii',
    concert_date_id: 'shinee_20260530',
    customer_country: 'Korea',
    country_code: '+82',
    customer_name: 'TEST CUSTOMER',
    customer_email: 'song.minki@gmail.com',
    phone_number: '010-3345-6625',
    payment_method: 'KAKAOPAY',
    pickup_time: '10:00',
    expected_suitcase_count: 1,
    expected_extra_bag_count: 0,
    note: 'test reservation from Apps Script'
  });

  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

/**
 * Creates one PayPal sample reservation.
 */
function testCreateReservationPayPalFinal() {
  const result = createReservationFinal({
    action: 'create_reservation',
    concert_id: 'shinee_world_vii',
    concert_date_id: 'shinee_20260531',
    customer_country: 'USA',
    country_code: '+1',
    customer_name: 'JANE TEST',
    customer_email: 'jeadee@naver.com',
    phone_number: '555-123-4567',
    payment_method: 'PAYPAL',
    pickup_time: '12:00',
    expected_suitcase_count: 1,
    expected_extra_bag_count: 1,
    note: 'paypal test reservation from Apps Script'
  });

  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

// ===== CarryGo Final Email Helpers =====

function sendPaymentInstructionEmailFinal_(reservation) {
  if (!reservation || !reservation.customer_email) throw new Error('customer_email is required for email');

  const subject = '[CarryGo] 신청이 접수되었습니다 / Payment Required';
  const body = buildPaymentInstructionEmailBodyFinal_(reservation);
  const htmlBody = buildPaymentInstructionEmailHtmlFinal_(reservation);

  MailApp.sendEmail({
    to: reservation.customer_email,
    subject: subject,
    body: body,
    htmlBody: htmlBody,
    name: 'CarryGo'
  });
}

function buildPaymentInstructionEmailBodyFinal_(r) {
  const paymentBlock = buildPaymentInstructionEmailPaymentBlockFinal_(r);
  const amount = formatAmountDisplayFinal_(r);

  return [
    '[예약 확정 전입니다]',
    '',
    '예약을 확정하려면 지금 ' + amount + '을 결제해 주세요.',
    '결제 확인 후 QR 코드와 짐 맡기기/찾기 안내 링크를 이메일로 보내드립니다.',
    '',
    '결제금액: ' + amount,
    '결제방법: ' + paymentBlock.methodLabel,
    '송금메모/메시지: ' + r.reservation_id,
    '',
    paymentBlock.ko,
    '',
    paymentBlock.noticeKo,
    '',
    '예약 정보',
    '- 예약번호: ' + r.reservation_id,
    '- 콘서트: ' + r.concert_title,
    '- 공연일: ' + r.concert_date + ' ' + r.concert_time,
    '- 장소: ' + r.venue,
    '- 짐 맡기는 시간: ' + formatKoreanTimeFinal_(r.pickup_time || '') + ' (30분 전부터 접수, 정시 마감)',
    '',
    '포함 사항: 기본요금 20,000원(캐리어 1개 보관) / 선택 시간 짐 맡기기 / 공연 종료 후 2시간 이내 수령',
    '',
    '신청 후 6시간 이내 결제가 확인되지 않으면 신청이 취소될 수 있습니다.',
    '현장 추가 결제: 추가 캐리어 1개당 ₩20,000, 추가 짐은 가방 1개당 ₩10,000. 한화가 없으면 추가 짐은 $10. 지퍼 및 잠금장치가 있는 가방에 한합니다. 쇼핑백·비닐봉투 등 쉽게 찢어지는 짐은 맡기실 수 없습니다.',
    '결제 후 고객 변심 취소 및 노쇼는 환불되지 않습니다.',
    '',
    'CarryGo',
    '',
    '---',
    '',
    '[Payment Required]',
    '',
    'To confirm your reservation, please pay ' + amount + ' now.',
    'After payment is verified, we will email your QR code and luggage drop-off/pickup guide link.',
    '',
    'Amount: ' + amount,
    'Payment method: ' + paymentBlock.methodLabel,
    'Payment note/message: ' + r.reservation_id,
    '',
    paymentBlock.en,
    '',
    paymentBlock.noticeEn,
    '',
    'Reservation Info',
    '- Reservation ID: ' + r.reservation_id,
    '- Concert: ' + r.concert_title,
    '- Date & Time: ' + r.concert_date + ' ' + r.concert_time,
    '- Venue: ' + r.venue,
    '- Drop-off Time: ' + formatKoreanTimeFinal_(r.pickup_time || '') + ' (check-in opens 30 min before and closes at the selected time)',
    '',
    'CarryGo'
  ].join('\n');
}


function buildPaymentInstructionEmailHtmlFinal_(r) {
  const paymentBlock = buildPaymentInstructionEmailPaymentBlockFinal_(r);
  const amount = formatAmountDisplayFinal_(r);
  const amountText = String(amount || '').replace('₩', '') + (String(amount || '').indexOf('₩') === 0 ? '원' : '');
  const reservationId = escapeHtmlFinal_(r.reservation_id);
  const method = String(r.payment_method || '').toUpperCase();
  const isKakao = method === 'KAKAOPAY';
  const isPaypal = method === 'PAYPAL';
  const notice = isKakao
    ? '<div style="margin-top:14px;padding:13px 14px;border:1.5px solid #b00020;border-radius:12px;background:#fff7f5;color:#8f1d1d;font-size:14px;line-height:1.55;font-weight:800;">카카오페이에서 금액이 자동 입력되지 않을 수 있습니다.<br>송금 화면에서 <strong>20000</strong>을 직접 입력하고, 송금 메모/메시지에 <strong>' + reservationId + '</strong>를 입력해 주세요.</div>'
    : isPaypal
      ? '<div style="margin-top:14px;padding:13px 14px;border:1px solid #ddd;border-radius:12px;background:#fafafa;color:#333;font-size:14px;line-height:1.55;font-weight:750;">PayPal 인보이스를 예약 시 입력한 이메일로 30분 이내 발송합니다.<br>결제 확인 후 QR 코드와 장소 안내를 보내드립니다.</div>'
      : '<div style="margin-top:14px;padding:13px 14px;border:1px solid #ddd;border-radius:12px;background:#fafafa;color:#333;font-size:14px;line-height:1.55;font-weight:750;">입금자명 또는 메모에 예약번호 <strong>' + reservationId + '</strong>를 입력해 주세요.</div>';
  const memoLabel = isPaypal ? '인보이스 기준 예약번호' : '송금메모/메시지';

  return [
    '<div style="margin:0;padding:0;background:#f7f2ea;font-family:Arial,\'Apple SD Gothic Neo\',\'Noto Sans KR\',sans-serif;color:#111;">',
    '<div style="max-width:560px;margin:0 auto;padding:18px;">',
    '<div style="background:#fff;border:1.5px solid #111;border-radius:18px;padding:20px;">',
    '<div style="font-size:13px;font-weight:900;letter-spacing:.06em;color:#8f1d1d;margin-bottom:10px;">신청이 접수되었습니다</div>',
    '<div style="font-size:24px;line-height:1.28;font-weight:900;letter-spacing:-.035em;margin-bottom:10px;">예약 확정을 위해<br>기본 이용료를 결제해 주세요.</div>',
    '<div style="font-size:15px;line-height:1.55;color:#444;font-weight:750;margin-bottom:18px;">결제 확인 후 QR 코드와 짐 맡기기/찾기 안내 링크를 이메일로 보내드립니다.</div>',
    '<div style="border:1.5px solid #111;border-radius:14px;padding:15px;background:#fff;margin-bottom:14px;">',
    '<div style="font-size:12px;font-weight:900;color:#777;letter-spacing:.12em;text-transform:uppercase;margin-bottom:8px;">Payment</div>',
    '<div style="font-size:17px;line-height:1.7;font-weight:850;">',
    '결제금액: <span style="font-size:24px;font-weight:950;">' + escapeHtmlFinal_(amountText) + '</span><br>',
    '결제방법: ' + escapeHtmlFinal_(paymentBlock.methodLabel) + '<br>',
    escapeHtmlFinal_(memoLabel) + ': <span style="font-size:18px;font-weight:950;">' + reservationId + '</span>',
    '</div>',
    '</div>',
    notice,
    paymentBlock.link ? '<a href="' + escapeHtmlFinal_(paymentBlock.link || '') + '" style="display:block;text-align:center;background:#111;color:#fff;text-decoration:none;border-radius:999px;padding:15px 14px;font-size:17px;font-weight:900;margin:16px 0;">' + escapeHtmlFinal_(paymentBlock.buttonLabel || '결제하기') + '</a>' : '',
    '<div style="margin-top:20px;border-top:1px solid #ddd;padding-top:15px;">',
    '<div style="font-size:15px;font-weight:900;margin-bottom:8px;">예약 정보</div>',
    '<div style="font-size:14px;line-height:1.65;color:#333;font-weight:700;">',
    '예약번호: ' + reservationId + '<br>',
    '콘서트: ' + escapeHtmlFinal_(r.concert_title) + '<br>',
    '공연일: ' + escapeHtmlFinal_(r.concert_date + ' ' + r.concert_time) + '<br>',
    '장소: ' + escapeHtmlFinal_(r.venue) + '<br>',
    '짐 맡기는 시간: ' + escapeHtmlFinal_(formatKoreanTimeFinal_(r.pickup_time || '')) + ' <span style="color:#777;">(30분 전부터 접수, 정시 마감)</span>',
    '</div>',
    '</div>',
    '<div style="margin-top:15px;font-size:12px;line-height:1.55;color:#666;font-weight:650;">포함 사항: 기본요금 20,000원(캐리어 1개 보관) / 선택 시간 짐 맡기기 / 공연 종료 후 2시간 이내 수령<br>현장 추가 결제: 추가 캐리어 1개당 20,000원, 추가 짐은 가방 1개당 10,000원. 한화가 없으면 추가 짐은 $10. 지퍼 및 잠금장치가 있는 가방에 한합니다. 쇼핑백·비닐봉투 등 쉽게 찢어지는 짐은 맡기실 수 없습니다.</div>',
    '</div>',
    '<div style="font-size:12px;color:#777;line-height:1.5;margin-top:14px;text-align:center;">CarryGo</div>',
    '</div>',
    '</div>'
  ].join('');
}

function buildPaymentInstructionEmailPaymentBlockFinal_(r) {
  const method = String(r.payment_method || '').toUpperCase();
  const reservationId = r.reservation_id;

  if (method === 'KAKAOPAY') {
    const link = 'https://qr.kakaopay.com/FOzMisaMr';
    return {
      methodLabel: 'KakaoPay',
      link: link,
      buttonLabel: '카카오페이 송금하기',
      noticeKo: '※ 카카오페이는 링크를 눌러도 금액이 자동 입력되지 않을 수 있습니다. 송금 화면에서 반드시 ₩20,000을 직접 입력하고, 송금 메모에 예약번호 ' + reservationId + '를 입력해 주세요.',
      noticeEn: '※ KakaoPay may not auto-fill the amount. Please enter ₩20,000 manually and add your Reservation ID ' + reservationId + ' in the transfer memo.',
      ko: [
        '카카오페이 송금 링크:',
        link
      ].join('\n'),
      en: [
        'KakaoPay link:',
        link
      ].join('\n')
    };
  }

  if (method === 'PAYPAL') {
    return {
      methodLabel: 'PayPal',
      link: '',
      buttonLabel: '',
      noticeKo: '※ PayPal 인보이스를 예약 시 입력한 이메일로 30분 이내 발송합니다. 결제 확인 후 QR과 장소 안내를 보내드립니다.',
      noticeEn: '※ We will send a PayPal invoice to your email within 30 minutes. Your QR code and pickup guide will be sent after payment is confirmed.',
      ko: [
        'PayPal 인보이스를 이메일로 보내드립니다.',
        '인보이스 발송 기준 예약번호: ' + reservationId
      ].join('\n'),
      en: [
        'A PayPal invoice will be sent to your email.',
        'Reservation ID: ' + reservationId
      ].join('\n')
    };
  }

  if (method === 'BANK') {
    const accountNo = getScriptPropertyFinal_('BANK_ACCOUNT_NO') || '{{bank_account}}';
    const holder = getScriptPropertyFinal_('BANK_ACCOUNT_HOLDER') || '{{account_holder}}';
    return {
      methodLabel: 'Bank Transfer',
      link: '',
      noticeKo: '※ 송금 메모에 예약번호 ' + reservationId + '를 입력해 주세요.',
      noticeEn: '※ Please add your Reservation ID ' + reservationId + ' in the transfer memo.',
      ko: [
        '계좌이체 / Bank Transfer',
        '결제금액: ₩20,000',
        '송금 메모: ' + reservationId,
        '은행: 신한은행',
        '계좌번호: ' + accountNo,
        '예금주: ' + holder
      ].join('\n'),
      en: [
        'Bank Transfer',
        'Amount: ₩20,000',
        'Transfer memo: ' + reservationId,
        'Bank: SHINHAN BANK',
        'Account No.: ' + accountNo,
        'Account Holder: ' + holder,
        '',
        'If you need international bank transfer information, please contact CarryGo.'
      ].join('\n')
    };
  }

  return { methodLabel: method, link: '', buttonLabel: '결제하기', noticeKo: '', noticeEn: '', ko: method, en: method };
}

function formatAmountDisplayFinal_(r) {
  if (String(r.currency).toUpperCase() === 'USD') return '$' + String(r.base_fee);
  return '₩' + Number(r.base_fee || 0).toLocaleString('ko-KR');
}

function testSendPaymentInstructionEmailFinal() {
  const result = createReservationFinal({
    action: 'create_reservation',
    concert_id: 'shinee_world_vii',
    concert_date_id: 'shinee_20260530',
    customer_country: 'Korea',
    country_code: '+82',
    customer_name: 'EMAIL TEST CUSTOMER',
    customer_email: 'song.minki@gmail.com',
    phone_number: '010-3345-6625',
    payment_method: 'KAKAOPAY',
    pickup_time: '10:00',
    expected_suitcase_count: 1,
    expected_extra_bag_count: 0,
    note: 'email test reservation from Apps Script'
  });

  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

// ===== CarryGo Final Payment Confirmation / QR Email =====

function confirmPaymentFinal(reservationId, paidAmount) {
  if (!reservationId) throw new Error('reservation_id is required');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const sh = getSheetFinal_(CARRYGO_SHEETS.RESERVATIONS);
    const rowNo = findReservationRowFinal_(reservationId);
    const row = getReservationObjectByRowFinal_(rowNo);
    const now = new Date();

    if (String(row.status || '') === 'CANCELLED') throw new Error('Cannot confirm a cancelled reservation: ' + reservationId);
    if (String(row.payment_status || '') === 'PAID' && String(row.status || '') === 'CONFIRMED') {
      return row;
    }

    const token = row.checkin_token || generateTokenFinal_();
    const checkinUrl = buildCheckinUrlFinal_(reservationId, token);
    const qrBlob = createQrPngBlobFinal_(checkinUrl, reservationId);

    setReservationValueFinal_(sh, rowNo, 'status', 'CONFIRMED');
    setReservationValueFinal_(sh, rowNo, 'confirmed_at', now);
    setReservationValueFinal_(sh, rowNo, 'payment_status', 'PAID');
    setReservationValueFinal_(sh, rowNo, 'paid_at', now);
    setReservationValueFinal_(sh, rowNo, 'paid_amount', paidAmount || row.base_fee || '');
    setReservationValueFinal_(sh, rowNo, 'checkin_token', token);
    setReservationValueFinal_(sh, rowNo, 'qr_checkin_url', checkinUrl);

    const updated = getReservationObjectByRowFinal_(rowNo);
    sendReservationConfirmedEmailFinal_(updated, qrBlob);
    return updated;
  } finally {
    lock.releaseLock();
  }
}

function sendReservationConfirmedEmailFinal_(reservation, qrBlob) {
  if (!reservation || !reservation.customer_email) throw new Error('customer_email is required for confirmation email');

  MailApp.sendEmail({
    to: reservation.customer_email,
    subject: '[CarryGo] 예약이 확정되었습니다 / Reservation Confirmed',
    body: buildReservationConfirmedEmailBodyFinal_(reservation),
    htmlBody: buildReservationConfirmedEmailHtmlFinal_(reservation),
    name: 'CarryGo',
    attachments: [qrBlob],
    inlineImages: {
      qrCode: qrBlob
    }
  });
}

function buildReservationConfirmedEmailBodyFinal_(r) {
  return [
    '안녕하세요. CarryGo입니다.',
    '',
    '예약이 확정되었습니다.',
    '현장에서 이 이메일에 첨부된 QR 코드를 스태프에게 보여주세요.',
    '',
    'QR이 보이지 않으면 아래 링크 또는 예약번호를 스태프에게 보여주세요. 고객이 링크 화면에서 직접 입력하거나 버튼을 누를 필요는 없습니다.',
    r.qr_checkin_url || '',
    '',
    '예약번호: ' + r.reservation_id,
    '콘서트: ' + r.concert_title,
    '공연일: ' + r.concert_date + ' ' + r.concert_time,
    '장소: ' + r.venue,
    '짐 맡기는 시간: ' + formatKoreanTimeFinal_(r.pickup_time || '') + ' (30분 전부터 접수, 정시 마감)',
    '결제상태: 결제 완료',
    '',
    '짐 맡기기/찾기 안내:',
    getConcertDateLinkFinal_(r.concert_date_id, 'pickup_drop_guide_link'),
    '',
    '기본 이용료 포함 사항:',
    '- 기본요금 20,000원에는 캐리어 1개 보관이 포함됩니다.',
    '- 선택한 시간에 짐 맡기기',
    '- 공연 종료 후 2시간 이내 수령',
    '- 캐리어 크기와 상관없이 1개 기준',
    '',
    '현장 추가 결제:',
    '- 추가 캐리어: 1개당 ₩20,000',
    '- 추가 짐: 가방 1개당 ₩10,000 / 한화가 없으면 $10 현금 가능',
    '- 추가 짐은 지퍼 및 잠금장치가 있는 가방에 한합니다. 쇼핑백·비닐봉투 등 쉽게 찢어지는 짐은 맡기실 수 없습니다.',
    '- 현장 환율 계산 없음',
    '',
    '짐 맡길 때 안내:',
    '- 선택한 시간 30분 전부터 접수 가능하며, 정시에 해당 시간대 접수가 마감됩니다. 늦으면 노쇼 처리될 수 있습니다.',
    '- 현장에서 예약 QR 또는 예약번호를 스태프에게 보여주세요.',
    '- CarryGo 스태프가 러기지택 / Luggage Tag를 짐에 부착합니다.',
    '- 고객용 러기지택은 공연 후 수령 시 필요하니 잃어버리지 말고 보관해 주세요.',
    '',
    '공연 후 수령:',
    '- 공연 종료 후 2시간 이내에 짐을 수령해 주세요.',
    '- 고객용 러기지택 번호와 짐 부착용 태그 번호를 확인한 후 짐을 드립니다.',
    '',
    '당일 미수령 안내:',
    '공연 종료 후 2시간 이내 수령이 원칙입니다.',
    '- 시간 내 수령하지 못한 경우 추가 보관 및 운영 비용으로 예약 1건당 30,000원이 부과됩니다.',
    '',
    '취소 및 환불:',
    '- 결제 후 고객 변심 취소 및 노쇼는 환불되지 않습니다.',
    '- 공연 취소 또는 CarryGo 운영 사정으로 서비스 제공이 불가한 경우 전액 환불됩니다.',
    '',
    '감사합니다.',
    'CarryGo',
    '',
    '---',
    '',
    'Hello, this is CarryGo.',
    '',
    'Your reservation has been confirmed.',
    'Please show the QR code attached to this email to CarryGo staff onsite.',
    '',
    'If the QR code is not visible, show the link below or your Reservation ID to staff. Customers do not need to enter anything or press any staff buttons on the link screen.',
    r.qr_checkin_url || '',
    '',
    'Reservation ID: ' + r.reservation_id,
    'Concert: ' + r.concert_title,
    'Date & Time: ' + r.concert_date + ' ' + r.concert_time,
    'Venue: ' + r.venue,
    'Drop-off Time: ' + formatKoreanTimeFinal_(r.pickup_time || '') + ' (check-in opens 30 min before and closes at the selected time)',
    'Payment Status: Paid',
    '',
    'Luggage Drop-off/Pickup Guide:',
    getConcertDateLinkFinal_(r.concert_date_id, 'pickup_drop_guide_link'),
    '',
    'Base Fee includes:',
    '- Base fee includes storage for 1 suitcase',
    '- Luggage drop-off at your selected time',
    '- Pickup within 2 hours after the concert ends',
    '- One suitcase regardless of size',
    '',
    'Onsite extra charges:',
    '- Additional suitcase: ₩20,000 each',
    '- Extra bag: ₩10,000 each / $10 cash if you do not have KRW',
    '- We do not calculate exchange rates onsite',
    '',
    'Drop-off instructions:',
    '- Check-in opens 30 minutes before your selected time and closes at the selected time. Late arrival may be treated as a no-show.',
    '- Please show your reservation QR code or Reservation ID to CarryGo staff onsite.',
    '- CarryGo staff will attach a Luggage Tag to your luggage.',
    '- Please keep your customer Luggage Tag until after-concert pickup.',
    '',
    'After-concert pickup:',
    '- Please pick up your luggage within 2 hours after the concert ends.',
    '- We will match your customer Luggage Tag number with the luggage tag before returning your luggage.',
    '',
    'Late pickup:',
    'Please pick up your luggage within 2 hours after the concert ends.',
    '- If luggage is not picked up within this time, an additional late pickup/storage fee of KRW 30,000 per booking will apply.',
    '',
    'Cancellation & Refund:',
    '- Customer cancellations and no-shows are non-refundable after payment.',
    '- If the concert is cancelled or CarryGo cannot provide the service due to operational reasons, a full refund will be provided.',
    '',
    'Thank you.',
    'CarryGo'
  ].join('\n');
}

function buildCheckinUrlFinal_(reservationId, token) {
  const adminUrl = getScriptPropertyFinal_('ADMIN_CHECKIN_URL') || 'https://songminki-cloud.github.io/carrygo/admin.html';
  return adminUrl + '?mode=checkin&id=' + encodeURIComponent(reservationId) + '&token=' + encodeURIComponent(token);
}

function createQrPngBlobFinal_(text, reservationId) {
  const encoded = encodeURIComponent(text);
  const url = 'https://quickchart.io/qr?size=900&margin=2&text=' + encoded;
  const response = UrlFetchApp.fetch(url);
  return response.getBlob().setName('CarryGo_QR_' + reservationId + '.png');
}

function generateTokenFinal_() {
  return Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
}

function findReservationRowFinal_(reservationId) {
  const sh = getSheetFinal_(CARRYGO_SHEETS.RESERVATIONS);
  const values = sh.getDataRange().getValues();
  const col = RESERVATIONS_HEADERS.indexOf('reservation_id');
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][col]) === String(reservationId)) return i + 1;
  }
  throw new Error('reservation_id not found: ' + reservationId);
}

function getReservationObjectByRowFinal_(rowNo) {
  const sh = getSheetFinal_(CARRYGO_SHEETS.RESERVATIONS);
  const values = sh.getRange(rowNo, 1, 1, RESERVATIONS_HEADERS.length).getValues()[0];
  const obj = {};
  RESERVATIONS_HEADERS.forEach((header, index) => obj[header] = values[index]);
  return obj;
}

function setReservationValueFinal_(sh, rowNo, header, value) {
  const col = RESERVATIONS_HEADERS.indexOf(header) + 1;
  if (col < 1) throw new Error('Reservation header not found: ' + header);
  sh.getRange(rowNo, col).setValue(value);
}

function getConcertDateLinkFinal_(concertDateId, field) {
  const rows = readSheetObjectsFinal_(CARRYGO_SHEETS.CONCERT_DATES, CONCERT_DATES_HEADERS);
  const row = rows.find(item => String(item.concert_date_id) === String(concertDateId));
  return row ? String(row[field] || '') : '';
}

function testConfirmLatestReservationFinal() {
  const rows = readSheetObjectsFinal_(CARRYGO_SHEETS.RESERVATIONS, RESERVATIONS_HEADERS);
  const unpaid = rows.filter(row => String(row.status) === 'UNPAID');
  if (!unpaid.length) throw new Error('No UNPAID reservation found for confirmation test');
  const latest = unpaid[unpaid.length - 1];
  const result = confirmPaymentFinal(latest.reservation_id, latest.base_fee);
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}


// ===== CarryGo Final Check-in Page / HTML Email =====

function buildReservationConfirmedEmailHtmlFinal_(r) {
  const pickupLink = escapeHtmlFinal_(getConcertDateLinkFinal_(r.concert_date_id, 'pickup_drop_guide_link'));
  const qrUrl = escapeHtmlFinal_(r.qr_checkin_url || '');

  return `
  <div style="font-family:Arial,Helvetica,sans-serif;max-width:640px;margin:0 auto;color:#111;line-height:1.5;">
    <h2 style="margin:0 0 12px;">CarryGo Reservation Confirmed</h2>
    <p style="font-size:16px;margin:0 0 16px;">예약이 확정되었습니다. 현장에서 아래 QR 코드를 스태프에게 보여주세요.</p>
    <p style="font-size:15px;margin:0 0 20px;">Your reservation has been confirmed. Please show the QR code below to CarryGo staff onsite.</p>

    <div style="text-align:center;margin:24px 0;padding:20px;border:2px solid #111;border-radius:16px;">
      <img src="cid:qrCode" alt="CarryGo QR Code" style="width:280px;max-width:90%;height:auto;display:block;margin:0 auto;" />
      <div style="font-size:20px;font-weight:bold;margin-top:12px;">${escapeHtmlFinal_(r.reservation_id)}</div>
    </div>

    <p style="font-size:14px;">QR이 보이지 않으면 아래 링크 또는 예약번호를 스태프에게 보여주세요. 고객이 링크 화면에서 직접 입력하거나 버튼을 누를 필요는 없습니다.<br/>If the QR code is not visible, show the link below or your Reservation ID to staff. Customers do not need to enter anything on the link screen.</p>
    <p><a href="${qrUrl}" style="color:#111;word-break:break-all;">${qrUrl}</a></p>

    <hr style="border:none;border-top:1px solid #ddd;margin:24px 0;" />

    <table style="width:100%;border-collapse:collapse;font-size:14px;">
      <tr><td style="padding:6px 0;font-weight:bold;">Reservation ID</td><td>${escapeHtmlFinal_(r.reservation_id)}</td></tr>
      <tr><td style="padding:6px 0;font-weight:bold;">Concert</td><td>${escapeHtmlFinal_(r.concert_title)}</td></tr>
      <tr><td style="padding:6px 0;font-weight:bold;">Date & Time</td><td>${escapeHtmlFinal_(r.concert_date)} ${escapeHtmlFinal_(r.concert_time)}</td></tr>
      <tr><td style="padding:6px 0;font-weight:bold;">Venue</td><td>${escapeHtmlFinal_(r.venue)}</td></tr>
      <tr><td style="padding:6px 0;font-weight:bold;">Drop-off Time</td><td>${escapeHtmlFinal_(formatKoreanTimeFinal_(r.pickup_time || ''))} · 30분 전부터 접수, 정시 마감</td></tr>
      <tr><td style="padding:6px 0;font-weight:bold;">Payment</td><td>Paid / 결제 완료</td></tr>
    </table>

    <h3 style="margin-top:24px;">Luggage Drop-off/Pickup Guide / 짐 맡기기·찾기 안내</h3>
    <p><a href="${pickupLink}" style="color:#111;word-break:break-all;">${pickupLink}</a></p>

    <h3 style="margin-top:24px;">Important / 중요 안내</h3>
    <ul style="padding-left:20px;">
      <li>Base fee includes storage for 1 suitcase, drop-off at the selected time, and pickup within 2 hours after the concert ends.</li>
      <li>기본요금 20,000원에는 캐리어 1개 보관, 선택 시간 짐 맡기기, 공연 종료 후 2시간 이내 수령이 포함됩니다.</li>
      <li>Additional suitcase: ₩20,000 each. Extra bag: ₩10,000 each. If you do not have KRW, $10 cash per extra bag is accepted. Shopping bags, plastic bags, or other easily torn items are not accepted. No exchange-rate calculation onsite.</li>
      <li>추가 캐리어: 1개당 ₩20,000. 추가 짐: 가방 1개당 ₩10,000. 한화가 없으면 추가 짐은 $10 현금 가능. 지퍼 및 잠금장치가 있는 가방에 한합니다. 쇼핑백·비닐봉투 등 쉽게 찢어지는 짐은 맡기실 수 없습니다. 현장 환율 계산 없음.</li>
      <li>Late pickup: If luggage is not picked up within 2 hours after the concert ends, an additional late pickup/storage fee of KRW 30,000 per booking will apply.</li>
      <li>당일 미수령: 공연 종료 후 2시간 이내 수령하지 못한 경우 추가 보관 및 운영 비용으로 예약 1건당 30,000원이 부과됩니다.</li>
      <li>Customer cancellations and no-shows are non-refundable after payment.</li>
      <li>결제 후 고객 변심 취소 및 노쇼는 환불되지 않습니다.</li>
    </ul>

    <p style="margin-top:28px;">Thank you.<br/>CarryGo</p>
  </div>`;
}

function renderCheckinPageFinal_(params) {
  const reservationId = String(params.id || '').trim();
  const token = String(params.token || '').trim();

  try {
    if (!reservationId || !token) throw new Error('Missing QR information.');
    const rowNo = findReservationRowFinal_(reservationId);
    const r = getReservationObjectByRowFinal_(rowNo);
    if (String(r.checkin_token || '') !== token) throw new Error('Invalid or expired QR code.');

    const pickupLink = escapeHtmlFinal_(getConcertDateLinkFinal_(r.concert_date_id, 'pickup_drop_guide_link'));
      const maskedName = maskNameFinal_(r.customer_name);

    const html = `
      <!doctype html>
      <html>
      <head>
        <meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no, viewport-fit=cover" />
        <title>CarryGo Reservation</title>
        <base target="_top">
        <style>
          body{font-family:Arial,Helvetica,sans-serif;margin:0;background:#f6f6f6;color:#111;}
          html,body{width:100%;min-width:0;margin:0;padding:0;overflow-x:hidden;-webkit-text-size-adjust:100%;}
          body{font-family:Arial,Helvetica,sans-serif;background:#f6f6f6;color:#111;}
          .wrap{width:100%;max-width:none;margin:0;padding:clamp(10px,2.8vw,22px);box-sizing:border-box;}
          .card{width:100%;box-sizing:border-box;background:#fff;border-radius:18px;padding:clamp(18px,4vw,34px);box-shadow:0 4px 18px rgba(0,0,0,.08);}
          .logoImage{display:block;width:clamp(150px,30vw,240px);max-width:58%;height:auto;margin:0 0 clamp(14px,3vw,24px);border:0;border-radius:0;}
          .logoSmall{display:inline-block;width:clamp(150px,30vw,260px);max-width:62%;border:2px solid #111;border-radius:8px;overflow:hidden;margin:0 0 clamp(14px,3vw,24px);background:#fff;vertical-align:top;}
          .logoSmallTop{display:flex;background:#050505;color:#fff;height:clamp(32px,6vw,54px);}
          .logoSmallName{flex:1;display:flex;align-items:center;justify-content:center;font-size:clamp(24px,4.8vw,44px);font-weight:900;letter-spacing:.06em;line-height:1;}
          .logoSmallSide{width:clamp(22px,4vw,36px);background:#fff;color:#111;display:flex;align-items:center;justify-content:center;border-left:2px solid #111;font-size:clamp(9px,1.8vw,15px);font-weight:900;letter-spacing:.13em;writing-mode:vertical-rl;}
          .logoSmallBottom{text-align:center;font-size:clamp(8px,1.7vw,14px);font-weight:900;letter-spacing:.30em;padding:clamp(5px,1.1vw,10px) 2px;border-top:2px solid #111;white-space:nowrap;}
          .brand{font-size:24px;font-weight:800;letter-spacing:.04em;margin-bottom:4px;}
          .sub{font-size:13px;color:#555;margin-bottom:20px;}
          .status{display:block;width:max-content;max-width:100%;background:#111;color:#fff;border-radius:999px;padding:clamp(8px,1.7vw,14px) clamp(13px,2.5vw,22px);font-size:clamp(16px,3.1vw,28px);font-weight:900;margin:0 0 clamp(18px,3vw,28px);line-height:1.18;}
          .row{border-top:1px solid #eee;padding:clamp(15px,3vw,28px) 0;}
          .label{font-size:clamp(15px,2.5vw,23px);color:#666;margin-bottom:clamp(5px,1vw,10px);}
          .value{font-size:clamp(24px,5.2vw,48px);font-weight:900;word-break:break-word;line-height:1.12;}.subnote{margin-top:8px;font-size:clamp(14px,2.5vw,22px);font-weight:800;color:#666;line-height:1.35;}
          a.button{display:block;text-align:center;background:#111;color:#fff;text-decoration:none;border-radius:12px;padding:clamp(17px,3.2vw,30px) 12px;margin-top:clamp(12px,2vw,20px);font-size:clamp(18px,3.4vw,32px);font-weight:900;}
          .note{font-size:clamp(15px,2.5vw,23px);color:#555;line-height:1.45;margin-top:clamp(18px,3vw,28px);}
          @media (max-width:480px){.wrap{padding:10px}.card{padding:20px}.logoImage{width:168px;max-width:60%;}.value{font-size:24px}.label{font-size:15px}.status{font-size:16px}a.button{font-size:18px}}
        </style>
      </head>
      <body>
        <div class="wrap">
          <div class="card">
            <img class="logoImage" src="data:image/jpeg;base64,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" alt="CarryGo Concert Luggage Care Seoul">
            <div class="row"><div class="label">Status / 상태</div><div class="value">${escapeHtmlFinal_(r.status || 'CONFIRMED')}</div></div>
            <div class="row"><div class="label">Reservation ID / 예약번호</div><div class="value">${escapeHtmlFinal_(r.reservation_id)}</div></div>
            <div class="row"><div class="label">Name / 이름</div><div class="value">${escapeHtmlFinal_(maskedName)}</div></div>
            <div class="row"><div class="label">Concert / 콘서트</div><div class="value">${escapeHtmlFinal_(r.concert_title)}</div></div>
            <div class="row"><div class="label">Date & Time / 공연일시</div><div class="value">${escapeHtmlFinal_(formatKoreanDateFinal_(r.concert_date))}<br>${escapeHtmlFinal_(formatKoreanTimeFinal_(r.concert_time))}</div></div>
            <div class="row"><div class="label">Venue / 장소</div><div class="value">${escapeHtmlFinal_(r.venue)}</div></div>
            <div class="row"><div class="label">Drop-off Time / 짐 맡기는 시간</div><div class="value">${escapeHtmlFinal_(formatKoreanTimeFinal_(r.pickup_time || ''))}</div><div class="subnote">${escapeHtmlFinal_(formatPickupWindowFinal_(r.pickup_time || ''))}</div></div>

            <a class="button" target="_top" href="${pickupLink}">Luggage Drop-off/Pickup Guide / 짐 맡기기·찾기 안내</a>
            ${buildStaffActionHtmlFinal_(r, token, params)}

            <div class="note">
              Please show this page or your QR code to CarryGo staff onsite.<br/>
              현장에서 이 화면 또는 QR 코드를 CarryGo 스태프에게 보여주세요.
            </div>
          </div>
        </div>
      </body>
      </html>`;

    return HtmlService.createHtmlOutput(html)
      .setTitle('CarryGo Reservation')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    const message = escapeHtmlFinal_(String(err && err.message ? err.message : err));
    return HtmlService.createHtmlOutput(`
      <div style="font-family:Arial,sans-serif;padding:24px;max-width:520px;margin:0 auto;">
        <h2>CarryGo</h2>
        <p>QR 확인 중 문제가 발생했습니다.</p>
        <p>There was a problem verifying this QR code.</p>
        <pre style="white-space:pre-wrap;background:#f5f5f5;padding:12px;border-radius:8px;">${message}</pre>
      </div>`).setTitle('CarryGo QR Error');
  }
}

function maskNameFinal_(name) {
  const text = String(name || '').trim();
  if (!text) return '';
  if (text.length <= 2) return text.charAt(0) + '*';
  return text.slice(0, 2) + '*'.repeat(Math.min(text.length - 2, 6));
}

function escapeHtmlFinal_(value) {
  return String(value === undefined || value === null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}


// ===== CarryGo Final Staff Login / Pickup Complete =====

function renderStaffLoginPageFinal_(params) {
  const reservationId = String(params.id || '').trim();
  const token = String(params.token || '').trim();
  const error = String(params.error || '').trim();
  const staffCode = String(params.staff_code || '').trim();

  if (staffCode) {
    const staff = findStaffByCodeFinal_(staffCode);
    if (staff) {
      const sessionToken = createStaffSessionFinal_(staff);
      return renderCheckinPageFinal_({
        mode: 'checkin',
        id: reservationId,
        token: token,
        staff_session: sessionToken
      });
    }
    return renderStaffLoginPageFinal_({
      mode: 'staff_login',
      id: reservationId,
      token: token,
      error: 'Invalid staff code.'
    });
  }

  const checkinHidden = reservationId && token
    ? `<input type="hidden" name="id" value="${escapeHtmlFinal_(reservationId)}"><input type="hidden" name="token" value="${escapeHtmlFinal_(token)}">`
    : '';
  const errorHtml = error ? `<div class="err">${escapeHtmlFinal_(error)}</div>` : '';

  const html = `
    <!doctype html>
    <html>
    <head>
      <meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no, viewport-fit=cover" />
      <title>CarryGo Staff Login</title>
      <base target="_top">
      <style>
        body{font-family:Arial,Helvetica,sans-serif;margin:0;background:#f6f6f6;color:#111;}
        .wrap{max-width:480px;margin:0 auto;padding:22px;}
        .card{background:#fff;border-radius:18px;padding:24px;box-shadow:0 4px 18px rgba(0,0,0,.08);}
        h1{font-size:26px;margin:0 0 8px;}
        p{color:#555;line-height:1.45;}
        input{box-sizing:border-box;width:100%;font-size:20px;padding:14px;border:1px solid #ccc;border-radius:12px;margin-top:12px;}
        button{width:100%;font-size:18px;font-weight:800;background:#111;color:#fff;border:0;border-radius:12px;padding:14px;margin-top:14px;}
        .err{background:#fff0f0;color:#b00020;border:1px solid #ffc8c8;border-radius:12px;padding:12px;margin-bottom:12px;font-size:14px;}
      </style>
    </head>
    <body>
      <div class="wrap">
        <div class="card">
          <h1>CarryGo Staff Login</h1>
          <p>스태프 코드를 입력해 주세요.<br/>Please enter your staff code.</p>
          ${errorHtml}
          <form method="get" target="_top">
            <input type="hidden" name="mode" value="staff_login">
            ${checkinHidden}
            <input name="staff_code" placeholder="Staff code" autocomplete="off" autofocus>
            <button type="submit">Login / 로그인</button>
          </form>
        </div>
      </div>
    </body>
    </html>`;

  return HtmlService.createHtmlOutput(html).setTitle('CarryGo Staff Login');
}

function renderStaffLogoutPageFinal_() {
  return HtmlService.createHtmlOutput(`
    <div style="font-family:Arial,sans-serif;padding:24px;max-width:520px;margin:0 auto;">
      <h2>CarryGo Staff Logout</h2>
      <p>로그아웃하려면 브라우저 탭을 닫거나 새 스태프 로그인으로 덮어써 주세요.</p>
      <p>Please close this browser tab or log in again with another staff code.</p>
    </div>`).setTitle('CarryGo Staff Logout');
}

function handlePickupCompleteFinal_(params) {
  return HtmlService.createHtmlOutput(`
    <div style="font-family:Arial,sans-serif;padding:24px;max-width:520px;margin:0 auto;">
      <h2>CarryGo Onsite Check-in Required</h2>
      <p>수량 확인, 현장 추가금, 러기지택 발급 후 접수완료 처리해야 합니다.</p>
      <p>Please use the QR check-in screen to confirm quantities, cash due, and luggage tags.</p>
    </div>`).setTitle('CarryGo Onsite Check-in Required');
}

function findStaffByCodeFinal_(staffCode) {
  const rows = readSheetObjectsFinal_(CARRYGO_SHEETS.STAFF, STAFF_HEADERS);
  const normalized = normalizeStaffCodeFinal_(staffCode);
  return rows.find(row => normalizeStaffCodeFinal_(row.staff_code) === normalized && isActiveFinal_(row.is_active));
}

function normalizeStaffCodeFinal_(value) {
  const raw = String(value === undefined || value === null ? '' : value).trim();
  const digits = raw.replace(/[^0-9]/g, '');
  if (!digits) return raw.toUpperCase();
  return String(Number(digits));
}

function createStaffSessionFinal_(staff) {
  const token = generateTokenFinal_();
  const expiresAt = new Date(Date.now() + 3 * 24 * 60 * 60 * 1000);
  const payload = {
    staff_id: staff.staff_id,
    staff_name: staff.staff_name,
    expires_at: expiresAt.getTime()
  };

  // CacheService cannot reliably keep a 3-day session; Apps Script cache max TTL is limited.
  // Use Script Properties for the MVP staff session instead.
  PropertiesService.getScriptProperties().setProperty('staff_session_' + token, JSON.stringify(payload));
  return token;
}

function validateStaffSessionFinal_(staffSession) {
  if (!staffSession) return null;
  const key = 'staff_session_' + staffSession;
  const props = PropertiesService.getScriptProperties();
  const raw = props.getProperty(key);
  if (!raw) return null;

  const payload = JSON.parse(raw);
  if (!payload.expires_at || Number(payload.expires_at) < Date.now()) {
    props.deleteProperty(key);
    return null;
  }
  return payload;
}

function redirectStaffLoginWithErrorFinal_(reservationId, token, error) {
  let url = buildStaffLoginUrlFinal_(reservationId, token);
  url += '&error=' + encodeURIComponent(error);
  return HtmlService.createHtmlOutput(`<script>window.location.replace(${JSON.stringify(url)});</script>`)
    .setTitle('CarryGo Staff Login');
}

function buildStaffLoginUrlFinal_(reservationId, token) {
  const webAppUrl = getScriptPropertyFinal_('WEB_APP_URL') || ScriptApp.getService().getUrl() || '';
  return webAppUrl + '?mode=staff_login&id=' + encodeURIComponent(reservationId || '') + '&token=' + encodeURIComponent(token || '');
}

function buildPickupCompleteUrlFinal_(reservationId, token, staffSession) {
  const webAppUrl = getScriptPropertyFinal_('WEB_APP_URL') || ScriptApp.getService().getUrl() || '';
  return webAppUrl + '?mode=pickup_complete&id=' + encodeURIComponent(reservationId || '') + '&token=' + encodeURIComponent(token || '') + '&staff_session=' + encodeURIComponent(staffSession || '');
}


function buildStaffActionHtmlFinal_(reservation, token, params) {
  const staffSession = String((params && params.staff_session) || '').trim();
  const staff = validateStaffSessionFinal_(staffSession);
  const status = String(reservation.status || '');
  const tags = String(reservation.luggage_tag_numbers || '').trim();

  if (status === 'PICKED_UP') {
    return `<div style="margin-top:14px;padding:20px 14px;border-radius:12px;background:#eefbea;color:#156a17;font-weight:900;text-align:center;font-size:clamp(22px,4.4vw,36px);line-height:1.22;">PICKED UP / 접수완료<br/><span style="display:block;margin-top:10px;font-size:clamp(30px,7vw,60px);letter-spacing:-.04em;">${escapeHtmlFinal_(tags || 'TAG NOT ASSIGNED')}</span><span style="display:inline-block;margin-top:8px;font-size:clamp(14px,2.6vw,22px);font-weight:800;">Staff: ${escapeHtmlFinal_(reservation.picked_up_by || (staff && staff.staff_id) || '')}</span></div>`;
  }

  return buildInlineStaffLoginHtmlFinal_(reservation, token);
}

function buildInlineStaffLoginHtmlFinal_(reservation, token) {
  const webAppUrl = getScriptPropertyFinal_('WEB_APP_URL') || ScriptApp.getService().getUrl() || '';
  const safeBase = escapeHtmlFinal_(webAppUrl);
  const safeId = escapeHtmlFinal_(reservation.reservation_id);
  const safeToken = escapeHtmlFinal_(token);
  const suitcaseCount = Math.max(1, Number(reservation.expected_suitcase_count || 1));
  const extraCount = Math.max(0, Number(reservation.expected_extra_bag_count || 0));

  return `
    <div id="staffBox" style="margin-top:18px;padding:18px;border:1px solid #ddd;border-radius:14px;background:#fafafa;">
      <div style="font-size:clamp(20px,3.8vw,30px);font-weight:900;margin-bottom:10px;">ONSITE CHECK-IN / 현장 접수</div>
      <div style="font-size:clamp(16px,3.2vw,24px);color:#555;line-height:1.45;margin-bottom:16px;font-weight:800;">실제 수량을 확인하면 추가 결제금액과 러기지택 번호가 자동 생성됩니다.</div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-bottom:12px;">
        <label style="font-size:clamp(15px,3vw,22px);font-weight:900;color:#555;">Suitcase / 캐리어<input id="actualSuitcaseInput" type="number" min="1" value="${suitcaseCount}" oninput="carryGoUpdateDueFinal()" style="box-sizing:border-box;width:100%;font-size:clamp(20px,4vw,32px);padding:14px;border:1px solid #ccc;border-radius:10px;margin-top:6px;"></label>
        <label style="font-size:clamp(15px,3vw,22px);font-weight:900;color:#555;">Extra / 추가짐<input id="actualExtraInput" type="number" min="0" value="${extraCount}" oninput="carryGoUpdateDueFinal()" style="box-sizing:border-box;width:100%;font-size:clamp(20px,4vw,32px);padding:14px;border:1px solid #ccc;border-radius:10px;margin-top:6px;"></label>
      </div>
      <div style="padding:14px;border:2px solid #111;border-radius:12px;background:#fff;margin-bottom:12px;">
        <div style="font-size:clamp(15px,3vw,22px);color:#555;font-weight:900;">현장 추가 결제금액 / Onsite Cash Due</div>
        <div id="onsiteDueBox" style="font-size:clamp(28px,5.4vw,44px);font-weight:950;letter-spacing:-.04em;line-height:1.05;">₩0</div>
        <div style="font-size:clamp(13px,2.7vw,20px);color:#666;line-height:1.45;font-weight:750;">기본요금에 포함된 캐리어 1개 제외. 추가 캐리어 ₩20,000 / 추가 짐 ₩10,000. 추가 짐: 지퍼/잠금장치 있는 가방만. 쇼핑백/비닐봉투 불가.</div>
      </div>
      <div style="margin:14px 0 16px;padding:14px;border-radius:12px;background:#fff;border:1px solid #ddd;font-size:clamp(16px,3.3vw,24px);line-height:1.45;color:#333;font-weight:850;">
        <div style="font-weight:950;margin-bottom:8px;">태그번호 발급 후 안내</div>
        <div>1. 태그번호를 짐과 동의서에 기재</div>
        <div>2. 러기지택을 짐에 부착</div>
        <div>3. 추가금이 있으면 현금 수납</div>
        <div>4. 태그번호가 보이게 사진 촬영</div>
      </div>
      <div id="staffLoginBox">
        <input id="staffCodeInput" placeholder="Staff code" autocomplete="off" style="box-sizing:border-box;width:100%;font-size:clamp(22px,4.8vw,34px);padding:18px;border:1px solid #ccc;border-radius:10px;margin:0 0 14px;">
      </div>
      <button type="button" onclick="carryGoOnsiteCheckinFinal()" style="width:100%;font-size:clamp(20px,4vw,30px);font-weight:900;background:#111;color:#fff;border:0;border-radius:10px;padding:19px 10px;">태그 발급 & 접수완료</button>
      <div id="staffMsg" style="font-size:clamp(16px,3.2vw,24px);color:#b00020;margin-top:14px;font-weight:850;line-height:1.4;"></div>
    </div>
    <script>
      const CARRYGO_WEBAPP_URL = '${safeBase}';
      const CARRYGO_RESERVATION_ID = '${safeId}';
      const CARRYGO_TOKEN = '${safeToken}';
      const CARRYGO_LOCAL_SESSION_KEY = 'carrygo_staff_session_v1';
      let CARRYGO_STAFF_SESSION = '';

      document.addEventListener('DOMContentLoaded', function() {
        carryGoUpdateDueFinal();
        carryGoTryStoredStaffSessionFinal();
      });

      function carryGoUpdateDueFinal() {
        const s = Math.max(1, Number(document.getElementById('actualSuitcaseInput').value || 1));
        const e = Math.max(0, Number(document.getElementById('actualExtraInput').value || 0));
        const due = Math.max(0, s - 1) * 20000 + e * 10000;
        document.getElementById('onsiteDueBox').textContent = '₩' + due.toLocaleString('ko-KR');
      }

      async function carryGoFetchJsonFinal(url) {
        const res = await fetch(url, { method: 'GET', cache: 'no-store' });
        const text = await res.text();
        try { return JSON.parse(text); } catch (e) { throw new Error('Invalid server response.'); }
      }

      async function carryGoTryStoredStaffSessionFinal() {
        const msg = document.getElementById('staffMsg');
        const stored = localStorage.getItem(CARRYGO_LOCAL_SESSION_KEY) || '';
        if (!stored) return;
        msg.style.color = '#555';
        msg.textContent = 'Staff session found. / 스태프 세션 확인됨';
        const url = CARRYGO_WEBAPP_URL + '?mode=staff_session_api&staff_session=' + encodeURIComponent(stored);
        try {
          const data = await carryGoFetchJsonFinal(url);
          if (!data.ok) throw new Error(data.error || 'Session expired.');
          CARRYGO_STAFF_SESSION = stored;
          document.getElementById('staffLoginBox').style.display = 'none';
          msg.textContent = 'Ready. Confirm quantity, collect cash, then issue tags. / 수량·현금 확인 후 태그 발급하세요.';
        } catch (err) {
          localStorage.removeItem(CARRYGO_LOCAL_SESSION_KEY);
          msg.textContent = 'Staff session expired. Please enter staff code.';
        }
      }

      async function carryGoOnsiteCheckinFinal() {
        const msg = document.getElementById('staffMsg');
        msg.style.color = '#555';
        msg.textContent = 'Saving...';
        const q = new URLSearchParams({
          mode: 'onsite_checkin_api',
          id: CARRYGO_RESERVATION_ID,
          token: CARRYGO_TOKEN,
          staff_session: CARRYGO_STAFF_SESSION,
          staff_code: document.getElementById('staffCodeInput') ? document.getElementById('staffCodeInput').value.trim() : '',
          actual_suitcase_count: document.getElementById('actualSuitcaseInput').value,
          actual_extra_bag_count: document.getElementById('actualExtraInput').value,
          onsite_cash_received: 'YES',
          onsite_tag_attached: 'YES',
          onsite_photo_taken: 'YES'
        });
        try {
          const data = await carryGoFetchJsonFinal(CARRYGO_WEBAPP_URL + '?' + q.toString());
          if (!data.ok) throw new Error(data.error || 'Check-in failed.');
          if (data.staff_session) localStorage.setItem(CARRYGO_LOCAL_SESSION_KEY, data.staff_session);
          document.getElementById('staffBox').innerHTML = '<div style="padding:20px 14px;border-radius:12px;background:#eefbea;color:#156a17;font-weight:900;text-align:center;font-size:clamp(22px,4.4vw,36px);line-height:1.22;">PICKED UP / 접수완료<br><span style="display:block;margin-top:10px;font-size:clamp(36px,8vw,68px);letter-spacing:-.04em;">' + escapeHtmlClientFinal(data.luggage_tag_numbers) + '</span><span style="display:block;margin-top:8px;font-size:clamp(18px,3vw,28px);">추가결제 ' + escapeHtmlClientFinal(data.onsite_due_display) + '</span><span style="display:block;margin-top:6px;font-size:13px;font-weight:700;">이 번호를 러기지택 고객용/짐부착용 양쪽에 기재하세요.</span></div>';
        } catch (err) {
          msg.style.color = '#b00020';
          msg.textContent = err.message || String(err);
        }
      }

      function escapeHtmlClientFinal(value) {
        return String(value || '').replace(/[&<>"']/g, function(ch) {
          return ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'})[ch];
        });
      }
    </script>`;
}

// ===== CarryGo Final Inline Staff API =====

function staffLoginApiFinal_(params) {
  try {
    const staffCode = String(params.staff_code || '').trim();
    if (!staffCode) throw new Error('Staff code is required.');
    const staff = findStaffByCodeFinal_(staffCode);
    if (!staff) throw new Error('Invalid staff code.');
    const sessionToken = createStaffSessionFinal_(staff);
    return jsonFinal_({
      ok: true,
      staff_session: sessionToken,
      staff_id: staff.staff_id,
      staff_name: staff.staff_name
    });
  } catch (err) {
    return jsonFinal_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function pickupCompleteApiFinal_(params) {
  return jsonFinal_({ ok: false, error: 'Use onsite_checkin_api to confirm counts, cash due, and luggage tags before PICKED_UP.' });
}

function formatDateTimeMaybeFinal_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return formatDateTimeFinal_(value);
  }
  return String(value);
}

function parseDateTimeFinal_(value) {
  if (!value) return null;
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) return value;
  const text = String(value || '').trim();
  const match = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (!match) return null;
  return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]), Number(match[4] || 0), Number(match[5] || 0), Number(match[6] || 0));
}

function isPaymentExpiredFinal_(paymentDueAt, now) {
  const due = parseDateTimeFinal_(paymentDueAt);
  if (!due) return false;
  return due.getTime() <= (now || new Date()).getTime();
}



function adminResetCheckinTestsApiFinal_(params) {
  try {
    validateAdminPinFinal_(params.admin_pin);
    const sh = getSheetFinal_(CARRYGO_SHEETS.RESERVATIONS);
    const values = sh.getDataRange().getValues();
    const changed = [];
    const clearHeaders = [
      'picked_up_at',
      'picked_up_by',
      'luggage_tag_numbers',
      'onsite_payment_method',
      'onsite_staff',
      'onsite_consent_flags',
      'actual_suitcase_count',
      'actual_extra_bag_count',
      'onsite_due_amount',
      'onsite_cash_received',
      'onsite_tag_attached',
      'onsite_photo_taken',
      'onsite_checkin_completed_at'
    ];
    const col = name => RESERVATIONS_HEADERS.indexOf(name);
    const statusCol = col('status');
    const paymentStatusCol = col('payment_status');
    const bookingChannelCol = col('booking_channel');
    const ridCol = col('reservation_id');
    const tagCol = col('luggage_tag_numbers');
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const hadTags = String(row[tagCol] || '').trim();
      const wasPickedUp = String(row[statusCol] || '').trim() === 'PICKED_UP';
      const isPaid = String(row[paymentStatusCol] || '').trim() === 'PAID';
      const isWalkIn = String(row[bookingChannelCol] || '').trim() === 'WALK_IN';
      if (!hadTags && !wasPickedUp) continue;
      clearHeaders.forEach(header => {
        const c = col(header);
        if (c >= 0) sh.getRange(i + 1, c + 1).setValue('');
      });
      if (wasPickedUp && isPaid && !isWalkIn) {
        sh.getRange(i + 1, statusCol + 1).setValue('CONFIRMED');
      }
      changed.push({ row: i + 1, reservation_id: row[ridCol], previous_status: row[statusCol], previous_tags: hadTags });
    }
    return jsonFinal_({ ok: true, changed_count: changed.length, changed: changed });
  } catch (err) {
    return jsonFinal_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}


function adminNormalizeLuggageTagsApiFinal_(params) {
  try {
    validateAdminPinFinal_(params.admin_pin);
    const sh = getSheetFinal_(CARRYGO_SHEETS.RESERVATIONS);
    const values = sh.getDataRange().getValues();
    const tagCol = RESERVATIONS_HEADERS.indexOf('luggage_tag_numbers') + 1;
    const ridCol = RESERVATIONS_HEADERS.indexOf('reservation_id');
    const changed = [];
    for (let i = 1; i < values.length; i++) {
      const current = String(values[i][tagCol - 1] || '').trim();
      if (!current) continue;
      const normalized = normalizeLuggageTagStringFinal_(current);
      if (normalized && normalized !== current) {
        sh.getRange(i + 1, tagCol).setValue(normalized);
        changed.push({ row: i + 1, reservation_id: values[i][ridCol], before: current, after: normalized });
      }
    }
    return jsonFinal_({ ok: true, changed_count: changed.length, changed: changed });
  } catch (err) {
    return jsonFinal_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}


function onsiteLookupApiFinal_(params) {
  try {
    const reservationId = String(params.id || '').trim();
    const token = String(params.token || '').trim();
    if (!reservationId || !token) throw new Error('Missing QR information.');
    const rowNo = findReservationRowFinal_(reservationId);
    const r = getReservationObjectByRowFinal_(rowNo);
    if (String(r.checkin_token || '') !== token) throw new Error('Invalid or expired QR code.');
    const normalizedTags = normalizeLuggageTagStringFinal_(r.luggage_tag_numbers || '');
    if (normalizedTags && normalizedTags !== String(r.luggage_tag_numbers || '').trim()) {
      const sh = getSheetFinal_(CARRYGO_SHEETS.RESERVATIONS);
      setReservationValueFinal_(sh, rowNo, 'luggage_tag_numbers', normalizedTags);
      r.luggage_tag_numbers = normalizedTags;
    }
    return jsonFinal_({ ok: true, reservation: adminReservationSummaryFinal_(r) });
  } catch (err) {
    return jsonFinal_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}


function onsiteCheckinApiFinal_(params) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const reservationId = String(params.id || '').trim();
    const token = String(params.token || '').trim();
    if (!reservationId || !token) throw new Error('Missing QR information.');

    let staff = validateStaffSessionFinal_(String(params.staff_session || '').trim());
    let sessionToken = String(params.staff_session || '').trim();
    if (!staff && params.staff_code) {
      staff = findStaffByCodeFinal_(params.staff_code);
      if (staff) sessionToken = createStaffSessionFinal_(staff);
    }
    if (!staff) throw new Error('Staff login is required.');

    const rowNo = findReservationRowFinal_(reservationId);
    const r = getReservationObjectByRowFinal_(rowNo);
    if (String(r.checkin_token || '') !== token) throw new Error('Invalid or expired QR code.');
    if (String(r.status || '') === 'RETURNED') throw new Error('This reservation is already returned.');
    if (String(r.status || '') === 'CANCELLED') throw new Error('This reservation is cancelled.');
    if (String(r.status || '') === 'PICKED_UP') {
      return jsonFinal_({
        ok: true,
        already_assigned: true,
        reservation_id: reservationId,
        status: 'PICKED_UP',
        luggage_tag_numbers: normalizeLuggageTagStringFinal_(r.luggage_tag_numbers || ''),
        actual_suitcase_count: r.actual_suitcase_count || r.expected_suitcase_count || '',
        actual_extra_bag_count: r.actual_extra_bag_count || r.expected_extra_bag_count || '',
        onsite_due_amount: r.onsite_due_amount || 0,
        onsite_due_display: '₩' + Number(r.onsite_due_amount || 0).toLocaleString('ko-KR'),
        onsite_cash_received: r.onsite_cash_received || '',
        onsite_tag_attached: r.onsite_tag_attached || '',
        onsite_photo_taken: r.onsite_photo_taken || '',
        picked_up_by: r.picked_up_by || '',
        picked_up_at: formatDateTimeMaybeFinal_(r.picked_up_at)
      });
    }
    if (String(r.status || '') !== 'CONFIRMED' || String(r.payment_status || '') !== 'PAID') {
      throw new Error('Reservation is not confirmed/paid yet.');
    }

    const actualSuitcase = normalizeCountFinal_(params.actual_suitcase_count || r.expected_suitcase_count, 1);
    const actualExtra = normalizeCountFinal_(params.actual_extra_bag_count || r.expected_extra_bag_count, 0);
    if (actualSuitcase < 1) throw new Error('actual_suitcase_count must be at least 1');
    const onsiteDue = Math.max(0, actualSuitcase - 1) * 20000 + actualExtra * 10000;
    const existingTags = normalizeLuggageTagStringFinal_(r.luggage_tag_numbers || '');
    const tags = existingTags || nextLuggageTagNumbersFinal_(actualSuitcase + actualExtra, r.concert_id);

    const sh = getSheetFinal_(CARRYGO_SHEETS.RESERVATIONS);
    const now = new Date();
    setReservationValueFinal_(sh, rowNo, 'expected_suitcase_count', actualSuitcase);
    setReservationValueFinal_(sh, rowNo, 'expected_extra_bag_count', actualExtra);
    setReservationValueFinal_(sh, rowNo, 'actual_suitcase_count', actualSuitcase);
    setReservationValueFinal_(sh, rowNo, 'actual_extra_bag_count', actualExtra);
    setReservationValueFinal_(sh, rowNo, 'onsite_due_amount', onsiteDue);
    setReservationValueFinal_(sh, rowNo, 'onsite_cash_received', 'YES');
    setReservationValueFinal_(sh, rowNo, 'onsite_tag_attached', 'YES');
    setReservationValueFinal_(sh, rowNo, 'onsite_photo_taken', 'YES');
    setReservationValueFinal_(sh, rowNo, 'onsite_checkin_completed_at', now);
    setReservationValueFinal_(sh, rowNo, 'onsite_payment_method', onsiteDue > 0 ? 'CASH' : 'NONE');
    setReservationValueFinal_(sh, rowNo, 'luggage_tag_numbers', tags);
    setReservationValueFinal_(sh, rowNo, 'onsite_staff', staff.staff_id || staff.staff_name || '');
    setReservationValueFinal_(sh, rowNo, 'onsite_consent_flags', mergeNoteFinal_(r.onsite_consent_flags || '', 'HARDCOPY_CONFIRMED=YES; ONSITE_CHECKIN_BY=' + (staff.staff_id || staff.staff_name || 'STAFF')));
    setReservationValueFinal_(sh, rowNo, 'status', 'PICKED_UP');
    setReservationValueFinal_(sh, rowNo, 'picked_up_at', now);
    setReservationValueFinal_(sh, rowNo, 'picked_up_by', staff.staff_id || staff.staff_name || 'STAFF');

    return jsonFinal_({
      ok: true,
      staff_session: sessionToken,
      reservation_id: reservationId,
      status: 'PICKED_UP',
      already_assigned: !!existingTags,
      luggage_tag_numbers: tags,
      actual_suitcase_count: actualSuitcase,
      actual_extra_bag_count: actualExtra,
      onsite_due_amount: onsiteDue,
      onsite_due_display: '₩' + Number(onsiteDue || 0).toLocaleString('ko-KR'),
      onsite_cash_received: 'YES',
      onsite_tag_attached: 'YES',
      onsite_photo_taken: 'YES',
      picked_up_by: staff.staff_id || staff.staff_name || 'STAFF',
      picked_up_at: formatDateTimeFinal_(now)
    });
  } catch (err) {
    return jsonFinal_({ ok: false, error: String(err && err.message ? err.message : err) });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function staffSessionApiFinal_(params) {
  try {
    const staffSession = String(params.staff_session || '').trim();
    const staff = validateStaffSessionFinal_(staffSession);
    if (!staff) throw new Error('Staff session expired. Please log in again.');
    return jsonFinal_({ ok: true, staff_id: staff.staff_id, staff_name: staff.staff_name });
  } catch (err) {
    return jsonFinal_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

// ===== CarryGo Final Admin Payment Confirmation =====

function renderAdminPageFinal_(params) {
  const webAppUrl = getScriptPropertyFinal_('WEB_APP_URL') || ScriptApp.getService().getUrl() || '';
  const adminPin = String(params.admin_pin || '').trim();
  let messageHtml = '';
  let listHtml = '';

  if (adminPin) {
    try {
      validateAdminPinFinal_(adminPin);
      const rows = readSheetObjectsFinal_(CARRYGO_SHEETS.RESERVATIONS, RESERVATIONS_HEADERS)
        .filter(row => String(row.status || '') === 'UNPAID')
        .slice(-50)
        .reverse();

      if (!rows.length) {
        listHtml = '<div class="card ok">No UNPAID reservations.</div>';
      } else {
        const rowsHtml = rows.map(row => {
          const amountDisplay = String(row.currency || '').toUpperCase() === 'USD'
            ? '$' + row.base_fee
            : '₩' + Number(row.base_fee || 0).toLocaleString('ko-KR');
          return [
            '<label class="row">',
            '<input type="checkbox" name="reservation_id" value="' + escapeHtmlFinal_(row.reservation_id) + '">',
            '<div>',
            '<div class="rid">' + escapeHtmlFinal_(row.reservation_id) + '</div>',
            '<div class="meta">',
            escapeHtmlFinal_(row.customer_name) + ' · ' + escapeHtmlFinal_(row.customer_email) + '<br>',
            escapeHtmlFinal_(row.concert_date) + ' ' + escapeHtmlFinal_(row.concert_time) + ' · ' + escapeHtmlFinal_(row.payment_method) + ' · ' + escapeHtmlFinal_(amountDisplay) + '<br>',
            'Due: ' + escapeHtmlFinal_(formatDateTimeMaybeFinal_(row.payment_due_at)),
            '</div>',
            '</div>',
            '</label>'
          ].join('');
        }).join('');

        listHtml = [
          '<form method="get" action="' + escapeHtmlFinal_(webAppUrl) + '">',
          '<input type="hidden" name="mode" value="admin_confirm_selected_page">',
          '<input type="hidden" name="admin_pin" value="' + escapeHtmlFinal_(adminPin) + '">',
          rowsHtml,
          '<div class="toolbar"><button type="submit">Confirm Selected / 선택 입금확인</button></div>',
          '</form>'
        ].join('');
      }
    } catch (err) {
      messageHtml = '<div class="err">' + escapeHtmlFinal_(String(err && err.message ? err.message : err)) + '</div>';
    }
  }

  const html = `
    <!doctype html>
    <html>
    <head>
      <meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no">
      <title>CarryGo Admin</title>
      <style>
        body{font-family:Arial,Helvetica,sans-serif;margin:0;background:#f7f4ef;color:#111;}
        .wrap{max-width:560px;margin:0 auto;padding:16px;}
        .card{background:#fff;border:2px solid #111;border-radius:18px;padding:18px;margin-bottom:14px;}
        h1{font-size:30px;line-height:.95;margin:0 0 8px;font-weight:900;letter-spacing:-.04em;}
        p{font-size:14px;color:#555;line-height:1.45;margin:0 0 14px;}
        label.pin{display:block;font-size:11px;font-weight:900;letter-spacing:.16em;text-transform:uppercase;color:#736f68;margin:12px 0 6px;}
        input.pin{box-sizing:border-box;width:100%;border:1.4px solid #d8d1c7;border-radius:12px;padding:14px;font-size:18px;font-weight:800;}
        button{width:100%;box-sizing:border-box;border:0;border-radius:999px;background:#050505;color:#fff;padding:15px 12px;font-size:15px;font-weight:900;margin-top:12px;}
        .row{border:1.4px solid #d8d1c7;border-radius:14px;padding:14px;margin-top:10px;background:#fff;display:grid;grid-template-columns:32px 1fr;gap:10px;align-items:start;}
        .row input{width:22px;height:22px;margin:4px 0 0;padding:0;}
        .rid{font-size:22px;font-weight:900;letter-spacing:-.03em;}
        .meta{font-size:13px;color:#555;line-height:1.45;margin-top:5px;}
        .ok{background:#ecfff7;border-color:#10b981;color:#065f46;}
        .err{background:#fff3f0;border:1px solid #f1b8ac;border-radius:12px;color:#8f3c32;font-size:13px;font-weight:800;margin:10px 0;padding:12px;white-space:pre-wrap;}
        .toolbar{position:sticky;bottom:0;background:#f7f4ef;padding:10px 0 4px;margin-top:10px;}
      </style>
    </head>
    <body>
      <div class="wrap">
        <div class="card">
          <h1>CarryGo<br>Admin</h1>
          <p>UNPAID 예약을 불러온 뒤 체크박스로 선택해서 입금확인 처리합니다.</p>
          <form method="get" action="${escapeHtmlFinal_(webAppUrl)}">
            <input type="hidden" name="mode" value="admin">
            <label class="pin">Admin PIN</label>
            <input class="pin" name="admin_pin" type="password" placeholder="PIN" value="${escapeHtmlFinal_(adminPin)}">
            <button type="submit">Load UNPAID</button>
          </form>
          ${messageHtml}
        </div>
        ${listHtml}
      </div>
    </body>
    </html>`;
  return HtmlService.createHtmlOutput(html).setTitle('CarryGo Admin');
}

function renderAdminConfirmPaymentPageFinal_(params) {
  try {
    validateAdminPinFinal_(params.admin_pin);
    const reservationId = String(params.reservation_id || '').trim();
    if (!reservationId) throw new Error('reservation_id is required');
    const rowNo = findReservationRowFinal_(reservationId);
    const row = getReservationObjectByRowFinal_(rowNo);
    confirmPaymentFinal(reservationId, row.base_fee || '');
    return HtmlService.createHtmlOutput(`
      <div style="font-family:Arial,sans-serif;max-width:520px;margin:0 auto;padding:20px;background:#f7f4ef;color:#111;">
        <div style="background:#ecfff7;border:2px solid #10b981;border-radius:18px;padding:18px;">
          <h1 style="margin:0 0 8px;font-size:30px;line-height:.95;">CONFIRMED</h1>
          <p style="font-size:16px;line-height:1.45;">${escapeHtmlFinal_(reservationId)} 입금확인 완료.<br>QR 확정메일을 발송했습니다.</p>
          <a style="display:block;text-align:center;text-decoration:none;border-radius:999px;background:#050505;color:#fff;padding:15px 12px;font-weight:900;margin-top:14px;" href="${escapeHtmlFinal_((getScriptPropertyFinal_('WEB_APP_URL') || ScriptApp.getService().getUrl() || '') + '?mode=admin&admin_pin=' + encodeURIComponent(params.admin_pin || ''))}">Back to Admin</a>
        </div>
      </div>`).setTitle('CarryGo Confirmed');
  } catch (err) {
    return HtmlService.createHtmlOutput(`
      <div style="font-family:Arial,sans-serif;max-width:520px;margin:0 auto;padding:20px;background:#f7f4ef;color:#111;">
        <div style="background:#fff3f0;border:2px solid #8f3c32;border-radius:18px;padding:18px;">
          <h1 style="margin:0 0 8px;font-size:30px;line-height:.95;">ERROR</h1>
          <pre style="white-space:pre-wrap;font-size:14px;line-height:1.45;">${escapeHtmlFinal_(String(err && err.message ? err.message : err))}</pre>
        </div>
      </div>`).setTitle('CarryGo Admin Error');
  }
}

function adminListUnpaidApiFinal_(params) {
  try {
    validateAdminPinFinal_(params.admin_pin);
    const rows = readSheetObjectsFinal_(CARRYGO_SHEETS.RESERVATIONS, RESERVATIONS_HEADERS);
    const now = new Date();
    const reservations = rows
      .filter(row => String(row.status || '') === 'UNPAID')
      .map(row => adminReservationSummaryFinal_(row))
      .map(row => {
        row.is_payment_expired = isPaymentExpiredFinal_(row.payment_due_at, now);
        return row;
      })
      .sort((a, b) => {
        if (a.is_payment_expired !== b.is_payment_expired) return a.is_payment_expired ? 1 : -1;
        return String(b.reservation_id || '').localeCompare(String(a.reservation_id || ''));
      })
      .slice(0, 100);
    return jsonFinal_({ ok: true, reservations: reservations });
  } catch (err) {
    return jsonFinal_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function adminCancelExpiredUnpaidApiFinal_(params) {
  try {
    validateAdminPinFinal_(params.admin_pin);
    const ids = splitIdsFinal_(params.reservation_id);
    if (!ids.length) throw new Error('reservation_id is required');
    const sh = getSheetFinal_(CARRYGO_SHEETS.RESERVATIONS);
    const now = new Date();
    const cancelled = [];
    const errors = [];
    ids.forEach(reservationId => {
      try {
        const rowNo = findReservationRowFinal_(reservationId);
        const row = getReservationObjectByRowFinal_(rowNo);
        if (String(row.status || '') !== 'UNPAID') throw new Error('UNPAID 상태만 취소할 수 있습니다.');
        if (!isPaymentExpiredFinal_(row.payment_due_at, now)) throw new Error('결제기한이 지나지 않았습니다.');
        setReservationValueFinal_(sh, rowNo, 'status', 'CANCELLED');
        setReservationValueFinal_(sh, rowNo, 'cancelled_at', now);
        setReservationValueFinal_(sh, rowNo, 'note', mergeNoteFinal_(row.note, '미입금 취소'));
        cancelled.push(reservationId);
      } catch (err) {
        errors.push(reservationId + ': ' + String(err && err.message ? err.message : err));
      }
    });
    if (errors.length) return jsonFinal_({ ok: false, cancelled: cancelled, error: errors.join('\n') });
    return jsonFinal_({ ok: true, cancelled: cancelled, status: 'CANCELLED' });
  } catch (err) {
    return jsonFinal_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function adminConfirmPaymentApiFinal_(params) {
  try {
    validateAdminPinFinal_(params.admin_pin);
    const ids = String(params.reservation_id || '').split(',').map(id => id.trim()).filter(Boolean);
    if (!ids.length) throw new Error('reservation_id is required');
    const confirmed = [];
    const errors = [];
    ids.forEach(reservationId => {
      try {
        const rowNo = findReservationRowFinal_(reservationId);
        const row = getReservationObjectByRowFinal_(rowNo);
        const currentStatus = String(row.status || '').toUpperCase();
        const paymentStatus = String(row.payment_status || '').toUpperCase();
        if (currentStatus === 'CONFIRMED' && paymentStatus === 'PAID') {
          confirmed.push(reservationId);
          return;
        }
        if (currentStatus !== 'UNPAID') throw new Error('UNPAID 상태만 입금확인 처리할 수 있습니다: ' + reservationId);
        if (paymentStatus === 'PAID' || paymentStatus === 'REFUNDED') throw new Error('이미 결제 처리된 예약입니다: ' + reservationId);
        const paidAmount = row.base_fee || '';
        confirmPaymentFinal(reservationId, paidAmount);
        confirmed.push(reservationId);
      } catch (err) {
        errors.push(reservationId + ': ' + String(err && err.message ? err.message : err));
      }
    });
    if (errors.length) return jsonFinal_({ ok: false, confirmed: confirmed, error: errors.join('\n') });
    return jsonFinal_({ ok: true, reservation_ids: confirmed, status: 'CONFIRMED' });
  } catch (err) {
    return jsonFinal_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function validateAdminPinFinal_(adminPin) {
  const expected = getScriptPropertyFinal_('ADMIN_PIN');
  if (!expected) throw new Error('ADMIN_PIN script property is not set.');
  if (String(adminPin || '').trim() !== String(expected).trim()) throw new Error('Invalid admin PIN.');
}

function adminListUnpaidClientFinal(adminPin) {
  try {
    validateAdminPinFinal_(adminPin);
    const rows = readSheetObjectsFinal_(CARRYGO_SHEETS.RESERVATIONS, RESERVATIONS_HEADERS);
    const reservations = rows
      .filter(row => String(row.status || '') === 'UNPAID')
      .slice(-50)
      .reverse()
      .map(row => ({
        reservation_id: row.reservation_id,
        customer_name: row.customer_name,
        customer_email: row.customer_email,
        concert_date: row.concert_date,
        concert_time: row.concert_time,
        pickup_time: row.pickup_time,
        booking_channel: row.booking_channel,
        luggage_tag_numbers: row.luggage_tag_numbers,
        expected_suitcase_count: row.expected_suitcase_count,
        expected_extra_bag_count: row.expected_extra_bag_count,
        payment_method: row.payment_method,
        amount_display: String(row.currency || '').toUpperCase() === 'USD' ? '$' + row.base_fee : '₩' + Number(row.base_fee || 0).toLocaleString('ko-KR'),
        payment_due_at: formatDateTimeMaybeFinal_(row.payment_due_at)
      }));
    return { ok: true, reservations: reservations };
  } catch (err) {
    return { ok: false, error: String(err && err.message ? err.message : err) };
  }
}

function adminConfirmPaymentsClientFinal(adminPin, reservationIds) {
  try {
    validateAdminPinFinal_(adminPin);
    const ids = Array.isArray(reservationIds) ? reservationIds : [];
    if (!ids.length) throw new Error('No reservation selected.');
    const confirmed = [];
    const errors = [];
    ids.forEach(id => {
      try {
        const reservationId = String(id || '').trim();
        if (!reservationId) return;
        const rowNo = findReservationRowFinal_(reservationId);
        const row = getReservationObjectByRowFinal_(rowNo);
        if (String(row.status || '') === 'CONFIRMED') {
          confirmed.push(reservationId);
          return;
        }
        confirmPaymentFinal(reservationId, row.base_fee || '');
        confirmed.push(reservationId);
      } catch (err) {
        errors.push(String(id) + ': ' + String(err && err.message ? err.message : err));
      }
    });
    if (errors.length) return { ok: false, confirmed: confirmed, error: errors.join('\n') };
    return { ok: true, confirmed: confirmed };
  } catch (err) {
    return { ok: false, error: String(err && err.message ? err.message : err) };
  }
}

function renderAdminConfirmSelectedPageFinal_(params) {
  try {
    validateAdminPinFinal_(params.admin_pin);
    let ids = params.reservation_id || [];
    if (!Array.isArray(ids)) ids = ids ? [ids] : [];
    ids = ids.map(id => String(id || '').trim()).filter(Boolean);
    if (!ids.length) throw new Error('선택된 예약이 없습니다.');

    const confirmed = [];
    const errors = [];
    ids.forEach(reservationId => {
      try {
        const rowNo = findReservationRowFinal_(reservationId);
        const row = getReservationObjectByRowFinal_(rowNo);
        if (String(row.status || '') === 'CONFIRMED') {
          confirmed.push(reservationId + ' (already confirmed)');
          return;
        }
        confirmPaymentFinal(reservationId, row.base_fee || '');
        confirmed.push(reservationId);
      } catch (err) {
        errors.push(reservationId + ': ' + String(err && err.message ? err.message : err));
      }
    });

    const webAppUrl = getScriptPropertyFinal_('WEB_APP_URL') || ScriptApp.getService().getUrl() || '';
    const backUrl = webAppUrl + '?mode=admin&admin_pin=' + encodeURIComponent(params.admin_pin || '');
    const confirmedHtml = confirmed.map(id => '<li>' + escapeHtmlFinal_(id) + '</li>').join('');
    const errorHtml = errors.length ? '<div style="background:#fff3f0;border:1px solid #f1b8ac;border-radius:12px;color:#8f3c32;padding:12px;margin-top:12px;"><b>Errors</b><br>' + errors.map(escapeHtmlFinal_).join('<br>') + '</div>' : '';

    return HtmlService.createHtmlOutput(`
      <div style="font-family:Arial,sans-serif;max-width:520px;margin:0 auto;padding:20px;background:#f7f4ef;color:#111;">
        <div style="background:#ecfff7;border:2px solid #10b981;border-radius:18px;padding:18px;">
          <h1 style="margin:0 0 8px;font-size:30px;line-height:.95;">CONFIRMED</h1>
          <p style="font-size:16px;line-height:1.45;">입금확인 완료. QR 확정메일을 발송했습니다.</p>
          <ul style="font-size:16px;line-height:1.5;font-weight:800;">${confirmedHtml}</ul>
          ${errorHtml}
          <a style="display:block;text-align:center;text-decoration:none;border-radius:999px;background:#050505;color:#fff;padding:15px 12px;font-weight:900;margin-top:14px;" href="${escapeHtmlFinal_(backUrl)}">Back to Admin</a>
        </div>
      </div>`).setTitle('CarryGo Confirmed');
  } catch (err) {
    const webAppUrl = getScriptPropertyFinal_('WEB_APP_URL') || ScriptApp.getService().getUrl() || '';
    return HtmlService.createHtmlOutput(`
      <div style="font-family:Arial,sans-serif;max-width:520px;margin:0 auto;padding:20px;background:#f7f4ef;color:#111;">
        <div style="background:#fff3f0;border:2px solid #8f3c32;border-radius:18px;padding:18px;">
          <h1 style="margin:0 0 8px;font-size:30px;line-height:.95;">ERROR</h1>
          <pre style="white-space:pre-wrap;font-size:14px;line-height:1.45;">${escapeHtmlFinal_(String(err && err.message ? err.message : err))}</pre>
          <a style="display:block;text-align:center;text-decoration:none;border-radius:999px;background:#050505;color:#fff;padding:15px 12px;font-weight:900;margin-top:14px;" href="${escapeHtmlFinal_(webAppUrl + '?mode=admin')}">Back to Admin</a>
        </div>
      </div>`).setTitle('CarryGo Admin Error');
  }
}

// ===== CarryGo Final Admin Operations API =====

function adminListByStatusApiFinal_(params) {
  try {
    validateAdminPinFinal_(params.admin_pin);
    const status = String(params.status || '').trim().toUpperCase();
    if (!status) throw new Error('status is required');
    const rows = readSheetObjectsFinal_(CARRYGO_SHEETS.RESERVATIONS, RESERVATIONS_HEADERS);
    const reservations = rows
      .filter(row => String(row.status || '').toUpperCase() === status)
      .filter(row => {
        const paymentStatus = String(row.payment_status || '').toUpperCase();
        if (status === 'CANCELLED') return paymentStatus === 'PAID' && !row.refunded_at;
        if (status === 'PICKED_UP') return paymentStatus === 'PAID';
        if (status === 'UNPAID') return paymentStatus !== 'PAID' && paymentStatus !== 'REFUNDED';
        return true;
      })
      .slice(-100)
      .reverse()
      .map(row => adminReservationSummaryFinal_(row));
    return jsonFinal_({ ok: true, reservations: reservations });
  } catch (err) {
    return jsonFinal_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function adminUpdateStatusApiFinal_(params) {
  try {
    validateAdminPinFinal_(params.admin_pin);
    const ids = splitIdsFinal_(params.reservation_id);
    const nextStatus = String(params.next_status || '').trim().toUpperCase();
    if (!ids.length) throw new Error('reservation_id is required');
    if (!['RETURNED', 'CANCELLED'].includes(nextStatus)) throw new Error('Unsupported next_status: ' + nextStatus);

    const updated = [];
    ids.forEach(reservationId => {
      const rowNo = findReservationRowFinal_(reservationId);
      const row = getReservationObjectByRowFinal_(rowNo);
      const currentStatus = String(row.status || '').toUpperCase();
      const paymentStatus = String(row.payment_status || '').toUpperCase();
      if (nextStatus === 'RETURNED') {
        if (currentStatus !== 'PICKED_UP') throw new Error('PICKED_UP 상태만 수령완료 처리할 수 있습니다: ' + reservationId);
        if (paymentStatus !== 'PAID') throw new Error('결제완료 예약만 수령완료 처리할 수 있습니다: ' + reservationId);
      }
      if (nextStatus === 'CANCELLED') {
        if (currentStatus !== 'UNPAID') throw new Error('UNPAID 상태만 예약취소 처리할 수 있습니다: ' + reservationId);
        if (paymentStatus === 'PAID' || paymentStatus === 'REFUNDED') throw new Error('결제완료 예약은 예약취소 탭에서 취소할 수 없습니다: ' + reservationId);
      }
      const sh = getSheetFinal_(CARRYGO_SHEETS.RESERVATIONS);
      const now = new Date();
      setReservationValueFinal_(sh, rowNo, 'status', nextStatus);
      if (nextStatus === 'RETURNED') setReservationValueFinal_(sh, rowNo, 'returned_at', now);
      if (nextStatus === 'CANCELLED') {
        setReservationValueFinal_(sh, rowNo, 'cancelled_at', now);
        if (params.note) setReservationValueFinal_(sh, rowNo, 'note', mergeNoteFinal_(row.note, params.note));
      }
      updated.push(reservationId);
    });
    return jsonFinal_({ ok: true, updated: updated, status: nextStatus });
  } catch (err) {
    return jsonFinal_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function adminRefundApiFinal_(params) {
  try {
    validateAdminPinFinal_(params.admin_pin);
    const ids = splitIdsFinal_(params.reservation_id);
    if (!ids.length) throw new Error('reservation_id is required');
    const refundAmount = String(params.refund_amount || '').trim();
    const refundMethod = String(params.refund_method || '').trim();
    const refundNote = String(params.refund_note || '').trim();
    if (!refundAmount) throw new Error('refund_amount is required');
    if (!refundMethod) throw new Error('refund_method is required');

    const updated = [];
    ids.forEach(reservationId => {
      const rowNo = findReservationRowFinal_(reservationId);
      const row = getReservationObjectByRowFinal_(rowNo);
      const currentStatus = String(row.status || '').toUpperCase();
      const paymentStatus = String(row.payment_status || '').toUpperCase();
      if (currentStatus !== 'CANCELLED') throw new Error('CANCELLED 상태만 환불기록할 수 있습니다: ' + reservationId);
      if (paymentStatus !== 'PAID') throw new Error('결제완료 취소건만 환불기록할 수 있습니다: ' + reservationId);
      if (row.refunded_at) throw new Error('이미 환불기록된 예약입니다: ' + reservationId);
      const sh = getSheetFinal_(CARRYGO_SHEETS.RESERVATIONS);
      const now = new Date();
      setReservationValueFinal_(sh, rowNo, 'refunded_at', now);
      setReservationValueFinal_(sh, rowNo, 'refund_amount', refundAmount);
      setReservationValueFinal_(sh, rowNo, 'refund_method', refundMethod);
      setReservationValueFinal_(sh, rowNo, 'refund_note', refundNote);
      setReservationValueFinal_(sh, rowNo, 'payment_status', 'REFUNDED');
      updated.push(reservationId);
    });
    return jsonFinal_({ ok: true, updated: updated });
  } catch (err) {
    return jsonFinal_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function adminCreateConcertApiFinal_(params) {
  try {
    validateAdminPinFinal_(params.admin_pin);
    const now = new Date();
    const rows = readSheetObjectsFinal_(CARRYGO_SHEETS.CONCERTS, CONCERTS_HEADERS);
    const title = requiredStringFinal_(params.concert_title, 'concert_title');
    const venue = requiredStringFinal_(params.venue, 'venue');
    const city = String(params.city || 'SEOUL').trim().toUpperCase();
    const concertId = String(params.concert_id || '').trim() || uniqueConcertIdFinal_(title, rows);
    const concertCode = String(params.concert_code || '').trim().toUpperCase() || uniqueConcertCodeFinal_(title, rows);
    if (rows.some(row => String(row.concert_id) === concertId)) throw new Error('concert_id already exists: ' + concertId);
    if (rows.some(row => String(row.concert_code).toUpperCase() === concertCode)) throw new Error('concert_code already exists: ' + concertCode);
    const obj = {
      concert_id: concertId,
      concert_code: concertCode,
      concert_title: title,
      venue: venue,
      city: city,
      is_active: String(params.is_active || 'TRUE').toUpperCase() !== 'FALSE',
      sort_order: normalizeCountFinal_(params.sort_order, rows.length + 1),
      created_at: now,
      updated_at: now
    };
    getSheetFinal_(CARRYGO_SHEETS.CONCERTS).appendRow(CONCERTS_HEADERS.map(h => obj[h] !== undefined ? obj[h] : ''));
    return jsonFinal_({ ok: true, concert: obj });
  } catch (err) {
    return jsonFinal_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function adminCreateConcertDateApiFinal_(params) {
  try {
    validateAdminPinFinal_(params.admin_pin);
    const now = new Date();
    const rows = readSheetObjectsFinal_(CARRYGO_SHEETS.CONCERT_DATES, CONCERT_DATES_HEADERS);
    const concertId = requiredStringFinal_(params.concert_id, 'concert_id');
    const concertDate = requiredStringFinal_(params.concert_date, 'concert_date');
    const concertTime = requiredStringFinal_(params.concert_time, 'concert_time');
    const concertDateId = String(params.concert_date_id || '').trim() || uniqueConcertDateIdFinal_(concertId, concertDate, concertTime, rows);
    if (rows.some(row => String(row.concert_date_id) === concertDateId)) throw new Error('concert_date_id already exists: ' + concertDateId);
    const obj = {
      concert_date_id: concertDateId,
      concert_id: concertId,
      concert_date: concertDate,
      concert_time: concertTime,
      pickup_time_options: normalizePickupTimeOptionsFinal_(params.pickup_time_options).join(','),
      pickup_drop_guide_link: String(params.pickup_drop_guide_link || '').trim(),
      next_day_pickup_guide_link: String(params.next_day_pickup_guide_link || '').trim(),
      location_change_guide_link: String(params.location_change_guide_link || '').trim(),
      is_active: String(params.is_active || 'TRUE').toUpperCase() !== 'FALSE',
      sort_order: normalizeCountFinal_(params.sort_order, rows.length + 1),
      created_at: now,
      updated_at: now
    };
    getSheetFinal_(CARRYGO_SHEETS.CONCERT_DATES).appendRow(CONCERT_DATES_HEADERS.map(h => obj[h] !== undefined ? obj[h] : ''));
    return jsonFinal_({ ok: true, concert_date: obj });
  } catch (err) {
    return jsonFinal_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function adminCreateConcertBundleApiFinal_(params) {
  try {
    validateAdminPinFinal_(params.admin_pin);
    const now = new Date();
    const title = requiredStringFinal_(params.concert_title, 'concert_title');
    const venue = requiredStringFinal_(params.venue, 'venue');
    const city = String(params.city || 'SEOUL').trim().toUpperCase();
    let dateLines = parseConcertDateLinesFinal_(params.date_lines);
    if (!dateLines.length && params.start_date && params.end_date && params.time_lines) {
      dateLines = parseConcertRangeTimesFinal_(params.start_date, params.end_date, params.time_lines);
    }
    if (!dateLines.length) throw new Error('date_lines or start_date/end_date/time_lines is required.');

    const concertRows = readSheetObjectsFinal_(CARRYGO_SHEETS.CONCERTS, CONCERTS_HEADERS);
    const dateRows = readSheetObjectsFinal_(CARRYGO_SHEETS.CONCERT_DATES, CONCERT_DATES_HEADERS);
    const concertId = String(params.concert_id || '').trim() || uniqueConcertIdFinal_(title, concertRows);
    const concertCode = String(params.concert_code || '').trim().toUpperCase() || uniqueConcertCodeFinal_(title, concertRows);
    if (concertRows.some(row => String(row.concert_id) === concertId)) throw new Error('concert_id already exists: ' + concertId);
    if (concertRows.some(row => String(row.concert_code).toUpperCase() === concertCode)) throw new Error('concert_code already exists: ' + concertCode);

    const concert = {
      concert_id: concertId,
      concert_code: concertCode,
      concert_title: title,
      venue: venue,
      city: city,
      is_active: true,
      sort_order: normalizeCountFinal_(params.sort_order, concertRows.length + 1),
      created_at: now,
      updated_at: now
    };
    getSheetFinal_(CARRYGO_SHEETS.CONCERTS).appendRow(CONCERTS_HEADERS.map(h => concert[h] !== undefined ? concert[h] : ''));

    const shDates = getSheetFinal_(CARRYGO_SHEETS.CONCERT_DATES);
    const createdDates = [];
    dateLines.forEach((item, idx) => {
      const concertDateId = uniqueConcertDateIdFinal_(concertId, item.date, item.time, dateRows.concat(createdDates));
      const obj = {
        concert_date_id: concertDateId,
        concert_id: concertId,
        concert_date: item.date,
        concert_time: item.time,
        pickup_time_options: normalizePickupTimeOptionsFinal_(params.pickup_time_options).join(','),
        pickup_drop_guide_link: String(params.pickup_drop_guide_link || '').trim(),
        next_day_pickup_guide_link: String(params.next_day_pickup_guide_link || '').trim(),
        location_change_guide_link: String(params.location_change_guide_link || '').trim(),
        is_active: true,
        sort_order: dateRows.length + idx + 1,
        created_at: now,
        updated_at: now
      };
      shDates.appendRow(CONCERT_DATES_HEADERS.map(h => obj[h] !== undefined ? obj[h] : ''));
      createdDates.push(obj);
    });
    return jsonFinal_({ ok: true, concert: concert, concert_dates: createdDates });
  } catch (err) {
    return jsonFinal_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function adminSetActiveApiFinal_(params) {
  try {
    validateAdminPinFinal_(params.admin_pin);
    const kind = String(params.kind || '').trim();
    const id = String(params.id || '').trim();
    const isActive = String(params.is_active || '').toUpperCase() === 'TRUE';
    if (!id) throw new Error('id is required');
    let sheetName, headers, idHeader;
    if (kind === 'concert') {
      sheetName = CARRYGO_SHEETS.CONCERTS; headers = CONCERTS_HEADERS; idHeader = 'concert_id';
    } else if (kind === 'concert_date') {
      sheetName = CARRYGO_SHEETS.CONCERT_DATES; headers = CONCERT_DATES_HEADERS; idHeader = 'concert_date_id';
    } else {
      throw new Error('Invalid kind: ' + kind);
    }
    const sh = getSheetFinal_(sheetName);
    const values = sh.getDataRange().getValues();
    const idCol = headers.indexOf(idHeader);
    const activeCol = headers.indexOf('is_active') + 1;
    const updatedCol = headers.indexOf('updated_at') + 1;
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][idCol]) === id) {
        sh.getRange(i + 1, activeCol).setValue(isActive);
        if (updatedCol > 0) sh.getRange(i + 1, updatedCol).setValue(new Date());
        return jsonFinal_({ ok: true, kind: kind, id: id, is_active: isActive });
      }
    }
    throw new Error('Not found: ' + id);
  } catch (err) {
    return jsonFinal_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}


function adminConcertsApiFinal_(params) {
  try {
    validateAdminPinFinal_(params.admin_pin);
    return jsonFinal_({
      ok: true,
      concerts: readSheetObjectsFinal_(CARRYGO_SHEETS.CONCERTS, CONCERTS_HEADERS),
      concert_dates: readSheetObjectsFinal_(CARRYGO_SHEETS.CONCERT_DATES, CONCERT_DATES_HEADERS)
    });
  } catch (err) {
    return jsonFinal_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function adminReservationSummaryFinal_(row) {
  return {
    reservation_id: row.reservation_id,
    status: row.status,
    customer_name: row.customer_name,
    customer_email: row.customer_email,
    concert_title: row.concert_title,
    concert_date: row.concert_date,
    concert_time: row.concert_time,
    pickup_time: row.pickup_time,
    booking_channel: row.booking_channel,
    luggage_tag_numbers: row.luggage_tag_numbers,
    onsite_due_amount: row.onsite_due_amount,
    onsite_cash_received: row.onsite_cash_received,
    onsite_tag_attached: row.onsite_tag_attached,
    onsite_photo_taken: row.onsite_photo_taken,
    onsite_checkin_completed_at: formatDateTimeMaybeFinal_(row.onsite_checkin_completed_at),
    actual_suitcase_count: row.actual_suitcase_count,
    actual_extra_bag_count: row.actual_extra_bag_count,
    expected_suitcase_count: row.expected_suitcase_count,
    expected_extra_bag_count: row.expected_extra_bag_count,
    payment_method: row.payment_method,
    payment_status: row.payment_status,
    paid_at: formatDateTimeMaybeFinal_(row.paid_at),
    amount_display: String(row.currency || '').toUpperCase() === 'USD' ? '$' + row.base_fee : '₩' + Number(row.base_fee || 0).toLocaleString('ko-KR'),
    payment_due_at: formatDateTimeMaybeFinal_(row.payment_due_at),
    picked_up_at: formatDateTimeMaybeFinal_(row.picked_up_at),
    picked_up_by: row.picked_up_by,
    returned_at: formatDateTimeMaybeFinal_(row.returned_at),
    cancelled_at: formatDateTimeMaybeFinal_(row.cancelled_at),
    refunded_at: formatDateTimeMaybeFinal_(row.refunded_at)
  };
}

function splitIdsFinal_(value) {
  return String(value || '').split(',').map(id => id.trim()).filter(Boolean);
}

function mergeNoteFinal_(oldNote, newNote) {
  const oldText = String(oldNote || '').trim();
  const newText = String(newNote || '').trim();
  if (!oldText) return newText;
  if (!newText) return oldText;
  return oldText + '\n' + newText;
}
function slugFinal_(text) {
  return String(text || '')
    .toLowerCase()
    .replace(/[^a-z0-9가-힣]+/g, '_')
    .replace(/^_+|_+$/g, '')
    .slice(0, 40) || 'concert';
}

function romanCodeFinal_(text) {
  const words = String(text || '').toUpperCase().match(/[A-Z0-9]+/g) || [];
  if (!words.length) return 'CG';
  if (words.length === 1) return words[0].slice(0, 2).padEnd(2, 'X');
  return (words[0][0] + words[1][0]).slice(0, 2);
}

function uniqueConcertIdFinal_(title, rows) {
  const base = slugFinal_(title);
  let id = base;
  let n = 2;
  while (rows.some(row => String(row.concert_id) === id)) id = base + '_' + n++;
  return id;
}

function uniqueConcertCodeFinal_(title, rows) {
  const base = romanCodeFinal_(title);
  let code = base;
  let n = 2;
  while (rows.some(row => String(row.concert_code).toUpperCase() === code)) code = base + String(n++).padStart(2, '0');
  return code;
}

function uniqueConcertDateIdFinal_(concertId, date, time, rows) {
  const dateKey = String(date || '').replace(/[^0-9]/g, '');
  const timeKey = String(time || '').replace(/[^0-9]/g, '').slice(0, 4) || '0000';
  const base = slugFinal_(concertId) + '_' + dateKey + '_' + timeKey;
  let id = base;
  let n = 2;
  while (rows.some(row => String(row.concert_date_id) === id)) id = base + '_' + n++;
  return id;
}

function parseConcertDateLinesFinal_(value) {
  return String(value || '').split(/\n+/).map(line => line.trim()).filter(Boolean).map(line => {
    const m = line.match(/^(\d{4}-\d{2}-\d{2})\s+([0-2]?\d:[0-5]\d)$/);
    if (!m) throw new Error('Invalid date line: ' + line + ' / Use YYYY-MM-DD HH:MM');
    return { date: m[1], time: m[2].padStart(5, '0') };
  });
}

function adminDeleteConcertApiFinal_(params) {
  try {
    validateAdminPinFinal_(params.admin_pin);
    const concertId = String(params.concert_id || '').trim();
    if (!concertId) throw new Error('concert_id is required');
    const reservationRows = readSheetObjectsFinal_(CARRYGO_SHEETS.RESERVATIONS, RESERVATIONS_HEADERS);
    if (reservationRows.some(row => String(row.concert_id) === concertId)) {
      throw new Error('Cannot delete concert with reservations. Disable it instead.');
    }
    deleteSheetRowsByValueFinal_(CARRYGO_SHEETS.CONCERT_DATES, CONCERT_DATES_HEADERS, 'concert_id', concertId);
    const deleted = deleteSheetRowsByValueFinal_(CARRYGO_SHEETS.CONCERTS, CONCERTS_HEADERS, 'concert_id', concertId);
    return jsonFinal_({ ok: true, deleted: deleted, concert_id: concertId });
  } catch (err) {
    return jsonFinal_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function adminDeleteConcertDateApiFinal_(params) {
  try {
    validateAdminPinFinal_(params.admin_pin);
    const concertDateId = String(params.concert_date_id || '').trim();
    if (!concertDateId) throw new Error('concert_date_id is required');
    const reservationRows = readSheetObjectsFinal_(CARRYGO_SHEETS.RESERVATIONS, RESERVATIONS_HEADERS);
    if (reservationRows.some(row => String(row.concert_date_id) === concertDateId)) {
      throw new Error('Cannot delete concert date with reservations. Disable it instead.');
    }
    const deleted = deleteSheetRowsByValueFinal_(CARRYGO_SHEETS.CONCERT_DATES, CONCERT_DATES_HEADERS, 'concert_date_id', concertDateId);
    return jsonFinal_({ ok: true, deleted: deleted, concert_date_id: concertDateId });
  } catch (err) {
    return jsonFinal_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function deleteSheetRowsByValueFinal_(sheetName, headers, field, value) {
  const sh = getSheetFinal_(sheetName);
  const values = sh.getDataRange().getValues();
  const col = headers.indexOf(field);
  if (col < 0) throw new Error('Header not found: ' + field);
  let deleted = 0;
  for (let i = values.length - 1; i >= 1; i--) {
    if (String(values[i][col]) === String(value)) {
      sh.deleteRow(i + 1);
      deleted++;
    }
  }
  return deleted;
}
function parseConcertRangeTimesFinal_(startDate, endDate, timeLines) {
  const start = parseIsoDateFinal_(startDate);
  const end = parseIsoDateFinal_(endDate);
  if (end.getTime() < start.getTime()) throw new Error('end_date must be after start_date');
  const dates = [];
  for (let d = new Date(start.getTime()); d.getTime() <= end.getTime(); d.setDate(d.getDate() + 1)) {
    dates.push(Utilities.formatDate(d, CARRYGO_TIMEZONE, 'yyyy-MM-dd'));
  }
  const times = String(timeLines || '').split(/\n+/).map(line => line.trim()).filter(Boolean);
  if (times.length !== dates.length) throw new Error('time_lines count must match date range days. Dates: ' + dates.length + ', times: ' + times.length);
  return dates.map((date, idx) => {
    const raw = times[idx];
    const m = raw.match(/^([0-2]?\d:[0-5]\d)$/);
    if (!m) throw new Error('Invalid time line: ' + raw + ' / Use HH:MM');
    return { date: date, time: m[1].padStart(5, '0') };
  });
}

function parseIsoDateFinal_(value) {
  const text = String(value || '').trim();
  const m = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) throw new Error('Invalid date: ' + value + ' / Use YYYY-MM-DD');
  return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
}

function adminUpdateConcertApiFinal_(params) {
  try {
    validateAdminPinFinal_(params.admin_pin);
    const concertId = String(params.concert_id || '').trim();
    if (!concertId) throw new Error('concert_id is required');
    const sh = getSheetFinal_(CARRYGO_SHEETS.CONCERTS);
    const values = sh.getDataRange().getValues();
    const idCol = CONCERTS_HEADERS.indexOf('concert_id');
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][idCol]) === concertId) {
        updateCellIfParamFinal_(sh, i + 1, CONCERTS_HEADERS, 'concert_code', params.concert_code, true);
        updateCellIfParamFinal_(sh, i + 1, CONCERTS_HEADERS, 'concert_title', params.concert_title, true);
        updateCellIfParamFinal_(sh, i + 1, CONCERTS_HEADERS, 'venue', params.venue, true);
        updateCellIfParamFinal_(sh, i + 1, CONCERTS_HEADERS, 'city', String(params.city || 'SEOUL').trim().toUpperCase(), true);
        updateCellIfParamFinal_(sh, i + 1, CONCERTS_HEADERS, 'updated_at', new Date(), false);
        return jsonFinal_({ ok: true, concert_id: concertId });
      }
    }
    throw new Error('concert not found: ' + concertId);
  } catch (err) {
    return jsonFinal_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function adminUpdateConcertDateApiFinal_(params) {
  try {
    validateAdminPinFinal_(params.admin_pin);
    const concertDateId = String(params.concert_date_id || '').trim();
    if (!concertDateId) throw new Error('concert_date_id is required');
    const sh = getSheetFinal_(CARRYGO_SHEETS.CONCERT_DATES);
    const values = sh.getDataRange().getValues();
    const idCol = CONCERT_DATES_HEADERS.indexOf('concert_date_id');
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][idCol]) === concertDateId) {
        updateCellIfParamFinal_(sh, i + 1, CONCERT_DATES_HEADERS, 'concert_date', params.concert_date, true);
        updateCellIfParamFinal_(sh, i + 1, CONCERT_DATES_HEADERS, 'concert_time', params.concert_time, true);
        if (params.pickup_time_options !== undefined) updateCellIfParamFinal_(sh, i + 1, CONCERT_DATES_HEADERS, 'pickup_time_options', normalizePickupTimeOptionsFinal_(params.pickup_time_options).join(','), false);
        updateCellIfParamFinal_(sh, i + 1, CONCERT_DATES_HEADERS, 'pickup_drop_guide_link', params.pickup_drop_guide_link, true);
        updateCellIfParamFinal_(sh, i + 1, CONCERT_DATES_HEADERS, 'next_day_pickup_guide_link', params.next_day_pickup_guide_link, true);
        updateCellIfParamFinal_(sh, i + 1, CONCERT_DATES_HEADERS, 'location_change_guide_link', params.location_change_guide_link, true);
        updateCellIfParamFinal_(sh, i + 1, CONCERT_DATES_HEADERS, 'updated_at', new Date(), false);
        return jsonFinal_({ ok: true, concert_date_id: concertDateId });
      }
    }
    throw new Error('concert_date not found: ' + concertDateId);
  } catch (err) {
    return jsonFinal_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function updateCellIfParamFinal_(sh, rowNo, headers, field, value, skipUndefined) {
  if (skipUndefined && (value === undefined || value === null)) return;
  const col = headers.indexOf(field) + 1;
  if (col < 1) throw new Error('Header not found: ' + field);
  sh.getRange(rowNo, col).setValue(value);
}

