/**
 * Ripple Rwanda booking and customer register.
 *
 * Deployment:
 * 1. Replace the code in the existing Apps Script project with this file.
 * 2. In Project Settings > Script properties, add RIPPLE_ADMIN_PIN with a strong private PIN.
 * 3. Deploy > Manage deployments > Edit > New version > Deploy.
 * 4. Keep the existing web-app deployment URL so the website endpoint remains unchanged.
 *
 * The public website may submit Website bookings through doPost().
 * Staff record WhatsApp bookings through the private-PIN protected ?view=whatsapp portal.
 */

const RIPPLE_BOOKINGS_SHEET = 'Bookings';
const RIPPLE_CUSTOMERS_SHEET = 'Customers';
const RIPPLE_ACTIVITY_SHEET = 'Activity';

const BOOKING_HEADERS = [
  'received', 'id', 'hub', 'zone', 'slot', 'size', 'fare',
  'vendor', 'rec', 'drop', 'customerId', 'phone', 'source', 'status'
];
const CUSTOMER_HEADERS = [
  'customerId', 'phone', 'name', 'firstSeen', 'lastSeen', 'totalBookings'
];
const ACTIVITY_HEADERS = ['timestamp', 'bookingId', 'customerId', 'channel', 'event', 'notes'];

function doGet(e) {
  if (e && e.parameter && e.parameter.view === 'whatsapp') {
    return HtmlService.createHtmlOutput(getWhatsAppPortalHtml())
      .setTitle('Ripple Rwanda | WhatsApp Booking Intake')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  return ContentService.createTextOutput('Ripple Rwanda booking endpoint is ready.');
}

function doPost(e) {
  try {
    const payload = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const result = saveBooking_(payload, 'Website');
    return ContentService.createTextOutput(JSON.stringify({ ok: true, result: result }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    console.error(error);
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: String(error.message || error) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/** Called by the Apps Script-hosted staff portal. */
function saveWhatsAppBooking(payload) {
  const expectedPin = PropertiesService.getScriptProperties().getProperty('RIPPLE_ADMIN_PIN');
  if (!expectedPin) throw new Error('Admin PIN is not configured. Ask the script owner to set RIPPLE_ADMIN_PIN in Script Properties.');
  if (!payload || String(payload.pin || '') !== expectedPin) throw new Error('Incorrect staff PIN.');
  return saveBooking_(payload, 'WhatsApp');
}

function saveBooking_(payload, defaultSource) {
  const phone = normalizePhone_(clean_(payload.phone || payload.vendorPhone || '', 30));
  if (!phone) throw new Error('Customer phone number is required.');
  const name = clean_(payload.vendor || payload.name || '', 100);
  if (!name) throw new Error('Customer name is required.');

  const customerId = customerIdForPhone_(phone);
  const bookingId = clean_(payload.id || ('RPL-' + Utilities.getUuid().slice(0, 8).toUpperCase()), 40);
  const now = new Date();
  const record = {
    received: now,
    id: bookingId,
    hub: clean_(payload.hub || '', 150),
    zone: clean_(payload.zone || '', 150),
    slot: clean_(payload.slot || '', 40),
    size: clean_(payload.size || '', 80),
    fare: clean_(payload.fare || '', 40),
    vendor: name,
    rec: clean_(payload.rec || '', 100),
    drop: clean_(payload.drop || '', 150),
    customerId: customerId,
    phone: phone,
    source: defaultSource,
    status: clean_(payload.status || 'New - needs manual confirmation', 80)
  };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bookings = ensureSheet_(ss, RIPPLE_BOOKINGS_SHEET, BOOKING_HEADERS);
  appendObjectRow_(bookings, BOOKING_HEADERS, record);

  const customers = ensureSheet_(ss, RIPPLE_CUSTOMERS_SHEET, CUSTOMER_HEADERS);
  upsertCustomer_(customers, customerId, phone, name, now);

  const activity = ensureSheet_(ss, RIPPLE_ACTIVITY_SHEET, ACTIVITY_HEADERS);
  appendObjectRow_(activity, ACTIVITY_HEADERS, {
    timestamp: now,
    bookingId: bookingId,
    customerId: customerId,
    channel: defaultSource,
    event: 'Booking recorded',
    notes: record.status
  });

  return { bookingId: bookingId, customerId: customerId, source: defaultSource, status: record.status };
}

function ensureSheet_(ss, name, requiredHeaders) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
    sheet.setFrozenRows(1);
    return sheet;
  }
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const current = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
  const missing = requiredHeaders.filter(h => current.indexOf(h) === -1);
  if (missing.length) sheet.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
  sheet.setFrozenRows(1);
  return sheet;
}

function appendObjectRow_(sheet, headers, obj) {
  const actualHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  const row = actualHeaders.map(header => Object.prototype.hasOwnProperty.call(obj, header) ? obj[header] : '');
  sheet.appendRow(row);
}

function upsertCustomer_(sheet, customerId, phone, name, now) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    sheet.appendRow([customerId, phone, name, now, now, 1]);
    return;
  }
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(String);
  const idCol = headers.indexOf('customerId') + 1;
  const phoneCol = headers.indexOf('phone') + 1;
  const nameCol = headers.indexOf('name') + 1;
  const firstCol = headers.indexOf('firstSeen') + 1;
  const lastCol = headers.indexOf('lastSeen') + 1;
  const countCol = headers.indexOf('totalBookings') + 1;
  const ids = sheet.getRange(2, idCol, lastRow - 1, 1).getDisplayValues().flat();
  const found = ids.indexOf(customerId);
  if (found === -1) {
    const row = new Array(headers.length).fill('');
    row[idCol - 1] = customerId;
    row[phoneCol - 1] = phone;
    row[nameCol - 1] = name;
    row[firstCol - 1] = now;
    row[lastCol - 1] = now;
    row[countCol - 1] = 1;
    sheet.appendRow(row);
  } else {
    const rowNumber = found + 2;
    sheet.getRange(rowNumber, phoneCol).setValue(phone);
    sheet.getRange(rowNumber, nameCol).setValue(name);
    sheet.getRange(rowNumber, lastCol).setValue(now);
    sheet.getRange(rowNumber, countCol).setValue(Number(sheet.getRange(rowNumber, countCol).getValue() || 0) + 1);
  }
}

function normalizePhone_(phone) {
  let digits = String(phone || '').replace(/[^0-9]/g, '');
  if (digits.startsWith('00')) digits = digits.slice(2);
  if (digits.startsWith('0') && digits.length === 10) digits = '250' + digits.slice(1);
  else if (digits.length === 9 && digits.startsWith('7')) digits = '250' + digits;
  return digits ? '+' + digits : '';
}

function customerIdForPhone_(phone) {
  const normalized = String(phone).replace(/[^0-9]/g, '');
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, normalized);
  const hex = digest.map(b => ('0' + ((b + 256) % 256).toString(16)).slice(-2)).join('').toUpperCase();
  return 'RPL-C-' + hex.slice(0, 8);
}

function clean_(value, maxLength) {
  let text = String(value == null ? '' : value).replace(/[\r\n\t]/g, ' ').trim().slice(0, maxLength || 300);
  if (/^[=+@]/.test(text)) text = "'" + text;
  return text;
}

function getWhatsAppPortalHtml() {
  return `<!doctype html><html><head><base target="_top"><meta name="viewport" content="width=device-width,initial-scale=1">
  <style>
    *{box-sizing:border-box}body{margin:0;background:#f1f5f9;color:#0f172a;font:16px system-ui,sans-serif}
    main{max-width:760px;margin:28px auto;padding:24px;background:white;border:1px solid #e2e8f0;border-radius:20px}
    h1{margin:0 0 8px;color:#065f46;font-size:26px}.muted{color:#64748b;font-size:14px}
    form{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-top:22px}
    label{display:block;font-size:13px;font-weight:700;margin-bottom:6px}
    input,select{width:100%;padding:12px;border:1px solid #cbd5e1;border-radius:10px;font:inherit;background:#fff}
    .wide{grid-column:1/-1}button{border:0;background:#047857;color:white;padding:14px;border-radius:10px;font-weight:800;font-size:15px;cursor:pointer}
    #message{padding:12px;border-radius:10px;margin-top:14px;display:none}.ok{background:#dcfce7;color:#166534}.err{background:#fee2e2;color:#991b1b}
    @media(max-width:600px){main{margin:12px;padding:18px}form{grid-template-columns:1fr}.wide{grid-column:auto}}
  </style></head><body><main>
    <h1>Ripple Rwanda | WhatsApp booking intake</h1>
    <p class="muted">Staff only. Enter details from a customer's WhatsApp message. A customer is matched by phone number, so repeat bookings update the same customer profile.</p>
    <form id="f">
      <div class="wide"><label>Staff PIN *</label><input name="pin" type="password" required autocomplete="current-password"></div>
      <div><label>Customer name *</label><input name="name" required maxlength="100"></div>
      <div><label>Customer phone *</label><input name="phone" required placeholder="+2507XXXXXXXX" maxlength="30"></div>
      <div><label>Pickup market / hub *</label><select name="hub" required><option value="">Choose pickup</option><option>Kimironko Market (KG 145 St)</option><option>Nyabugogo Main Market</option><option>Downtown / Nyarugenge Market</option><option>Remera / Gisimenti</option><option>Kicukiro Centre</option></select></div>
      <div><label>Destination zone *</label><select name="zone" required><option value="">Choose destination</option><option>Remera / Kimironko / Bibare (Zone A)</option><option>Kacyiru / Kimihurura / CBD (Zone B)</option><option>Nyamirambo / Nyabugogo / Gikondo (Zone C)</option><option>Kicukiro / Kanombe / Masaka (Zone D)</option></select></div>
      <div><label>Parcel size *</label><select name="size" required><option value="">Choose size</option><option value="Small Parcel / Shoebox">Small Parcel / Shoebox</option><option value="Medium Box / Bag">Medium Box / Bag</option><option value="Large Sack / Crate">Large Sack / Crate</option><option value="Bulk Mini-Van">Bulk Mini-Van</option></select></div>
      <div><label>Departure slot</label><select name="slot"><option>10:00 AM</option><option>1:00 PM</option><option>4:00 PM</option><option>To be confirmed</option></select></div>
      <div><label>Fare (RWF)</label><input name="fare" placeholder="Confirm with customer"></div>
      <div><label>Recipient name</label><input name="rec"></div>
      <div><label>Delivery location</label><input name="drop"></div>
      <div><label>Status</label><select name="status"><option>New - needs manual confirmation</option><option>Confirmed</option><option>Cancelled</option><option>Needs follow-up</option></select></div>
      <div class="wide"><button type="submit">Save WhatsApp booking to Google Sheets</button></div>
    </form>
    <div id="message" role="status" aria-live="polite"></div>
    <p class="muted">Save only details needed to arrange the parcel. Do not enter passwords, payment card details, or unrelated personal information.</p>
  </main><script>
    const form=document.getElementById('f'),msg=document.getElementById('message');
    form.addEventListener('submit',function(ev){
      ev.preventDefault();const data=Object.fromEntries(new FormData(form).entries());
      msg.style.display='block';msg.className='';msg.textContent='Saving booking…';
      google.script.run.withSuccessHandler(function(r){
        msg.className='ok';msg.textContent='Saved. Booking: '+r.bookingId+' | Customer: '+r.customerId+' | Channel: WhatsApp';
        const pin=data.pin;form.reset();form.elements.pin.value=pin;
      }).withFailureHandler(function(err){
        msg.className='err';msg.textContent=err&&err.message?err.message:'Could not save. Please try again.';
      }).saveWhatsAppBooking(data);
    });
  </script></body></html>`;
}
