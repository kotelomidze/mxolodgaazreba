/**
 * საგამოცდო ნაშრომების მიმღები — Google Apps Script.
 *
 * გაშვება (ერთხელ):
 *   1. script.google.com → New project
 *   2. წაშალე იქ არსებული კოდი და ჩასვი ეს ფაილი მთლიანად
 *   3. Deploy → New deployment → ტიპი: Web app
 *        Execute as:      Me
 *        Who has access:  Anyone
 *   4. დააკოპირე მიღებული URL და ჩასვი config.json-ის webhookUrl ველში
 *
 * თუ ცხრილშიც გინდა ჩანაწერების დაგროვება, შექმენი Google Sheet,
 * მისამართიდან ამოიღე ID და ჩასვი SHEET_ID-ში. სურვილისამებრია.
 */

var RECIPIENT = 'lomidze.kote1@gmail.com';
var SHEET_ID = '';

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var name = ((data.student && data.student.firstName || '') + ' ' +
                (data.student && data.student.lastName || '')).trim() || 'უცნობი მოსწავლე';

    var answers = data.comprehensionAnswers || {};
    var key = data.correctAnswers || {};
    var numbers = Object.keys(answers).concat(Object.keys(key))
      .filter(function (v, i, a) { return a.indexOf(v) === i; })
      .sort(function (a, b) { return Number(a) - Number(b); });

    var correct = 0;
    var keyKnown = 0;
    var rows = [];

    numbers.forEach(function (n) {
      var given = answers[n] || '—';
      var right = key[n] || '';
      var state = 'unknown';
      if (right) {
        keyKnown++;
        state = (given === right) ? 'ok' : 'bad';
        if (state === 'ok') correct++;
      }
      rows.push({ n: n, given: given, right: right, state: state });
    });

    var score = (data.comprehensionTotal != null) ? data.comprehensionTotal : correct;
    var maxScore = data.comprehensionMaxPoints || numbers.length;

    var scoreText;
    if (keyKnown === numbers.length && numbers.length > 0) {
      scoreText = score + ' / ' + maxScore;
    } else if (keyKnown > 0) {
      scoreText = correct + ' / ' + keyKnown + ' (ნაწილობრივი — გასაღები არასრულია)';
    } else {
      scoreText = 'ვერ დაითვალა — სწორი პასუხები შევსებული არ არის';
    }

    var duration = formatDuration(data.durationSeconds);

    // --- უბრალო ტექსტი (სათადარიგოდ) ---
    var plain = [];
    plain.push('მოსწავლე: ' + name);
    plain.push('ტექსტი: ' + (data.testTitle || data.testId || '—'));
    plain.push('გაგზავნის დრო: ' + formatDate(data.submittedAt));
    plain.push('ხანგრძლივობა: ' + duration);
    plain.push('');
    plain.push('წაკითხულის გააზრება — ' + scoreText);
    rows.forEach(function (r) {
      plain.push('  ' + r.n + '. ' + r.given +
        (r.state === 'ok' ? '  ✓' : r.state === 'bad' ? '  ✗ (სწორი: ' + r.right + ')' : ''));
    });

    // --- HTML: ✓ მწვანე, ✗ წითელი ---
    var html = [];
    html.push('<div style="font-family:Arial,\'Noto Sans Georgian\',sans-serif;font-size:14px;color:#0f172a;max-width:640px">');
    html.push('<p style="margin:0 0 4px"><b>მოსწავლე:</b> ' + esc(name) + '</p>');
    html.push('<p style="margin:0 0 4px"><b>ტექსტი:</b> ' + esc(data.testTitle || data.testId || '—') + '</p>');
    html.push('<p style="margin:0 0 4px"><b>გაგზავნის დრო:</b> ' + esc(formatDate(data.submittedAt)) + '</p>');
    html.push('<p style="margin:0 0 16px"><b>ხანგრძლივობა:</b> ' + esc(duration) + '</p>');

    html.push('<h3 style="margin:0 0 8px;font-size:15px;border-top:1px solid #e2e8f0;padding-top:12px">' +
              'წაკითხულის გააზრება — ' + esc(scoreText) + '</h3>');
    html.push('<table style="border-collapse:collapse;font-size:14px">');
    rows.forEach(function (r) {
      var color = r.state === 'ok' ? '#16a34a' : r.state === 'bad' ? '#dc2626' : '#64748b';
      var mark = r.state === 'ok' ? '✓' : r.state === 'bad' ? '✗' : '–';
      var extra = r.state === 'bad' ? ' <span style="color:#64748b">(სწორი: ' + esc(r.right) + ')</span>' : '';
      html.push('<tr>' +
        '<td style="padding:3px 10px 3px 0;color:#64748b">' + esc(r.n) + '.</td>' +
        '<td style="padding:3px 10px 3px 0;font-weight:bold">' + esc(r.given) + '</td>' +
        '<td style="padding:3px 10px 3px 0;color:' + color + ';font-weight:bold;font-size:16px">' + mark + '</td>' +
        '<td style="padding:3px 0">' + extra + '</td></tr>');
    });
    html.push('</table>');

    html.push('</div>');

    MailApp.sendEmail({
      to: RECIPIENT,
      subject: 'ნაშრომი: ' + name + ' — ' + (data.testTitle || data.testId || ''),
      body: plain.join('\n'),
      htmlBody: html.join('')
    });

    if (SHEET_ID) {
      var sheet = SpreadsheetApp.openById(SHEET_ID).getSheets()[0];
      if (sheet.getLastRow() === 0) {
        sheet.appendRow(['თარიღი', 'მოსწავლე', 'ტექსტი', 'ქულა', 'მაქსიმუმი', 'ხანგრძლივობა']);
      }
      sheet.appendRow([
        new Date(), name, data.testTitle || data.testId,
        score, maxScore, duration
      ]);
    }

    return json({ status: 'success' });

  } catch (error) {
    // შეცდომაც იმეილით მოვა, რომ ნაშრომი უკვალოდ არ დაიკარგოს
    try {
      MailApp.sendEmail(RECIPIENT, 'ნაშრომის მიღების შეცდომა',
        String(error) + '\n\n' + (e && e.postData ? e.postData.contents : '(მონაცემები არ არის)'));
    } catch (ignored) {}
    return json({ status: 'error', error: String(error) });
  }
}

function formatDate(iso) {
  try {
    return Utilities.formatDate(new Date(iso), 'Asia/Tbilisi', 'dd.MM.yyyy HH:mm');
  } catch (e) {
    return iso || '';
  }
}

function formatDuration(seconds) {
  var s = Number(seconds);
  if (!s || s < 0) return 'უცნობია';
  var h = Math.floor(s / 3600);
  var m = Math.floor((s % 3600) / 60);
  var sec = s % 60;
  if (h) return h + ' სთ ' + m + ' წთ';
  if (m) return m + ' წთ ' + sec + ' წმ';
  return sec + ' წმ';
}

function esc(v) {
  return String(v == null ? '' : v)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * შესამოწმებლად: გაუშვი ეს ფუნქცია რედაქტორიდან (Run ღილაკი).
 * პირველად Google ნებართვას მოგთხოვს. თუ იმეილი მოვიდა, ყველაფერი წესრიგშია.
 */
function testEmail() {
  doPost({
    postData: {
      contents: JSON.stringify({
        student: { firstName: 'ტესტი', lastName: 'ტესტიშვილი' },
        testTitle: 'შოთა რუსთაველი – „ვეფხისტყაოსანი“',
        submittedAt: new Date().toISOString(),
        comprehensionAnswers: { 3: 'ა', 4: 'ბ', 5: 'გ' },
        correctAnswers: { 3: 'ა', 4: 'ბ', 5: 'დ' },
        comprehensionTotal: 2,
        comprehensionMaxPoints: 15,
        durationSeconds: 1875
      })
    }
  });
}
