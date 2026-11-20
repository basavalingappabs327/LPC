// Google Apps Script – paste into script.google.com (Code.gs)
// TEMPLATE_ID = ID of the Google Doc made from LPC_template.docx (…/document/d/<ID>/edit)
const TEMPLATE_ID = '1wJPWuIV2iRJHTBItNgq2cE2Lyg3WmShBbXpnTRPeH-Q';

function doGet() { return json({ok: true, msg: 'LPC service is running'}); }

function doPost(e) {
  let copy;
  try {
    const req = JSON.parse(e.postData.contents);
    const data = req.data || {};
    const format = req.format === 'docx' ? 'docx' : 'pdf';

    copy = DriveApp.getFileById(TEMPLATE_ID).makeCopy('LPC_tmp_' + Date.now());
    const id = copy.getId();
    const doc = DocumentApp.openById(id);
    const body = doc.getBody();

    // Remove unused extra pay/deduction rows ({xe1}/{xd1} …) so no blank rows remain
    const tables = body.getTables();
    if (tables.length > 0) {
      const t0 = tables[0];
      for (let r = t0.getNumRows() - 1; r >= 0; r--) {
        const m = t0.getRow(r).getCell(1).getText().match(/\{xe(\d+)\}/);
        if (m && !data['xe' + m[1]] && !data['xd' + m[1]]) t0.removeRow(r);
      }
    }

    // Replace every {token} / {{token}}. Key = token with braces, spaces, commas removed.
    // Skip {pagebreak} — handled separately as a real page break.
    const tokens = {};
    (body.getText().match(/\{+[^{}]*\}+/g) || []).forEach(function (t) {
      if (!/pagebreak/i.test(t)) tokens[t] = true;
    });
    Object.keys(tokens).forEach(function (tok) {
      const key = tok.replace(/[{}\s,]/g, '');
      const val = String(data[key] == null ? '' : data[key]).replace(/\\/g, '\\\\').replace(/\$/g, '\\$');
      body.replaceText(tok.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), val);
    });

    // Turn {pagebreak} into a real page break
    insertPageBreaks(body);

    doc.saveAndClose();

    let blob;
    if (format === 'pdf') {
      blob = copy.getAs('application/pdf');
    } else {
      blob = UrlFetchApp.fetch('https://docs.google.com/document/d/' + id + '/export?format=docx',
        {headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()}}).getBlob();
    }
    const who = String(data.employee || data.gpf || 'form').replace(/[^\w\-]+/g, '_');
    return json({
      ok: true, filename: 'LPC_' + who + '.' + format,
      mime: format === 'pdf' ? 'application/pdf' : 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      data: Utilities.base64Encode(blob.getBytes())
    });
  } catch (err) {
    return json({ok: false, error: String(err)});
  } finally {
    if (copy) copy.setTrashed(true); // each request uses its own temp copy → safe for many users
  }
}

/** Find {pagebreak} markers, remove them, and insert a real page break at that position. */
function insertPageBreaks(body) {
  const marker = '{pagebreak}';
  const escaped = marker.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  let range = body.findText(escaped);
  while (range) {
    const el = range.getElement();
    const start = range.getStartOffset();
    const end = range.getEndOffsetInclusive();
    // Delete the placeholder text
    if (el.editAsText) {
      el.asText().deleteText(start, end);
    }
    // Insert page break just before the paragraph that held the marker
    const parent = el.getParent();
    try {
      const idx = body.getChildIndex(parent);
      body.insertPageBreak(idx);
    } catch (err) {
      body.appendPageBreak();
    }
    range = body.findText(escaped);
  }
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
