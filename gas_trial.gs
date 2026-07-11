// =====================================================
// アスユメ Labo. メンタルスクリーニング【お試し版】GAS
// =====================================================

const TRIAL_SS_ID = '1vNg4jYMMzlTr0M-Rl0jkzbCpilHkJLds6iHiZDmdW_c'; // ← スプレッドシートIDを後で設定

function doGet(e) {
  return ContentService.createTextOutput('ok');
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    if (data.action === 'submit') {
      saveToSheet(data);
      sendEmail(data);
    }
    return ContentService.createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// =====================================================
// シートへ保存
// =====================================================
function saveToSheet(data) {
  const ss = SpreadsheetApp.openById(TRIAL_SS_ID);
  let sheet = ss.getSheetByName('📊 スクリーニング結果');
  if (!sheet) {
    sheet = ss.insertSheet('📊 スクリーニング結果');
    sheet.appendRow(['受診日時','会社名','氏名','部署','年齢','性別','メール','総合スコア','判定','仕事の量','仕事のコントロール','職場サポート','身体的・心理的反応','活気・仕事の満足感']);
    sheet.getRange(1,1,1,14).setFontWeight('bold').setBackground('#1E293B').setFontColor('white');
    sheet.setFrozenRows(1);
  }

  const secs = parseSectionText(data.section_text);
  const row = [
    new Date().toLocaleString('ja-JP'),
    data.company || '',
    data.name || '',
    data.dept || '',
    data.age || '',
    data.gender || '',
    data.email || '',
    data.total_score,
    data.risk_label,
    secs['仕事の量'] || '',
    secs['仕事のコントロール'] || '',
    secs['職場サポート'] || '',
    secs['身体的・心理的反応'] || '',
    secs['活気・仕事の満足感'] || ''
  ];
  sheet.appendRow(row);

  const lastRow = sheet.getLastRow();
  const score = data.total_score;
  const bg = score >= 75 ? '#DCFCE7' : score >= 55 ? '#FEF9C3' : score >= 40 ? '#FFEDD5' : '#FEE2E2';
  sheet.getRange(lastRow, 1, 1, 14).setBackground(bg);
}

function parseSectionText(text) {
  const result = {};
  if (!text) return result;
  text.split('\n').forEach(line => {
    const m = line.match(/^(.+?):\s*(\d+)点?$/);
    if (m) result[m[1].trim()] = parseInt(m[2]);
  });
  return result;
}

// =====================================================
// メール送信
// =====================================================
function sendEmail(data) {
  if (!data.email) return;
  const score = data.total_score;
  const rc = score >= 75 ? '#10b981' : score >= 55 ? '#f59e0b' : score >= 40 ? '#f97316' : '#ef4444';
  const body = `<div style="font-family:'Hiragino Sans',Arial,sans-serif;max-width:600px;margin:0 auto;">
    <div style="background:linear-gradient(135deg,#29ABE2,#1E90D8);padding:24px;text-align:center;border-radius:12px 12px 0 0;">
      <p style="color:white;font-size:20px;font-weight:bold;margin:0;">アスユメ Labo.</p>
      <p style="color:rgba(255,255,255,.8);font-size:12px;margin:6px 0 0;">メンタルスクリーニング結果レポート（お試し版）</p>
    </div>
    <div style="background:white;padding:32px;border:1px solid #e2e8f0;">
      <p style="font-size:16px;color:#1E293B;">${data.company || ''} ${data.name || 'ご回答者'} 様</p>
      <p style="color:#475569;font-size:14px;margin-bottom:20px;">メンタルスクリーニング（お試し版）の結果をお送りします。</p>
      <div style="background:${rc};border-radius:12px;padding:20px;text-align:center;margin:20px 0;">
        <p style="color:rgba(255,255,255,.8);font-size:12px;margin:0;">総合評価</p>
        <p style="color:white;font-size:56px;font-weight:900;margin:4px 0;line-height:1;">${score}</p>
        <p style="color:white;font-size:14px;margin:0;">/ 100点 | <b>${data.risk_label}</b></p>
      </div>
      <p style="background:#f8fafc;padding:14px;border-left:4px solid ${rc};border-radius:8px;font-size:13px;color:#475569;line-height:1.7;">${data.comment || ''}</p>
      <p style="font-weight:bold;color:#0F172A;margin-top:24px;font-size:14px;">セクション別スコア</p>
      <p style="font-size:13px;color:#475569;white-space:pre-line;line-height:2;">${data.section_text || ''}</p>
      <div style="background:#eff6ff;border-radius:10px;padding:16px;margin-top:24px;text-align:center;">
        <p style="font-size:14px;font-weight:700;color:#1a5f7a;margin-bottom:8px;">アスユメ相談室にご相談ください</p>
        <p style="font-size:13px;color:#475569;margin-bottom:4px;">有料版では5領域×10問（計50問）の詳細な分析が受けられます。</p>
      </div>
      <p style="font-size:11px;color:#94A3B8;border-top:1px solid #F1F5F9;padding-top:16px;margin-top:24px;">
        ※本レポートは医療診断ではありません。<br>
        アスユメ Labo. | ${Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy年MM月dd日')}
      </p>
    </div>
  </div>`;
  GmailApp.sendEmail(data.email, '【アスユメ Labo.】メンタルスクリーニング結果（お試し版）', '', { htmlBody: body });
}
