/**
 * スプレッドシートを開いた時にカスタムメニューを追加
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📅 カレンダー連携')
    .addItem('スケジュールをカレンダーに登録', 'syncScheduleToCalendar')
    .addToUi();
}

/**
 * スケジュールをGoogleカレンダーに登録するメイン処理
 */
function syncScheduleToCalendar() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('シート1');
  if (!sheet) {
    SpreadsheetApp.getUi().alert('「シート1」が見つかりませんでした。');
    return;
  }

  // --- カレンダーの指定 ---
  // デフォルト（自身のメインカレンダー）を使用する場合:
  const calendar = CalendarApp.getDefaultCalendar();
  
  // 特定の共有カレンダー等に登録したい場合は、上記をコメントアウトし以下にカレンダーID（メールアドレス形式）を入力してください:
  // const calendar = CalendarApp.getCalendarById('xxx@group.calendar.google.com');

  // データ範囲の取得（3行目以降）
  const lastRow = sheet.getLastRow();
  if (lastRow < 3) {
    SpreadsheetApp.getUi().alert('登録対象のデータがありません。');
    return;
  }

  // B列(2列目)〜H列(8列目)のデータを取得
  const data = sheet.getRange(3, 2, lastRow - 2, 7).getValues();

  // 列の対応定義
  const milestones = [
    { colIdx: 0, title: '【イベント実施】' },         // B列: イベント実施日
    { colIdx: 2, title: '【企画書レビュー】' },       // D列: 企画書レビュー
    { colIdx: 3, title: '【バナー依頼】' },           // E列: バナー依頼
    { colIdx: 4, title: '【バナー受取】' },           // F列: バナー受取
    { colIdx: 5, title: '【イベントページレビュー】' }, // G列: イベントページレビュー
    { colIdx: 6, title: '【公開期日】' }              // H列: イベント公開日期日
  ];

  let createdCount = 0;

  data.forEach((row, index) => {
    const rowNum = index + 3;
    // C列（イベント名）。空欄の場合は実施日などから仮の名称を設定
    const eventName = row[1] ? String(row[1]).trim() : `(名称未定イベント_行${rowNum})`;

    milestones.forEach(m => {
      const dateVal = row[m.colIdx];
      const targetDate = parseDateValue(dateVal);

      if (targetDate) {
        const title = `${m.title} ${eventName}`;
        const description = `スプレッドシート「Findy Studentイベント公開までのスケジュール」より自動登録\n対象行: ${rowNum}行目`;

        // 既存の同一タイトルの重複登録を防ぐチェック
        const existingEvents = calendar.getEventsForDay(targetDate, { search: title });
        const isAlreadyRegistered = existingEvents.some(e => e.getTitle() === title);

        if (!isAlreadyRegistered) {
          calendar.createAllDayEvent(title, targetDate, { description: description });
          createdCount++;
        }
      }
    });
  });

  SpreadsheetApp.getUi().alert(`カレンダー連携が完了しました！\n新しく ${createdCount} 件の予定を登録しました。`);
}

/**
 * セルの値（Date型オブジェクトまたは「2026/9/29 (火)」等の文字列）をDate型に変換する関数
 */
function parseDateValue(val) {
  if (!val) return null;
  if (val instanceof Date && !isNaN(val.getTime())) {
    return val;
  }
  if (typeof val === 'string') {
    // YYYY/MM/DD または YYYY-MM-DD の数字部分を抽出
    const match = val.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
    if (match) {
      return new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
    }
  }
  return null;
}
