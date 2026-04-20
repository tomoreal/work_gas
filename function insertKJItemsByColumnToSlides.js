function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('GASツール')
    .addItem('スライドに書き出し', 'insertKJItemsByColumnToSlides')
    .addToUi();
}

function insertKJItemsByColumnToSlides() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  const selection = spreadsheet.getActiveRangeList();
  if (!selection) {
    ui.alert('範囲を選択してから実行してください。');
    return;
  }

  const columnIndices = new Set();
  selection.getRanges().forEach(range => {
    const startCol = range.getColumn();
    for (let i = 0; i < range.getNumColumns(); i++) {
      columnIndices.add(startCol + i - 1);
    }
  });

  const sortedIndices = Array.from(columnIndices).sort((a, b) => a - b);
  if (sortedIndices.length === 0) {
    ui.alert('処理対象の列が選択されていません。列を選択してから実行してください。');
    return;
  }

  // API呼び出しなしで列名(A,B,C...)をローカル計算
  const colNames = sortedIndices.map(colIndexToLetter);

  if (ui.alert('実行確認', `選択された以下の列を処理しますか？\n${colNames.join(', ')}`, ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) {
    return;
  }

  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow === 0 || lastColumn === 0) {
    ui.alert('シートにデータがありません。');
    return;
  }

  // シート全体を1回のAPI呼び出しで取得
  const allData = sheet.getRange(1, 1, lastRow, lastColumn).getValues();

  const spreadsheetFile = DriveApp.getFileById(spreadsheet.getId());
  const parents = spreadsheetFile.getParents();
  const targetFolder = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();

  const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm');
  const newPresentation = SlidesApp.create(`いい会社とは_ ${targetFolder.getName()}_${dateStr}`);
  DriveApp.getFileById(newPresentation.getId()).moveTo(targetFolder);

  const SHAPE = SlidesApp.ShapeType.RECTANGLE;
  const COLS = 3, ROWS = 5;
  const MARGIN = 10, GAP = 10;
  const SLIDE_WIDTH = newPresentation.getPageWidth();
  const SLIDE_HEIGHT = newPresentation.getPageHeight();
  const BOX_WIDTH = (SLIDE_WIDTH - 2 * MARGIN - (COLS - 1) * GAP) / COLS;
  const BOX_HEIGHT = (SLIDE_HEIGHT - 2 * MARGIN - (ROWS - 1) * GAP) / ROWS;
  const PER_SLIDE = COLS * ROWS;
  let addedSlideCount = 0;

  sortedIndices.forEach(colIdx => {
    const values = allData.map(row => row[colIdx]).filter(v => v !== null && v !== '');
    if (values.length === 0) return;

    let page = newPresentation.appendSlide();
    addedSlideCount++;

    values.forEach((text, i) => {
      if (i > 0 && i % PER_SLIDE === 0) {
        page = newPresentation.appendSlide();
        addedSlideCount++;
      }
      const pos = i % PER_SLIDE;
      const x = MARGIN + (pos % COLS) * (BOX_WIDTH + GAP);
      const y = MARGIN + Math.floor(pos / COLS) * (BOX_HEIGHT + GAP);
      const textRange = page.insertShape(SHAPE, x, y, BOX_WIDTH, BOX_HEIGHT).getText();
      textRange.setText(String(text));
      textRange.getTextStyle().setFontSize(24);
      textRange.getParagraphStyle().setParagraphAlignment(SlidesApp.ParagraphAlignment.START);
    });
  });

  if (addedSlideCount === 0) {
    DriveApp.getFileById(newPresentation.getId()).setTrashed(true);
    ui.alert('書き出せるデータがありませんでした。');
    return;
  }

  ui.alert('完了', 'スライドの作成が完了しました。', ui.ButtonSet.OK);
}

function colIndexToLetter(idx) {
  let letter = '';
  let n = idx + 1;
  while (n > 0) {
    letter = String.fromCharCode(64 + (n % 26 || 26)) + letter;
    n = Math.floor((n - 1) / 26);
  }
  return letter;
}
