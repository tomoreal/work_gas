/**
 * スプレッドシートが開かれた時に実行される関数
 * カスタムメニューを追加します
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('GASツール')
    .addItem('スライドに書き出し', 'insertKJItemsByColumnToSlides')
    .addToUi();
}

function insertKJItemsByColumnToSlides() {
  // 1. アクティブなスプレッドシートとシートを取得
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = spreadsheet.getActiveSheet();

  // ▼▼▼ 1. 処理したい列の範囲をアクティブな選択範囲から取得 ▼▼▼
  const selection = spreadsheet.getActiveRangeList();
  if (!selection) {
    SpreadsheetApp.getUi().alert('範囲が取得できませんでした。列を選択してから実行してください。');
    return;
  }

  const columns = [];
  selection.getRanges().forEach(range => {
    const startCol = range.getColumn();
    const numCols = range.getNumColumns();
    for (let i = 0; i < numCols; i++) {
      // 列全体のA1表記を取得 (例: 'A:A')
      const colLetter = sheet.getRange(1, startCol + i).getA1Notation().replace(/[0-9]/g, '');
      const colRange = `${colLetter}:${colLetter}`;
      if (!columns.includes(colRange)) {
        columns.push(colRange);
      }
    }
  });

  if (columns.length === 0) {
    SpreadsheetApp.getUi().alert('処理対象の列が選択されていません。列を選択してから実行してください。');
    return;
  }

  // 実行確認ダイアログ
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '実行確認',
    `選択された以下の列を処理しますか？\n${columns.join(', ')}`,
    ui.ButtonSet.OK_CANCEL
  );

  if (response !== ui.Button.OK) {
    return;
  }

  // 2. スライドの新規作成とフォルダ処理

  // 2-1. スプレッドシートのファイルオブジェクトを取得
  const spreadsheetFile = DriveApp.getFileById(spreadsheet.getId());

  // 2-2. 親フォルダを取得（最初の親フォルダを対象とする）
  const parents = spreadsheetFile.getParents();
  let targetFolder = DriveApp.getRootFolder(); // 親フォルダがない場合のデフォルト
  if (parents.hasNext()) {
    targetFolder = parents.next();
  }

  // 2-3. 新しいスライドのタイトルを決定
  const folderName = targetFolder.getName();
  const newSlideTitle = `いい会社とは_ ${folderName}`;

  // 2-4. 新しいプレゼンテーションを作成（デフォルトでマイドライブのルートに作成される）
  const newPresentation = SlidesApp.create(newSlideTitle);
  const newSlideId = newPresentation.getId();

  // 2-5. 作成したスライドファイルを目的のフォルダに移動
  // 変更後（推奨される moveTo メソッドを使用）
  const newSlideFile = DriveApp.getFileById(newSlideId);


  // 作成したスライドファイルを目的のフォルダへ移動する
  // これにより、ファイルは自動的にルートフォルダから削除され、
  // targetFolderの配下に入ります。
  newSlideFile.moveTo(targetFolder);


  // 3. 処理対象のプレゼンテーションを新規作成したものに設定
  const presentation = newPresentation;


  // 4. スライドへのデータ書き込み（以下の処理は変更なし）

  // ▼▼▼ 2. 配列の各列に対してループ処理を実行 ▼▼▼
  columns.forEach(columnRange => {
    // 現在の列から値を取得し、空のセルを除外
    const values = sheet.getRange(columnRange).getValues().flat().filter(v => v !== '');


    // ▼▼▼ 3. 列にデータがなければ、その列の処理をスキップ ▼▼▼
    if (values.length === 0) {
      return; // 次の列の処理へ移る
    }


    // ▼▼▼ 4. 新しいスライドを末尾に追加 ▼▼▼
    const page = presentation.appendSlide();


    // シェイプの初期位置やサイズを定義（スライドごとにリセットされる）
    const startX = 50;
    let startY = 50;
    const boxWidth = 200;
    const boxHeight = 60;
    const gap = 20;


    // ▼▼▼ 5. 取得した値を使ってシェイプをスライドに追加 ▼▼▼
    values.forEach(text => {
      const shape = page.insertShape(SlidesApp.ShapeType.RECTANGLE, startX, startY, boxWidth, boxHeight);
      shape.getText().setText(text);
      shape.getText().getTextStyle().setFontSize(14);
      startY += boxHeight + gap; // 次のシェイプのためにY座標をずらす
    });
  });
}
