/**
 * ** 概要 **
 * テキストを複製して，片方をアウトライン化，もう片方をテキスト用のレイヤーに移動するスクリプトです
 * スクリプト実行時に選択済みのアイテムがあれば，その内のテキストに対してのみ上記の処理をします
 *
 * ** 注意 **
 * - テキスト用レイヤーかどうかは，レイヤー名と最下位レイヤーかどうかで判断します
 * - デフォルトのテキスト用レイヤー名は「Texts」です
 *   変更したい場合は43行目を書き換えてください
 * - レガシーテキストには対応していません
 * - 本スクリプトは自己責任においてご利用ください
 * - 本スクリプトに不具合や改善すべき点がございましたら，作成元まで連絡いただけると幸いです
 *
 * ** 動作 **
 * 1. 最下位のテキスト用レイヤーがない場合は作成します
 *
 * -- スクリプト実行時に選択済みのアイテムがある場合 --
 * 2. 選択範囲の中からテキストのみを抽出します
 * ※ テキスト用レイヤー内のテキストと空テキストは抽出しません
 *
 * -- スクリプト実行時に選択済みのアイテムがない場合 --
 * 2. 表示されていて，かつロックされていないテキストをすべて抽出します
 * ※ テキスト用レイヤー内のテキストと空テキストは抽出しません
 *
 * 3. 2で抽出したテキストをそれぞれ複製します
 * 4. 複製した方のテキストをテキスト用レイヤーに移動します
 * 5. 複製元のテキストのアウトラインを作成します
 * 6. 結果をアラートして終了します
 */

// @target 'illustrator'

(function () {
  if (app.documents.length <= 0) {
    alert('ファイルを開いてください。');
    return;
  }
  var doc = app.activeDocument;
  var layList = doc.layers;
  var textList = doc.textFrames;

  var result = ''; // 最後にアラートする文章
  var layName = 'Texts'; // レイヤー名を変える場合はここを変更

  // 最下位のテキスト用レイヤーがない場合は作成
  var bottomLay = layList[layList.length - 1];
  if (bottomLay.name === layName) {
    var textLay = bottomLay;
  } else {
    var shouldAdd = confirm(
      '最下位レイヤーが「' +
        layName +
        '」ではありません。\n' +
        '最下位レイヤー「' +
        layName +
        '」を作成しますか？'
    );
    if (shouldAdd) {
      var textLay = layList.add();
      textLay.name = layName;
      textLay.move(doc, ElementPlacement.PLACEATEND);
      result += '最下位レイヤー「' + layName + '」を作成しました。\n';
    } else {
      alert('処理を中断しました。');
      return;
    }
  }

  // テキスト用レイヤーの状態を記録
  var isVisible = textLay.visible;
  var isLocked = textLay.locked;

  // 処理の対象となるテキストを収集
  var targetTexts = collectTargetTexts(textList, textLay);
  var lenTarget = targetTexts.length;

  if (lenTarget > 0) {
    // テキスト用レイヤーを一時的に表示＆ロック解除
    textLay.visible = true;
    textLay.locked = false;

    // テキストを複製 + テキスト用レイヤーに移動 + 複製元テキストのアウトラインを作成
    for (var i = 0; i < lenTarget; i++) {
      var copied = targetTexts[i].duplicate();
      copied.move(textLay, ElementPlacement.PLACEATBEGINNING);
      targetTexts[i].createOutline();
    }
    result += lenTarget + '個のテキストのアウトラインを作成しました。';
  }

  // テキスト用レイヤーの状態を戻す
  textLay.visible = isVisible;
  textLay.locked = isLocked;

  // 結果をアラート
  if (result === '') result += '処理対象のテキストがありませんでした。';
  alert(result);
})();

/**
 * 処理の対象となるテキストを収集する
 * @param {Array<TextFrameItem>} textList activeDocument.textFrames
 * @param {Layer} textLay テキスト用レイヤー
 * @returns {Array<TextFrameItem>}
 */
function collectTargetTexts(textList, textLay) {
  var result = [];
  var len = textList.length;

  // テキスト用レイヤー内から抽出されないようにレイヤーをロック
  // app.selectionからも外れる
  textLay.locked = true;

  // 選択済みのアイテムがあれば，その中から空でないテキストを抽出
  if (app.selection.length > 0) {
    for (var i = 0; i < len; i++) {
      var cur = textList[i];
      if (cur.selected && isNotEmptyFrame(cur)) result.push(cur);
    }
  } else {
    // 編集可能で（自身と親が表示かつアンロック）空でないテキストを抽出
    for (var i = 0; i < len; i++) {
      var cur = textList[i];
      if (cur.editable && isNotEmptyFrame(cur)) result.push(cur);
    }
  }
  return result;
}

/**
 * テキストフレームが空かどうかを返す
 * @param {TextFrameItem} text テキストフレーム
 * @returns {boolean}
 */
function isNotEmptyFrame(text) {
  return text.contents.length > 0;
}
