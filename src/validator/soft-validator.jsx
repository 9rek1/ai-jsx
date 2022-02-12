/**
 * ** 概要 **
 * グレスケ印刷物用の.aiファイルのデータチェッカー
 *
 * ** 本スクリプトで対応できないもの **
 * - アピアランス
 * - スウォッチ・ブラシ・シンボルなど
 * - 複合シェイプ・ブレンド・エンベロープなど
 * - オブジェクトのカラーはInDesignのプリフライトで
 *
 * ** 自動で行う処理 **
 * - レイヤーとオブジェクトのロックを解除
 * - 空レイヤーを削除
 * - 必要に応じてラスタライズ効果設定を修正
 *
 * ** 確認事項 **
 * - ドキュメントのカラーモードがCMYKか
 * - アートボードが1つか
 *
 * ** 以下についてその有無を確認 **
 * - 非表示レイヤー
 * - 非表示オブジェクト
 * - 無色オブジェクト
 * - 孤立点
 * - 塗り付きの直線パス
 * - 塗り付きのオープンパス
 * - テキストフレーム
 * - グラフアイテム
 * - ラスター画像
 * - リンクオブジェクト
 * - アートボード外オブジェクト
 */

// @target 'illustrator'

/**
 * メインプロセス
 */
(function () {
  if (app.documents.length <= 0) {
    alert('ファイルを開いてください。');
    return;
  }
  var doc = app.activeDocument;
  var result = ''; // 最後にアラートする文章

  if (doc.documentColorSpace !== DocumentColorSpace.CMYK) {
    result += '【重要】ドキュメントのカラーモードがRGBです。\n';
  }

  // ドキュメントのラスタライズ効果設定
  modifyRasterSettings(doc);

  // レイヤーに関する処理
  var hiddenLays = handleLayers(doc);
  if (hiddenLays.length > 0) {
    result += hiddenLays.length + '個の非表示レイヤーがあります。\n';
  }

  // すべてのオブジェクトのロックを解除
  app.executeMenuCommand('unlockAll');
  unselectAll();

  // 非表示オブジェクトを取得
  var hiddenItems = getHiddenItems();
  if (hiddenItems.length > 0) {
    result += hiddenItems.length + '個の非表示オブジェクトがあります。\n';
  }

  // パスアイテムに関する処理
  result += handlePathItems(doc);

  // テキストフレームをカウント
  var texts = doc.textFrames;
  if (texts.length > 0) {
    result +=
      texts.length + '個のテキストがあります。アウトライン化してください。\n';
  }

  // グラフアイテムをカウント
  var graphs = doc.graphItems;
  if (graphs.length > 0) {
    result += graphs.length + '個のグラフがあります。拡張してください。\n';
  }

  // ラスターアイテムをカウント
  var rasterImages = doc.rasterItems;
  if (rasterImages.length > 0) {
    result += rasterImages.length + '個のラスター画像があります。\n';
  }

  // リンクアイテムをカウント
  var linkedImages = doc.placedItems;
  if (linkedImages.length > 0) {
    result += linkedImages.length + '個のリンクオブジェクトがあります。\n';
  }

  // アートボード外オブジェクトを取得
  if (doc.artboards.length > 1) {
    result += 'アートボードが' + doc.artboards.length + '個あります。\n';
  } else {
    var outBoardItems = getOutBoardItems();
    if (outBoardItems.length > 0) {
      result +=
        outBoardItems.length + '個のアートボード外オブジェクトがあります。\n';
    }
  }

  // 非表示オブジェクトを，非表示に戻す
  for (var i = 0, len = hiddenItems.length; i < len; i++) {
    hiddenItems[i].hidden = true;
  }

  // 非表示レイヤーを，非表示に戻す
  for (var i = 0, len = hiddenLays.length; i < len; i++) {
    hiddenLays[i].visible = false;
  }

  // 結果を表示
  if (result === '') result = '安全なデータです。すごい！';
  alert(result);
})();

/**
 * ドキュメントのラスタライズ効果設定を確認・修正する
 * @param {Document} docObj ドキュメント
 */
function modifyRasterSettings(docObj) {
  var settings = docObj.rasterEffectSettings;
  if (settings.colorModel !== RasterizationColorModel.GRAYSCALE) {
    var shouldModify = confirm(
      'ドキュメントのラスタライズ効果設定のカラーモードを\n' +
        'グレースケールに変更しますか？'
    );
    if (shouldModify) settings.colorModel = RasterizationColorModel.GRAYSCALE;
  }
  if (settings.resolution !== 350) {
    var shouldModify = confirm(
      'ドキュメントのラスタライズ効果設定の解像度を\n' +
        settings.resolution +
        'から350に変更しますか'
    );
    if (shouldModify) settings.resolution = 350;
  }
}

/**
 * すべてのレイヤーを表示＆ロック解除 + 空レイヤーを削除 +
 * 非表示レイヤーを返す
 * @param {Document} docObj ドキュメント
 * @returns {Array<Layer>} 非表示レイヤーのリスト
 */
function handleLayers(docObj) {
  var result = []; // 非表示レイヤーを入れるリスト
  (function rec(docObj) {
    var lays = docObj.layers;
    var len = lays.length;
    for (var i = len - 1; i >= 0; i--) {
      var currentLay = lays[i];
      // レイヤーのロックを解除
      currentLay.locked = false;
      // 空レイヤーを削除
      if (isEmptyLayer(currentLay)) {
        currentLay.visible = true;
        currentLay.remove();
        continue;
      }
      // 非表示レイヤーはリストに格納
      if (!currentLay.visible) {
        currentLay.visible = true; // レイヤーを表示
        result.push(currentLay);
      }
      rec(currentLay); // サブレイヤーのために再帰呼び出し
    }
  })(docObj);
  return result;
}

/**
 * レイヤーが完全に空かどうかを返す
 * @param {Layer} layObj レイヤー
 * @returns {boolean}
 */
function isEmptyLayer(layObj) {
  return layObj.pageItems.length === 0 && layObj.layers.length === 0;
}

/**
 * すべての選択を解除する
 */
function unselectAll() {
  app.selection = null;
}

/**
 * すべてのオブジェクトを表示 + 非表示オブジェクトを返す
 * @returns {Array<PageItem>}
 */
function getHiddenItems() {
  app.executeMenuCommand('showAll'); // すべてのオブジェクトを表示
  var result = app.selection;
  unselectAll();
  return result;
}

/**
 * ドキュメント内のパスについて次のものがないか確認する。
 * 無色オブジェクト + 孤立点 + 塗り付きの直線パス・オープンパス
 * @param {Document} docObj ドキュメント
 * @returns {string} 確認の結果
 */
function handlePathItems(docObj) {
  var paths = docObj.pathItems;
  var len = paths.length;
  var colorlessItems = [];
  var strayPoints = [];
  var filledLines = [];
  var filledOpenPaths = [];
  for (var i = len - 1; i >= 0; i--) {
    var currentPath = paths[i];
    // 無色オブジェクトについて
    if (isColorless(currentPath)) {
      colorlessItems.push(currentPath);
    }
    // 孤立点について
    if (currentPath.pathPoints.length <= 1) {
      strayPoints.push(currentPath);
      continue;
    }
    // ガイド以外の塗りがついているオープンパスについて
    if (!currentPath.guides && currentPath.filled && !currentPath.closed) {
      if (isLinear(currentPath)) {
        filledLines.push(currentPath);
      } else {
        filledOpenPaths.push(currentPath);
      }
    }
  }

  var result = '';
  if (colorlessItems.length > 0) {
    result += colorlessItems.length + '個の無色オブジェクトがあります。\n';
  }
  if (strayPoints.length > 0) {
    result += strayPoints.length + '個の孤立点があります。\n';
  }
  if (filledLines.length > 0) {
    result += filledLines.length + '個の塗り付きの直線パスがあります。\n';
  }
  if (filledOpenPaths.length > 0) {
    result +=
      filledOpenPaths.length + '個の塗り付きのオープンパスがあります。\n';
  }
  return result;
}

/**
 * 無色オブジェクトかどうかを返す
 * @param {PathItem} pathObj パス
 * @returns {boolean}
 */
function isColorless(pathObj) {
  return !pathObj.filled && !pathObj.stroked && !pathObj.clipping;
}

/**
 * 一直線のパスかどうかを返す
 * @param {PathItem} pathObj パス
 * @returns {boolean}
 */
function isLinear(pathObj) {
  // パスの面積がほぼ0かどうかで判定
  return Math.round(Math.abs(pathObj.area) * 100000) === 0;
}

/**
 * アートボード外オブジェクトを取得する
 * @returns {Array<PageItem>}
 */
function getOutBoardItems() {
  app.executeMenuCommand('selectallinartboard');
  app.redraw(); // 選択範囲反転のために必要
  app.executeMenuCommand('Inverse menu item');
  var result = app.selection;
  unselectAll();
  return result;
}
