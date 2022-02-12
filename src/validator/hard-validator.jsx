// @target 'illustrator'

/**
 * 「再表示しない」ボックスがついた確認ダイアログのクラス
 */
var Confirmer = (function () {
  /**
   * コンストラクタ
   * @param {string} [title="重要な確認事項があります"] （省略可）ダイアログのタイトル
   * @param {string} [yesBtnTxt="はい"] （省略可）承諾ボタンのテキスト
   * @param {string} [noBtnTxt="いいえ"] （省略可）拒否ボタンのテキスト
   */
  var Confirmer = function (title, yesBtnTxt, noBtnTxt) {
    // newをつけ忘れた場合に備えて
    if (!(this instanceof Confirmer)) {
      return new Confirmer();
    }
    this.shouldConfirm = true; // 確認ダイアログを表示するかどうか
    this.result = false; // 確認の結果

    // デフォルト引数
    if (title === undefined) title = '重要な確認事項があります';
    if (yesBtnTxt === undefined) yesBtnTxt = 'はい';
    if (noBtnTxt === undefined) noBtnTxt = 'いいえ';

    this.title = title;
    this.yesBtnTxt = yesBtnTxt;
    this.noBtnTxt = noBtnTxt;
  };

  /**
   * 確認の結果を返す。必要に応じて確認ダイアログを表示する
   * @param {string} body ダイアログの本文
   * @returns {boolean} 確認の結果
   */
  Confirmer.prototype.confirmIfNeed = function (body) {
    if (this.shouldConfirm) {
      var dlg = new Window('dialog', this.title);
      dlg.add('statictext', undefined, body, { multiline: true });

      var btnGrp = dlg.add('group');
      var yesBtn = btnGrp.add('button', undefined, this.yesBtnTxt);
      yesBtn.onClick = function () {
        dlg.close(1); // dlg.show()の返り値が1になる
      };
      var noBtn = btnGrp.add('button', undefined, this.noBtnTxt);
      noBtn.onClick = function () {
        dlg.close(0); // dlg.show()の返り値が0になる
      };

      var checkbox = dlg.add(
        'checkbox',
        undefined,
        'このダイアログを再表示しない'
      );
      checkbox.value = false; // チェックボックスのデフォルトの値

      var result = dlg.show();
      if (result === 2) return false; // ダイアログが閉じられた場合の処理
      this.result = result;
      this.shouldConfirm = !checkbox.value;
    }
    return this.result;
  };

  /**
   * どのオブジェクトについてConfirmer.confirmIfNeedしているかわかるように
   * 事前に対象オブジェクトを選択する
   * @param {PageItem} item
   */
  Confirmer.prototype.selectItemIfNeed = function (item) {
    if (this.shouldConfirm) {
      app.selection = item;
      app.redraw(); // 選択を描画
    }
  };

  return Confirmer;
})();

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
  var hiddenLays = handleLayers(doc); // 返り値は非表示レイヤーのリスト

  // すべてのオブジェクトのロックを解除
  app.executeMenuCommand('unlockAll');
  unselectAll();

  // 確認ダイアログで許可があれば，非表示オブジェクトを削除
  var hiddenItems = removeHiddenItems(); // 返り値は非表示オブジェクトのリスト

  // パスアイテムに関する処理
  handlePathItems(doc);

  // テキストのアウトラインを作成
  outlineTexts(doc);

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

  // アートボード外オブジェクトを削除
  if (doc.artboards.length > 1) {
    result += 'アートボードが' + doc.artboards.length + '個あります。\n';
  } else {
    removeOutBoardItems();
  }

  // 削除しなかった非表示オブジェクトを，非表示に戻す
  for (var i = 0, len = hiddenItems.length; i < len; i++) {
    hiddenItems[i].hidden = true;
  }

  // 削除しなかった非表示レイヤーを，非表示に戻す
  for (var i = 0, len = hiddenLays.length; i < len; i++) {
    hiddenLays[i].visible = false;
  }

  // 結果を表示
  result += 'データチェックが完了しました。';
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
 * 許可があれば非表示レイヤーを削除 + 削除しなかった非表示レイヤーを返す
 * @param {Document} docObj ドキュメント
 * @returns {Array<Layer>} 削除しなかった非表示レイヤーのリスト
 */
function handleLayers(docObj) {
  var result = []; // 非表示レイヤーを入れるリスト
  var confirmer = new Confirmer(
    '非表示のレイヤーがあります',
    '削除する',
    '削除しない'
  );

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
      // 非表示レイヤーに関する処理
      if (!currentLay.visible) {
        currentLay.visible = true; // レイヤーを表示
        var shouldRemove = confirmer.confirmIfNeed(
          'レイヤー「' +
            currentLay.name +
            '」は非表示レイヤーです。オブジェクトごと削除しますか？'
        );
        if (shouldRemove) {
          currentLay.remove();
          continue;
        } else {
          result.push(currentLay);
        }
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
 * すべてのオブジェクトを表示 + 非表示オブジェクトを確認の上で削除 +
 * 削除しなかった非表示オブジェクトを返す
 * @returns {Array<PageItem>} 削除しなかった非表示オブジェクトのリスト
 */
function removeHiddenItems() {
  var result = []; // 非表示オブジェクトを入れるリスト
  var confirmer = new Confirmer(
    '非表示のオブジェクトがあります',
    '削除する',
    '削除しない'
  );

  app.executeMenuCommand('showAll'); // すべてのオブジェクトを表示
  var hiddenItems = app.selection;
  var len = hiddenItems.length;
  for (var i = len - 1; i >= 0; i--) {
    var currentItem = hiddenItems[i];
    var removed = removeNeedlessItem(
      currentItem,
      '非表示のオブジェクト',
      confirmer
    );
    if (!removed) result.push(currentItem);
  }
  unselectAll();
  return result;
}

/**
 * オブジェクトを選択して確認ダイアログを出す。そこで許可があれば削除する。
 * 削除したかどうかを返す
 * @param {PageItem} item オブジェクト
 * @param {string} kind 不要アイテムの種類。ダイアログの本文に用いる
 * @param {Confirmer} confirmer
 * @returns {boolean} 削除したかどうか
 */
function removeNeedlessItem(item, kind, confirmer) {
  confirmer.selectItemIfNeed(item);
  var shouldRemove = confirmer.confirmIfNeed(
    '現在選択されているオブジェクトは' + kind + 'です。削除しますか？'
  );
  if (shouldRemove) {
    item.remove();
    return true;
  } else {
    return false;
  }
}

/**
 * ドキュメント内のパスについて次の処理をする。
 * 確認の上で無色オブジェクトを削除 + 確認の上で孤立点を削除 +
 * 塗り付きのオープンパスを適切に処理
 * @param {Document} docObj ドキュメント
 */
function handlePathItems(docObj) {
  var paths = docObj.pathItems;
  var len = paths.length;
  var cfmColorless = new Confirmer(
    '無色オブジェクトがあります',
    '削除する',
    '削除しない'
  );
  var cfmStrayPoint = new Confirmer(
    '孤立点があります',
    '削除する',
    '削除しない'
  );
  var cfmFilledLine = new Confirmer('直線パスに塗りがついています');

  for (var i = len - 1; i >= 0; i--) {
    var currentPath = paths[i];
    // 無色オブジェクトについて
    if (isColorless(currentPath)) {
      var removed = removeNeedlessItem(
        currentPath,
        '無色オブジェクト',
        cfmColorless
      );
      if (removed) continue;
    }
    // 孤立点について
    if (currentPath.pathPoints.length <= 1) {
      removeNeedlessItem(currentPath, '孤立点', cfmStrayPoint);
      continue;
    }
    // ガイド以外の塗りがついているオープンパスについて
    if (!currentPath.guides && currentPath.filled && !currentPath.closed) {
      if (isLinear(currentPath)) {
        fixFilledLine(currentPath, cfmFilledLine);
      } else {
        fixFilledOpenPath(currentPath);
      }
    }
  }
  unselectAll();
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
 * 塗り付きの直線パスを処理する。処理したかどうかを返す
 * @param {PathItem} pathObj 塗り付きの直線パス
 * @param {Confirmer} confirmer
 * @returns {boolean}
 */
function fixFilledLine(pathObj, confirmer) {
  if (pathObj.stroked) {
    confirmer.selectItemIfNeed(pathObj);
    var shouldErase = confirmer.confirmIfNeed(
      '現在選択されている直線パスには，塗りがついています。塗りを消去しますか？'
    );
    if (shouldErase) pathObj.filled = false;
    return !shouldErase;
  } else {
    // 線なし塗り付きのヘアラインパスを削除する
    return removeNeedlessItem(pathObj, 'ヘアラインパス', confirmer);
  }
}

/**
 * 塗り付きのオープンパスを処理する。処理したかどうかを返す
 * @param {PathItem} pathObj 塗りオープンパス
 * @returns {boolean}
 */
function fixFilledOpenPath(pathObj) {
  if (!pathObj.stroked) {
    // 線がない場合は塗りを維持しながら閉じる
    closePath(pathObj);
    return true;
  }
  if (isCompounded(pathObj)) {
    app.selection = pathObj;
    app.redraw(); // 選択を描画
    alert(
      '現在選択されているパスは複合パス内の塗り付きオープンパスです。\n' +
        'このスクリプトでは自動対処しないので，後で個別対応をお願いします。'
    );
    return false;
  }
  // パスを複製し，片方は線を消去して閉じる，もう片方は塗りを消去して線のみにする
  var copied = pathObj.duplicate();
  pathObj.stroked = false;
  closePath(pathObj);
  copied.filled = false;
  return true;
}

/**
 * 塗りの形を維持しながらパスを閉じる
 * @param {PathItem} pathObj オープンパス
 */
function closePath(pathObj) {
  var anchors = pathObj.pathPoints;
  var firstAnchor = anchors[0];
  var lastAnchor = anchors[anchors.length - 1];

  pathObj.closed = true; // ハンドルはそのままでパスを閉じる
  if (
    // 始点アンカーと終点アンカーが重なっているとき
    firstAnchor.anchor[0] === lastAnchor.anchor[0] &&
    firstAnchor.anchor[1] === lastAnchor.anchor[1]
  ) {
    // 終点アンカーを削除しても見た目が変わらないように始点アンカーの内向きハンドルを操作
    firstAnchor.leftDirection = lastAnchor.leftDirection;
    lastAnchor.remove();
  } else {
    // 始点アンカーの内向きハンドルを削除
    firstAnchor.leftDirection = firstAnchor.anchor;
    // 終点アンカーの外向きハンドルを削除
    lastAnchor.rightDirection = lastAnchor.anchor;
  }
}

/**
 * 複合パスの一部かどうかを返す
 * @param {PathItem} pathObj パス
 * @returns {boolean}
 */
function isCompounded(pathObj) {
  var parentObj = pathObj.parent;
  if (parentObj.typename === 'Layer') {
    return false;
  } else if (parentObj.typename === 'CompoundPathItem') {
    return true;
  } else {
    // パスのparentがレイヤーでも複合パスでもない，つまりグループの場合は再帰呼び出し
    return isCompounded(parentObj);
  }
}

/**
 * ドキュメント内のテキストのアウトラインを作成する
 * @param {Document} docObj ドキュメント
 */
function outlineTexts(docObj) {
  var texts = docObj.textFrames;
  var len = texts.length;
  var confirmer = new Confirmer('テキストフレームがあります');

  for (var i = len - 1; i >= 0; i--) {
    var currentText = texts[i];
    confirmer.selectItemIfNeed(currentText);
    var shouldOutline = confirmer.confirmIfNeed(
      '現在選択されているアイテムはテキストフレームです。アウトライン化しますか？'
    );
    if (shouldOutline) currentText.createOutline();
  }
}

/**
 * アートボード外オブジェクトを削除する
 */
function removeOutBoardItems() {
  var confirmer = new Confirmer(
    'アートボード外オブジェクトがあります',
    '削除する',
    '削除しない'
  );

  app.executeMenuCommand('selectallinartboard');
  app.redraw(); // 選択範囲反転のために必要
  app.executeMenuCommand('Inverse menu item');
  var outBoardItems = app.selection;
  var len = outBoardItems.length;
  for (var i = len - 1; i >= 0; i--) {
    removeNeedlessItem(
      outBoardItems[i],
      'アートボード外のオブジェクト',
      confirmer
    );
  }
  unselectAll();
}
