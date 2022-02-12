/**
 * ** 概要 **
 * 見た目を維持しながら塗り付きのオープンパスを自動処理するスクリプトです
 *
 * ** 注意 **
 * 1. アピアランスについて
 * 本スクリプトではパスを閉じる際にアンカーのハンドルを操作します
 * ハンドルに左右されるアピアランス（「パスの変形」項目など）がついている場合，見た目が変わってしまいます
 * 不安な場合は事前にアピアランス分割することをおすすめします
 *
 * 2. シンボルアイテムやプラグインアイテムについて
 * シンボルや複合シェイプやブレンドといったアイテム内にある塗り付きのオープンパスは本スクリプトでは検出できません
 * それらをチェックしたい場合は，各アイテムを拡張するなどしてから本スクリプトを実行してください
 *
 * 3. 複合パス内にある線付き塗りオープンパスについて
 * 複合パス内では，動作12（線と塗りを別のパスに分割）ができません
 * 複合パス内の線付き塗りオープンパスは，スクリプト実行後に個別で対応してください
 * 該当パスがあった場合は，スクリプト処理中に該当パスを選択してアラートします
 *
 * 4. その他
 * - レイヤーとオブジェクトのロックは自動で解除します
 * - ガイドには手をつけないので，オープンパスの数を確認する際は注意してください
 * - 本スクリプトは「見た目を変えずに」塗り付きのオープンパスを処理するものです
 *   線を足しながらオープンパスを閉じたい場合は「ナイフツールで囲む」といった方法が有用です
 * - 本スクリプトは自己責任においてご利用ください
 * - 本スクリプトに不具合や改善すべき点がございましたら，作成者まで連絡いただけると幸いです
 *
 * ** 動作 **
 * -- レイヤーに関する処理 --
 * 1. 全レイヤーのロックを解除します
 * 2. 空レイヤーを削除します
 * 3. 非表示レイヤーがあれば，ダイアログを出して削除するか確認します
 * 4. ダイアログで許可があれば非表示レイヤーを削除します
 *
 * -- オブジェクトに関する前処理 --
 * 5. 全オブジェクトのロックを解除します
 * 6. 非表示アイテムがあれば，ダイアログを出して削除するか確認します
 * 7. ダイアログで許可があれば非表示オブジェクトを削除します
 * 8. 孤立点があれば，ダイアログを出して削除するか確認します
 * 9. ダイアログで許可があれば孤立点を削除します
 *
 * -- 塗り付きのオープンパスの処理 --
 * 10. 塗りオープンパスのうち，直線パス（塗りの面積はないが塗り色がついている）を処理します（**詳細**の1）
 * 11. 塗りオープンパスのうち，線がついていないものはそのままパスを閉じます（**詳細**の2）
 * 12. 線がついている塗りオープンパスは，線と塗りを別々のパスにして，塗りのパスのみを閉じます（**詳細**の3）
 *
 * -- 結果のアラート --
 * 13. 処理しなかった孤立点と塗りオープンパスの個数をアラートして終了します
 *
 * ** 詳細 **
 * 1. 動作10（塗りがついている直線の処理）
 * 塗りはあるが線はないヘアラインパスと，塗りも線もついているパスで処理が分かれます
 * 前者の場合はパスを削除します。後者の場合は塗りを消去します
 * 処理の前にはダイアログを出して確認します。不許可の場合は何もしません
 * ※直線パスでもクローズパスの場合は本スクリプトの処理の対象外になります
 *
 * 2. 動作11（線がついていない塗りオープンパスの処理）
 * 塗り部分の見た目を変えないように始点アンカーと終点アンカーをつなげて，パスを閉じます
 * 処理としては「パスの連結」(Ctrl(⌘)+J)に近いです（なお「パスの連結」では見た目が変わる場合があります）
 *
 * 3. 動作12（線がついている塗りオープンパスの処理）
 * 前述の方法では，パスに線がついている場合はアンカー同士をつなげた分の線が足されてしまいます
 * そうならないように，線と塗りを別のパスに分割して，塗りのパスにだけ前述の処理をしてパスを閉じます
 * 線と塗りの分割では，パスを複製し，片方は塗りを消去，もう片方は線を消去しています
 */

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

  // レイヤーに関する処理
  var hiddenLays = handleLayers(doc); // 返り値は非表示レイヤーのリスト

  // すべてのオブジェクトのロックを解除
  app.executeMenuCommand('unlockAll');
  unselectAll();

  // 確認ダイアログで許可があれば，非表示オブジェクトを削除
  var hiddenItems = removeHiddenItems(); // 返り値は非表示オブジェクトのリスト

  // 塗り付きのオープンパスに関する処理
  var paths = doc.pathItems;
  var lenPath = paths.length;
  var spCount = 0; // 処理しなかった孤立点の個数
  var fopCount = 0; // 処理しなかった塗りオープンパスの個数
  var cfmStrayPoint = new Confirmer(
    '孤立点があります',
    '削除する',
    '削除しない'
  );
  var cfmFilledLine = new Confirmer('直線パスに塗りがついています');
  for (var i = lenPath - 1; i >= 0; i--) {
    var currentPath = paths[i];
    // 孤立点について
    if (currentPath.pathPoints.length <= 1) {
      var removed = removeNeedlessItem(currentPath, '孤立点', cfmStrayPoint);
      if (!removed) spCount++;
      continue;
    }
    // ガイド以外の塗りがついているオープンパスについて
    if (!currentPath.guides && currentPath.filled && !currentPath.closed) {
      if (isLinear(currentPath)) {
        var fixed = fixFilledLine(currentPath, cfmFilledLine);
        if (!fixed) fopCount++;
      } else {
        var fixed = fixFilledOpenPath(currentPath);
        if (!fixed) fopCount++;
      }
    }
  }
  if (spCount > 0) {
    result += spCount + '個の孤立点が未処理です。\n';
  }
  if (fopCount > 0) {
    result += fopCount + '個の塗り付きのオープンパスが未処理です。';
  }
  unselectAll();

  // 削除しなかった非表示オブジェクトを，非表示に戻す
  for (var i = 0, len = hiddenItems.length; i < len; i++) {
    hiddenItems[i].hidden = true;
  }

  // 削除しなかった非表示レイヤーを，非表示に戻す
  for (var i = 0, len = hiddenLays.length; i < len; i++) {
    hiddenLays[i].visible = false;
  }

  // 結果を表示
  if (result === '') result = '塗り付きのオープンパスは0個になりました。';
  alert(result);
})();

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
 * @returns {boolean}
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
