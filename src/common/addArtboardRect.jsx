/**
 * ** 概要 **
 * アートボードサイズの四角形を作成します
 * http://gorolib.blog.jp/archives/53547822.html より
 *
 * ** 使い方 **
 * - アートボードサイズでクリッピングしたい時に便利です
 * - 実行してできた四角形のサイズを変えた後，「オブジェクト」→「アートボード」→「選択オブジェクトに合わせる」も便利です
 * - ちなみにアートボードに合わせて整列は整列パネルからできます
 *
 * ** 動作 **
 * 1. アクティブなアートボードの4頂点の座標を取得します
 * 2. そのサイズの四角形をアクティブなレイヤーに追加します
 * 3. 四角形の塗りと線を消し，選択したら終了です
 */

// @target 'illustrator'

(function () {
  if (app.documents.length <= 0) {
    alert('ファイルを開いてください。');
    return;
  }
  var doc = app.activeDocument;
  var boardIndex = doc.artboards.getActiveArtboardIndex();
  var boardRect = doc.artboards[boardIndex].artboardRect;
  var x0 = boardRect[0];
  var y0 = -boardRect[1];
  var x1 = boardRect[2];
  var y1 = -boardRect[3];

  var activeLay = doc.activeLayer;
  if (activeLay.locked) {
    alert('アクティブレイヤーがロックされています。');
    return;
  }
  if (!activeLay.visible) {
    alert('アクティブレイヤーは非表示です。');
    return;
  }

  var targetRect = activeLay.pathItems.rectangle(-y0, x0, x1 - x0, y1 - y0);
  targetRect.filled = false;
  targetRect.stroked = false;
  targetRect.selected = true;
})();
