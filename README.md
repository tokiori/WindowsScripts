# WindowsScripts
WSFで使用頻度が高い（環境が変わってもよく再利用する）スクリプト、ツール

## 送るメニューに登録する機能

- [wsf]コピーしてリネーム（タイムスタンプ付与）.wsf
  - 対象パスを同階層にコピーしてタイムスタンプ付与したリネームを行う。

- [wsf]ハッシュ生成（MD5,SHA1,SHA256）.wsf
  - 対象ファイルのハッシュを生成して同階層に「元ファイル名.hash.txt」として保存する。

- [wsf]ブックの全シートでA1セル選択.wsf
  - 対象ファイルの全シートのフォーカスをA1にして保存する。
  - [YES]ボタン押下で上書き
  - [NO]ボタン押下で同階層に「元ファイル名_a1focus_YYYYMMDDhhmmss」を付与して保存


## 使い方

1. このプロジェクトの全ファイルをダウンロードして、任意のフォルダに保管する

2. 「wsf.js.tools.register.wsf」を実行して送るメニューに登録する

※2. のフォルダに対してのショートカットを送るメニューに登録するため、登録後に2.のフォルダを移動した際は再度実行して登録しなおすこと


## フォルダ構成

- wsf.js
  - WSFで使用頻度が高い機能（jscript）をライブラリ化したもの。

- wsf.js.test.wsf
  - 「wsf.js」の簡易的な機能確認スクリプト（サンプル）。
  - 実行すると「wsf.js.test_[YYYYMMDDhhmmss].wsf」の実行ログが作成される。

- tools.sendtoフォルダ
  - 「wsf.js」を利用したwindowsの送るメニューに登録するスクリプト。

- htaフォルダ
  - 「wsf.js」を利用したHTA起動実行WSFスクリプト（サンプル）。

- wsf.js.tools.register.wsf
  - 「tools.sendto」フォルダ内のスクリプトを登録するスクリプト。
  - 実行すると「tools.sendto」フォルダ内のスクリプトをカレントユーザの送るメニューに登録する。


## jscriptライブラリ（wsf.js）

- js直下
  - よく使う機能を詰め込んだもの。（jQueryライクなextendとeachや、型判定、オブジェクトのソート、最大最少長取得等。）

- js.str
  - 文字列操作（主にリプレース関連）機能。最近のJavascriptだとバッククオート「`」とかで実現できるやつとか。

- js.date
  - 主にDate型にFormat機能（YYYYMMDDhhmmssmsをリプレース）を付加したもの。

- js.path
  - FileSystemObject関連をWrapしたもの。

- js.log
  - ロガー機能を集約したもの（「js.log.on()」でデフォルトログファイルに出力）。

- js.book
  - ExcelAppをWrapしたいもの。

- js.cmd
  - Windowsコマンド関連を集約したもの。
  - 「js.cmd.hta()」でHTAを実行する。（引数渡し可能）

## License 
* Copyright (c) 2022 tokiori
* This software is Released under the MIT license.
* see https://opensource.org/licenses/MIT
