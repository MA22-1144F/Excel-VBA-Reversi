# Excel-VBA-Reversi

## Overview
VBAを利用してExcel Sheet上でリバーシがプレイできます。  
セルのクリックによる操作と座標の入力による操作に対応しています。  
PERSONAL.XLSBに保存して動作させることを想定しています。  

* Othello_2025：CPU対戦版(3段階の難易度設定)  
* Othello_P2_2025：2プレイヤー対戦版

## Requirement
Microsoft Excel上で動作します。  

## Usage
1. Othello_2025またはOthello_P2_2025を任意の場所に保存
2. Excelで新規WorkSheetを開く
3. 開発タブのVisual BasicまたはAlt+F11でVBE(Visual Basic Editor)を開く
4. VBAProject一覧からVBAProject (PERSONAL.XLSB)を選択し右クリック
5. ファイルのインポートで保存したOthello_2025またはOthello_P2_2025を開く
6. 上書き保存ボタンまたはCtrl+Sで保存
以下オプション設定  
7. 任意のExcel WorkSheetに戻り、ファイル→オプション→リボンのユーザー設定の順で遷移
8. 任意のユーザー設定グループにPERSONAL.XLSB!StartReversiGameまたはPERSONAL.XLSB!StartReversiGame2Pのマクロを追加
9. お好みで名前やアイコンを変更
10. Excel WorkSheetに戻り、設定したマクロのアイコンを選択してプレイを開始

## Author
[Instagram](https://www.instagram.com/nattotoasto?igsh=NWNtdHhnY3A4NDQ0 "nattotoasto")

## Licence
MIT
