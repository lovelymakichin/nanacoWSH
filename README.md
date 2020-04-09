# nanacoギフト登録
 
えらべる倶楽部等の福利厚生サービスでnanacoギフトを購入した場合、
1000円単位にギフトコードが分かれてしまい、複数回の処理が必要となるため
プログラムした。例）10万円=100回
 
## 簡単な説明
 
nanacoCode.txtを作成し、ギフトコードを張り付ける。1行1コード
'cscript nanacologconf.vbs'を実行

 
## 機能
 
- ログ出力
 
## 必要要件
 
- Windows
- vbs実行を許可
- カレントディレクトリにlogフォルダを作成
- 文字コードANSI
 
## 使い方
 
1. nanacologconf.vbsをカレントディレクトリにコピー
2. nanacoCode.txtにギフトコードを張り付け、カレントディレクトリにコピー
3. 'cscript nanacologconf.vbs'を実行
