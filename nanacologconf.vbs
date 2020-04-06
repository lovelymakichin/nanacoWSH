'Dim anchor As HTMLAnchorElement
' ファイル入出力の定数
Const conForReading = 1, conForWriting = 2, conForAppending = 8

' iniファイル名
Const FolderName = "D:\Program\nanacoWSH\"
Const conIniFileName = "nanaco.ini"
Const inputFileName = "nanacoCode.txt"

' 格納先Dictionaryオブジェクトの作成
Set objSectionDic = CreateObject("Scripting.Dictionary")



Set fso = WScript.CreateObject("Scripting.FileSystemObject")
' ファイルのOPEN
Set inputFile = fso.OpenTextFile(FolderName+inputFileName, conForReading, False)
If Err.Number <> 0 Then
' エラーメッセージを出力
wscript.echo "inputファイル名:" & FolderName+inputFileName
wscript.quit(1)
End If

Set inifso = WScript.CreateObject("Scripting.FileSystemObject")
' ファイルのOPEN
Wscript.echo "inputファイル名:" & FolderName+conIniFileName
Set objIniFile = inifso.OpenTextFile(FolderName+conIniFileName,conForReading,False)
If Err.Number <> 0 Then
' エラーメッセージを出力
Wscript.echo "inputファイル名:" & FolderName+conIniFileName
Wscript.quit(1)
End If

' ファイルのREAD
strReadLine = objIniFile.ReadLine
Do While objIniFile.AtEndofStream = False
' ステートメント開始行を検索
If (strReadLine <> " ") And (StrComp("[]", (Left(strReadLine, 1) & Right(strReadLine, 1))) = 0) Then
' セクション名を取得
strSection = Mid(strReadLine, 2, (Len(strReadLine) - 2))
' キー用Dictionaryオブジェクト作成
Set objKeyDic = CreateObject("Scripting.Dictionary")
' ファイルの最終行になるまでLoop
Do While objIniFile.AtEndofStream = False
strReadLine = objIniFile.ReadLine
If (strReadLine <> "") And (StrComp(";", Left(strReadLine, 1)) <> 0) Then
' 次のステートメント開始行が出現したら、Loop終了
If StrComp("[]", (Left(strReadLine, 1) & Right(strReadLine, 1))) = 0 Then
Exit Do
End If
' １セクション内の定義をDictionaryオブジェクトに格納する
arrReadLine = Split(strReadLine, "=", 2, vbTextCompare)
objKeyDic.Add UCase(arrReadLine(0)), arrReadLine(1)
End If
Loop
' オブジェクトに格納する
objSectionDic.Add UCase(strSection), objKeyDic
Else
strReadLine = objIniFile.ReadLine
End If
Loop
 
' ファイルのCLOSE
objIniFile.Close



dim logfso
dim logf
dim strDate
 
set logfso = CreateObject("Scripting.FileSystemObject")
strDate = Replace(FormatDateTime(Date(),0),"/","")+Replace(FormatDateTime(Time(),0),":","")

set logf = fso.OpenTextFile(FolderName+"log\exec" + StrDate + ".log", 8, True)

Do Until inputFile.AtEndOfStream
  Dim registerUrl
  Dim nanacoCode
  nanacoCode = inputFile.ReadLine 

Call use_ie(nanacoCode)

Loop

logf.Close
Set logfso = Nothing 
Set logf = Nothing
Set logfile = Nothing
Set strDate = Nothing
Set nanacoCode = Nothing


Sub use_ie(nanacoCode)

' １．登録用URL（PC）をIEで開く
    ' IE起動
    Set ie = CreateObject("InternetExplorer.Application")
    registerUrl = "https://www.nanaco-net.jp/pc/emServlet?gid="+nanacoCode
    ie.Navigate registerUrl
    ie.Visible = True
    waitIE ie

' ２．「nanaco番号」と「カード記載の番号」を自動入力し、ログインボタンをクリック
    ' nanaco番号を入力
    ie.Document.getElementById("nanacoNumber02").Value = funcIniFileGetString(objSectionDic, "ma", "nanaco16")
    WScript.Sleep 100
    
    ' カード記載の番号を入力
    ie.Document.getElementById("cardNumber").Value = funcIniFileGetString(objSectionDic, "ma", "nanaco7")
    WScript.Sleep 100
    
    ' ログインボタンクリック
    ie.Document.all("loginPass02").Click
    waitIE ie
     WScript.Sleep 100
    
 ' ３．会員メニューの「nanacoギフト登録」をクリック
For Each anchor In ie.document.getElementsByTagName("A")
If InStr(anchor.innerText, "nanacoギフト登録") > 0 Then
anchor.Click
Exit For
End If
Next
WScript.Sleep 1000
     
     
' ４．「ご利用約款に同意の上、登録」をクリック
For Each anchorss In ie.document.getElementsByTagName("input")
If InStr(anchorss.outerHtml, "/member/image/gift100/btn_400.gif") > 0 Then
anchorss.Click
Exit For
End If
Next
WScript.Sleep 1000
     
     
' ５．ギフトID登録フォームで「確認画面へ」ボタンをクリック。（ギフトIDは登録用URLをクリックしていれば自動入力）
    Set objShell = CreateObject("Shell.Application")
    Set objIE2 = objShell.Windows(objShell.Windows.Count - 1) '新しいウィンドウのオブジェクトを取得
    objIE2.Visible = True 'True：IEを表示 , False：IEを非表示
     WScript.Sleep 100
For Each anchorss In objIE2.document.getElementsByTagName("input")
If InStr(anchorss.outerHtml, "submit-button") > 0 Then
anchorss.Click
Exit For
End If
Next
WScript.Sleep 1000


' ６．最後に登録を押して、IEの画面を閉じる
  ' ギフトID登録内容確認画面で「登録する」ボタンを押す 
For Each anchorss In objIE2.document.getElementsByTagName("input")
If InStr(anchorss.outerHtml, "登録する") > 0 Then
anchorss.Click
logf.WriteLine( Date() & " " & Time() & ": "+nanacoCode)
Exit For
End If
Next

WScript.Sleep 1000

' 制御を破棄
    ie.Quit
    Set ie = Nothing
    
    objIE2.Quit
    Set objIE2 = Nothing
    

End Sub

' IEがビジー状態の間待ちます
Sub waitIE(ie)
    
    Do While ie.Busy = True Or ie.readystate <> 4
        WScript.Sleep 100
    Loop
    
    WScript.Sleep 1000
 
End Sub

Function funcIniFileGetString(objDictionary, strSection, strKey)
 
Dim objTempdic
 
strSection = UCase(strSection)
strKey = UCase(strKey)
 
If objDictionary.Exists(strSection) Then
Set objTempdic = objDictionary.Item(strSection)
If objTempdic.Exists(strKey) Then
funcIniFileGetString = objDictionary(strSection)(strKey)
Exit Function
End If
End If
 
funcIniFileGetString = ""
 
End Function
