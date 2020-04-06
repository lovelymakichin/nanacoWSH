'Dim anchor As HTMLAnchorElement
' �t�@�C�����o�͂̒萔
Const conForReading = 1, conForWriting = 2, conForAppending = 8

' ini�t�@�C����
Const FolderName = "D:\Program\nanacoWSH\"
Const conIniFileName = "nanaco.ini"
Const inputFileName = "nanacoCode.txt"

' �i�[��Dictionary�I�u�W�F�N�g�̍쐬
Set objSectionDic = CreateObject("Scripting.Dictionary")



Set fso = WScript.CreateObject("Scripting.FileSystemObject")
' �t�@�C����OPEN
Set inputFile = fso.OpenTextFile(FolderName+inputFileName, conForReading, False)
If Err.Number <> 0 Then
' �G���[���b�Z�[�W���o��
wscript.echo "input�t�@�C����:" & FolderName+inputFileName
wscript.quit(1)
End If

Set inifso = WScript.CreateObject("Scripting.FileSystemObject")
' �t�@�C����OPEN
Wscript.echo "input�t�@�C����:" & FolderName+conIniFileName
Set objIniFile = inifso.OpenTextFile(FolderName+conIniFileName,conForReading,False)
If Err.Number <> 0 Then
' �G���[���b�Z�[�W���o��
Wscript.echo "input�t�@�C����:" & FolderName+conIniFileName
Wscript.quit(1)
End If

' �t�@�C����READ
strReadLine = objIniFile.ReadLine
Do While objIniFile.AtEndofStream = False
' �X�e�[�g�����g�J�n�s������
If (strReadLine <> " ") And (StrComp("[]", (Left(strReadLine, 1) & Right(strReadLine, 1))) = 0) Then
' �Z�N�V���������擾
strSection = Mid(strReadLine, 2, (Len(strReadLine) - 2))
' �L�[�pDictionary�I�u�W�F�N�g�쐬
Set objKeyDic = CreateObject("Scripting.Dictionary")
' �t�@�C���̍ŏI�s�ɂȂ�܂�Loop
Do While objIniFile.AtEndofStream = False
strReadLine = objIniFile.ReadLine
If (strReadLine <> "") And (StrComp(";", Left(strReadLine, 1)) <> 0) Then
' ���̃X�e�[�g�����g�J�n�s���o��������ALoop�I��
If StrComp("[]", (Left(strReadLine, 1) & Right(strReadLine, 1))) = 0 Then
Exit Do
End If
' �P�Z�N�V�������̒�`��Dictionary�I�u�W�F�N�g�Ɋi�[����
arrReadLine = Split(strReadLine, "=", 2, vbTextCompare)
objKeyDic.Add UCase(arrReadLine(0)), arrReadLine(1)
End If
Loop
' �I�u�W�F�N�g�Ɋi�[����
objSectionDic.Add UCase(strSection), objKeyDic
Else
strReadLine = objIniFile.ReadLine
End If
Loop
 
' �t�@�C����CLOSE
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

' �P�D�o�^�pURL�iPC�j��IE�ŊJ��
    ' IE�N��
    Set ie = CreateObject("InternetExplorer.Application")
    registerUrl = "https://www.nanaco-net.jp/pc/emServlet?gid="+nanacoCode
    ie.Navigate registerUrl
    ie.Visible = True
    waitIE ie

' �Q�D�unanaco�ԍ��v�Ɓu�J�[�h�L�ڂ̔ԍ��v���������͂��A���O�C���{�^�����N���b�N
    ' nanaco�ԍ������
    ie.Document.getElementById("nanacoNumber02").Value = funcIniFileGetString(objSectionDic, "ma", "nanaco16")
    WScript.Sleep 100
    
    ' �J�[�h�L�ڂ̔ԍ������
    ie.Document.getElementById("cardNumber").Value = funcIniFileGetString(objSectionDic, "ma", "nanaco7")
    WScript.Sleep 100
    
    ' ���O�C���{�^���N���b�N
    ie.Document.all("loginPass02").Click
    waitIE ie
     WScript.Sleep 100
    
 ' �R�D������j���[�́unanaco�M�t�g�o�^�v���N���b�N
For Each anchor In ie.document.getElementsByTagName("A")
If InStr(anchor.innerText, "nanaco�M�t�g�o�^") > 0 Then
anchor.Click
Exit For
End If
Next
WScript.Sleep 1000
     
     
' �S�D�u�����p�񊼂ɓ��ӂ̏�A�o�^�v���N���b�N
For Each anchorss In ie.document.getElementsByTagName("input")
If InStr(anchorss.outerHtml, "/member/image/gift100/btn_400.gif") > 0 Then
anchorss.Click
Exit For
End If
Next
WScript.Sleep 1000
     
     
' �T�D�M�t�gID�o�^�t�H�[���Łu�m�F��ʂցv�{�^�����N���b�N�B�i�M�t�gID�͓o�^�pURL���N���b�N���Ă���Ύ������́j
    Set objShell = CreateObject("Shell.Application")
    Set objIE2 = objShell.Windows(objShell.Windows.Count - 1) '�V�����E�B���h�E�̃I�u�W�F�N�g���擾
    objIE2.Visible = True 'True�FIE��\�� , False�FIE���\��
     WScript.Sleep 100
For Each anchorss In objIE2.document.getElementsByTagName("input")
If InStr(anchorss.outerHtml, "submit-button") > 0 Then
anchorss.Click
Exit For
End If
Next
WScript.Sleep 1000


' �U�D�Ō�ɓo�^�������āAIE�̉�ʂ����
  ' �M�t�gID�o�^���e�m�F��ʂŁu�o�^����v�{�^�������� 
For Each anchorss In objIE2.document.getElementsByTagName("input")
If InStr(anchorss.outerHtml, "�o�^����") > 0 Then
anchorss.Click
logf.WriteLine( Date() & " " & Time() & ": "+nanacoCode)
Exit For
End If
Next

WScript.Sleep 1000

' �����j��
    ie.Quit
    Set ie = Nothing
    
    objIE2.Quit
    Set objIE2 = Nothing
    

End Sub

' IE���r�W�[��Ԃ̊ԑ҂��܂�
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
