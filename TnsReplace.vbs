Option Explicit

'旧サーバーアドレスその１
Const Search1 = "192.168.0.102"
'新サーバーアドレスその１
Const Replace1 = "192.168.0.202"


'旧サーバーアドレスその２
Const Search2 = "192.168.0.103"
'新サーバーアドレスその２
Const Replace2 = "192.168.0.203"

'オラクルインストールパス（ORACLE_HOME）
Const oraHome = "C:\Program Files (x86)\app\oracle"
Const ora11Key = "OraClient11g_home1"



'tnsnames.oraファイルのフォルダ
Const orgFilePath = "\network\admin\"
'tnsnames.oraファイル名
Const orgFile = "tnsnames.ora"
'tnsnames.oraバックアップファイル名
Const bakFile = "tnsnames.ora.org"

'Oracle Home取得する
Function GetOraHome(ByVal vHOME_NAME)
    Dim WSH, wExec
    Dim sCmd, Result, strPath
    Dim lngPos
    
    Set WSH = CreateObject("WScript.Shell")
    sCmd = "reg query HKEY_LOCAL_MACHINE\SOFTWARE\ORACLE\KEY_" & vHOME_NAME & " /v ""ORACLE_HOME"""
    sCmd = "reg query HKEY_LOCAL_MACHINE\SOFTWARE\ /f pain*" '/v ""Path"""
    Set wExec = WSH.Exec("%ComSpec% /c " & sCmd)
    Do While wExec.Status = 0
        WScript.Sleep 100
    Loop

    Result = wExec.StdOut.ReadAll
    lngPos = InStr(Result, vbCrLf)
    strPath = Trim(Mid(Result,lngPos + 2, InStr(lngPos + 1, Result, vbCrLf) - 3))
    'strPath = Trim(Mid(Result, InStr(Result, "REG_SZ") + 6, LenB(Result)))

    GetOraHome = strPath

    Set wExec = Nothing
    Set WSH = Nothing

End Function

'現在日時を取得する
Private Function getSysDateTime()
    Dim strFormattedDate
     
    'yyyy/mm/dd hh:mm:ss 形式の文字列で現在日時を取得
    strFormattedDate = Now()
     
    'yyyy/mm/dd hh:mm:ss から / を削除
    strFormattedDate = Replace(strFormattedDate, "/", "")
     
    'yyyy/mm/dd hh:mm:ss から : を削除
    strFormattedDate = Replace(strFormattedDate, ":", "")
     
    'yyyy/mm/dd hh:mm:ss からスペースを削除
    strFormattedDate = Replace(strFormattedDate, " ", "")
    getSysDateTime = strFormattedDate
End Function

'サーバーアドレス置換
Private Function FindAndReplace(ByVal line, ByVal search, ByVal replaceMent)
    Dim objRegExp       ' 正規表現オブジェクト
    Dim strRepBefore    ' 置換前の文字列
    Dim strRepAfter     ' 置換後の文字列

    Set objRegExp = CreateObject("VBScript.RegExp")
    objRegExp.Pattern = "HOST[^0-9]+" & search
    objRegExp.IgnoreCase  = true '大文字と小文字を区別しない

    strRepAfter = line
    If objRegExp.Execute(line).Count > 0 Then
        strRepAfter = objRegExp.Replace(line, "HOST = " & replaceMent)
        'WScript.Echo "置換後の文字列は " & strRepAfter & " です。"
    End If

    Set objRegExp = Nothing

    FindAndReplace = strRepAfter
End Function

'ファイルテキスト編集
Private Sub ReplaceHostAddress(ByVal objFso, ByVal orgPath)
    Dim reader,writer
    Dim writeText, line

    '読取モードでテキストを開く
    Set reader = objFso.OpenTextFile(orgPath , 1)

    Do Until reader.AtEndOfStream = True    '終了行まで繰り返し
        line = reader.ReadLine 
        line = FindAndReplace(line, Search1, Replace1)
        line = FindAndReplace(line, Search2, Replace2)
        
        writeText = writeText & line & vbCrLf
    Loop

    reader.Close  '読取モード閉じる
    Set reader = Nothing
    
    Set writer = objFso.OpenTextFile(orgPath , 2)
    writer.Write  writeText  'ファイル書き込み
    writer.Close
    Set writer = Nothing
End Sub

'メイン処理
Private Sub procecc(ByVal oracleHome)
    Dim objFso
    Dim orgPath, targetPath
    
    orgPath = oracleHome & orgFilePath & orgFile
    targetPath = oracleHome & orgFilePath & bakFile & "." & getSysDateTime

    Set objFso = CreateObject("Scripting.FileSystemObject")
    
    'tnsnames.oraファイルが見つからない場合処理中止
    If objFso.FileExists(orgPath) = False Then
        WScript.Echo "[" & orgPath & "]ファイルが見つかりません。"
        Exit Sub
    End If

    'オリジナルファイルバックアップ
    Call objFso.CopyFile(orgPath, targetPath, False)
    'tnsnames.oraファイル編集
    Call ReplaceHostAddress(objFso, orgPath)
    
    Set objFso = Nothing
End Sub



'開始メソッド
Sub main()
    Call procecc(oraHome)
    WScript.Echo "tnsnames.oraファイルを変更しました。"
End Sub

'エラー発生時にも処理を続行するよう設定
On Error Resume Next

Dim objShell
Dim oracleHome

Set objShell = CreateObject("Shell.Application")
If Wscript.Arguments.Count = 0 then
    'バッチファイルを管理者権限で実行する
    objShell.ShellExecute "wscript.exe", WScript.ScriptFullName & " runas", "", "runas", 1
    Wscript.Quit
end if

oracleHome = GetOraHome("OraClient11g_home1")
If oracleHome = "" Then
    WScript.Echo "Not Find Oracle_Home"
Else
    WScript.Echo "Oracle_Home = " & oracleHome
    WScript.Echo "Oracle_Home_Len = " & Len(oracleHome)
End If


'Call main()

'エラーになった場合の処理
If Err.Number <> 0 Then
  
  WScript.Echo "予期しないエラーが発生しました。" & vbCrLf & "エラー番号：" & Err.Number & vbCrLf & "エラー詳細：" & Err.Description
  Err.Clear
End If