Option Explicit

'���T�[�o�[�A�h���X���̂P
Const Search1 = "192.168.0.102"
'�V�T�[�o�[�A�h���X���̂P
Const Replace1 = "192.168.0.202"


'���T�[�o�[�A�h���X���̂Q
Const Search2 = "192.168.0.103"
'�V�T�[�o�[�A�h���X���̂Q
Const Replace2 = "192.168.0.203"

'�I���N���C���X�g�[���p�X�iORACLE_HOME�j
Const oraHome = "C:\Program Files (x86)\app\oracle"
Const ora11Key = "OraClient11g_home1"



'tnsnames.ora�t�@�C���̃t�H���_
Const orgFilePath = "\network\admin\"
'tnsnames.ora�t�@�C����
Const orgFile = "tnsnames.ora"
'tnsnames.ora�o�b�N�A�b�v�t�@�C����
Const bakFile = "tnsnames.ora.org"

'Oracle Home�擾����
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

'���ݓ������擾����
Private Function getSysDateTime()
    Dim strFormattedDate
     
    'yyyy/mm/dd hh:mm:ss �`���̕�����Ō��ݓ������擾
    strFormattedDate = Now()
     
    'yyyy/mm/dd hh:mm:ss ���� / ���폜
    strFormattedDate = Replace(strFormattedDate, "/", "")
     
    'yyyy/mm/dd hh:mm:ss ���� : ���폜
    strFormattedDate = Replace(strFormattedDate, ":", "")
     
    'yyyy/mm/dd hh:mm:ss ����X�y�[�X���폜
    strFormattedDate = Replace(strFormattedDate, " ", "")
    getSysDateTime = strFormattedDate
End Function

'�T�[�o�[�A�h���X�u��
Private Function FindAndReplace(ByVal line, ByVal search, ByVal replaceMent)
    Dim objRegExp       ' ���K�\���I�u�W�F�N�g
    Dim strRepBefore    ' �u���O�̕�����
    Dim strRepAfter     ' �u����̕�����

    Set objRegExp = CreateObject("VBScript.RegExp")
    objRegExp.Pattern = "HOST[^0-9]+" & search
    objRegExp.IgnoreCase  = true '�啶���Ə���������ʂ��Ȃ�

    strRepAfter = line
    If objRegExp.Execute(line).Count > 0 Then
        strRepAfter = objRegExp.Replace(line, "HOST = " & replaceMent)
        'WScript.Echo "�u����̕������ " & strRepAfter & " �ł��B"
    End If

    Set objRegExp = Nothing

    FindAndReplace = strRepAfter
End Function

'�t�@�C���e�L�X�g�ҏW
Private Sub ReplaceHostAddress(ByVal objFso, ByVal orgPath)
    Dim reader,writer
    Dim writeText, line

    '�ǎ惂�[�h�Ńe�L�X�g���J��
    Set reader = objFso.OpenTextFile(orgPath , 1)

    Do Until reader.AtEndOfStream = True    '�I���s�܂ŌJ��Ԃ�
        line = reader.ReadLine 
        line = FindAndReplace(line, Search1, Replace1)
        line = FindAndReplace(line, Search2, Replace2)
        
        writeText = writeText & line & vbCrLf
    Loop

    reader.Close  '�ǎ惂�[�h����
    Set reader = Nothing
    
    Set writer = objFso.OpenTextFile(orgPath , 2)
    writer.Write  writeText  '�t�@�C����������
    writer.Close
    Set writer = Nothing
End Sub

'���C������
Private Sub procecc(ByVal oracleHome)
    Dim objFso
    Dim orgPath, targetPath
    
    orgPath = oracleHome & orgFilePath & orgFile
    targetPath = oracleHome & orgFilePath & bakFile & "." & getSysDateTime

    Set objFso = CreateObject("Scripting.FileSystemObject")
    
    'tnsnames.ora�t�@�C����������Ȃ��ꍇ�������~
    If objFso.FileExists(orgPath) = False Then
        WScript.Echo "[" & orgPath & "]�t�@�C����������܂���B"
        Exit Sub
    End If

    '�I���W�i���t�@�C���o�b�N�A�b�v
    Call objFso.CopyFile(orgPath, targetPath, False)
    'tnsnames.ora�t�@�C���ҏW
    Call ReplaceHostAddress(objFso, orgPath)
    
    Set objFso = Nothing
End Sub



'�J�n���\�b�h
Sub main()
    Call procecc(oraHome)
    WScript.Echo "tnsnames.ora�t�@�C����ύX���܂����B"
End Sub

'�G���[�������ɂ������𑱍s����悤�ݒ�
On Error Resume Next

Dim objShell
Dim oracleHome

Set objShell = CreateObject("Shell.Application")
If Wscript.Arguments.Count = 0 then
    '�o�b�`�t�@�C�����Ǘ��Ҍ����Ŏ��s����
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

'�G���[�ɂȂ����ꍇ�̏���
If Err.Number <> 0 Then
  
  WScript.Echo "�\�����Ȃ��G���[���������܂����B" & vbCrLf & "�G���[�ԍ��F" & Err.Number & vbCrLf & "�G���[�ڍׁF" & Err.Description
  Err.Clear
End If