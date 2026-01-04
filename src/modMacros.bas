Attribute VB_Name = "modMacros"
Option Explicit

' サクラエディタ インストールキー
Private Const REGKEY_SAKURA_INSTALL As String = "HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\sakura editor_is1\InstallLocation"
' サクラエディタ ファイル名
Private Const SAKURA_FILE_NAME      As String = "sakura.exe"

'-------------------------------------------------------------------------------
' 関数名 : OpenSakura
' 概要   : 結果シートのカレント行をサクラエディタで開く。
' 引数   : なし
' 戻り値 : なし
'-------------------------------------------------------------------------------
Public Sub OpenSakura()
    Dim sakuraInstallPath   As String
    Dim sakuraFilePath      As String
    Dim ws                  As Worksheet
    Dim wsh                 As Object
    
    Dim filePath            As String
    Dim pos                 As String
    
    Dim cmdOption           As String
    Dim cmd                 As String
    
    Dim arr                 As Variant
    
    If LCase(ActiveWorkbook.Name) <> LCase(ThisWorkbook.Name) Then
    
        Exit Sub
    End If
    
    Set wsh = CreateObject("WScript.Shell")
    
    On Error Resume Next
    sakuraInstallPath = wsh.RegRead(REGKEY_SAKURA_INSTALL)
    On Error GoTo 0
    
    sakuraFilePath = CombinePath(sakuraInstallPath, SAKURA_FILE_NAME)
    
    If Len(RTrim(sakuraInstallPath)) = 0 Or Len(Dir(sakuraFilePath, vbNormal)) = 0 Then
        
        MsgBox "サクラエディタがインストールされていません。", vbExclamation + vbOKOnly
        
        GoTo Exit_OpenSakura
    End If
    
    Set ws = ThisWorkbook.Sheets(WSNM_RESULT)
    
    ws.Activate
    
    filePath = ws.Range(ADDR_RESULT_PATH).Value
    pos = ws.Cells(ActiveCell.Row, COLIDX_RESULT_POSITION).Value
    
    If Len(filePath) = 0 Then
    
        GoTo Exit_OpenSakura
    End If
    
    If Len(Dir(filePath, vbNormal)) = 0 Then
    
        MsgBox "ファイルが見つかりません。", vbExclamation + vbOKOnly
    
        GoTo Exit_OpenSakura
    End If
    
    cmd = """" & sakuraFilePath & """"
    
    If Len(pos) > 0 Then
        
        arr = Split(pos, ",")
        
        If UBound(arr) >= 1 Then
        
            cmd = cmd & " -Y=" & arr(0) & " -X=" & arr(1)
        End If
    End If
    
    cmd = cmd & " """ & filePath & """"
    
    Shell cmd, vbNormalFocus
    
Exit_OpenSakura:

    Set wsh = Nothing
End Sub
