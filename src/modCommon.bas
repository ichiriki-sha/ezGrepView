Attribute VB_Name = "modCommon"
Option Explicit

'WinApt
Private Declare PtrSafe Function GetWindowThreadProcessId Lib "user32" ( _
                                    ByVal hWnd As LongPtr, ByRef lpdwProcessId As Long) As Long
                                    
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)

' 文字コード
Public Const ENCODING_UTF8  As String = "UTF-8"
Public Const ENCODING_SJIS  As String = "Shift_JIS"

'-------------------------------------------------------------------------------
' 関数名 : RemovePathSeparator
' 概要   : パス文字列の末尾にある区切り文字を削除します。
'          例）"C:\Work\" → "C:\Work"
' 引数   : folderPath    - 対象のパス文字列
'          pathSeparator - パス区切り文字（省略時 "\"）
' 戻り値 : 末尾の区切り文字を除去したパス文字列
'-------------------------------------------------------------------------------
Public Function RemovePathSeparator(ByVal folderPath As String, Optional ByVal pathSeparator = "\") As String

    If Len(RTrim(folderPath)) = 0 Then
        
        RemovePathSeparator = folderPath
        
        Exit Function
    End If

    If Right(folderPath, 1) = pathSeparator Then
        
        folderPath = Left(folderPath, Len(folderPath) - 1)
    End If

    RemovePathSeparator = folderPath
End Function

'-------------------------------------------------------------------------------
' 関数名 : AddPathSeparator
' 概要   : パス文字列の末尾に区切り文字を付加します。
'          既に付加されている場合は何もしません。
'          例）"C:\Work" → "C:\Work\"
' 引数   : folderPath    - 対象のパス文字列
'          pathSeparator - パス区切り文字（省略時 "\"）
' 戻り値 : 末尾に区切り文字を付加したパス文字列
'-------------------------------------------------------------------------------
Public Function AddPathSeparator(ByVal folderPath As String, Optional ByVal pathSeparator = "\") As String

    If Len(RTrim(folderPath)) = 0 Then
        
        AddPathSeparator = folderPath
        
        Exit Function
    End If

    If Right(folderPath, 1) <> pathSeparator Then
        
        folderPath = folderPath & pathSeparator
    End If

    AddPathSeparator = folderPath
End Function

'-------------------------------------------------------------------------------
' 関数名 : CombinePath
' 概要   : フォルダパスとファイル名を結合してフルパスを生成します。
'          フォルダパス末尾の区切り文字有無を吸収します。
' 引数   : folderPath    - フォルダパス
'          fileName      - ファイル名
'          pathSeparator - パス区切り文字（省略時 "\"）
' 戻り値 : 結合されたフルパス文字列
'-------------------------------------------------------------------------------
Public Function CombinePath(ByVal folderPath As String, ByVal FileName As String, Optional ByVal pathSeparator = "\") As String

    If Len(RTrim(folderPath)) = 0 And Len(RTrim(FileName)) = 0 Then
        
        CombinePath = folderPath
        
        Exit Function
    End If

    CombinePath = AddPathSeparator(folderPath, pathSeparator) & FileName
End Function

'-------------------------------------------------------------------------------
' 関数名 : GetParentPath
' 概要   : 指定されたパスから親フォルダのパスを取得します。
'          例）"C:\Work\Test.txt" → "C:\Work"
' 引数   : filePath      - 対象のパス文字列
'          pathSeparator - パス区切り文字（省略時 "\"）
' 戻り値 : 親フォルダのパス
'          親が存在しない場合は空文字を返します
'-------------------------------------------------------------------------------
Public Function GetParentPath(ByVal filePath As String, Optional ByVal pathSeparator = "\") As String
    Dim pos     As Long
    Dim parent  As String

    filePath = RemovePathSeparator(filePath)
    
    pos = InStrRev(filePath, pathSeparator)
    
    If pos > 0 Then
    
        parent = Left(filePath, pos - 1)
    End If

    GetParentPath = parent
End Function

'-------------------------------------------------------------------------------
' 関数名 : GetFileName
' 概要   : 指定されたパスからファイル名を取得します。
'          例）"C:\Work\Test.txt" → "Test.txt"
' 引数   : filePath      - 対象のパス文字列
'          pathSeparator - パス区切り文字（省略時 "\"）
' 戻り値 : ファイル名
'-------------------------------------------------------------------------------
Public Function GetFileName(ByVal filePath As String, Optional ByVal pathSeparator = "\") As String
    Dim pos     As Long
    Dim file    As String

    file = filePath
    
    pos = InStrRev(filePath, pathSeparator)
    
    If pos > 0 Then
    
        file = mid(filePath, pos + 1)
    End If

    GetFileName = file
End Function

'-------------------------------------------------------------------------------
' 関数名 : GetExtension
' 概要   : 指定されたパスから拡張子を取得します。
'          例）"Test.txt" → "txt"
' 引数   : filePath      - 対象のパス文字列
'          pathSeparator - パス区切り文字（省略時 "\"）
' 戻り値 : 拡張子（ドットなし）
'          拡張子が存在しない場合は空文字を返します
'-------------------------------------------------------------------------------
Public Function GetExtension(ByVal filePath As String, Optional ByVal pathSeparator = "\") As String
    Dim pos     As Long
    Dim file    As String
    Dim ext     As String

    ext = ""
    
    file = GetFileName(filePath, pathSeparator)
    
    pos = InStrRev(file, ".")
    
    If pos > 0 Then
    
        ext = mid(file, pos + 1)
    End If

    GetExtension = ext
End Function

'-------------------------------------------------------------------------------
' 関数名 : GetTempDirectory
' 概要   : OS が提供する一時フォルダのパスを取得します。
' 引数   : なし
' 戻り値 : 一時フォルダのパス
'-------------------------------------------------------------------------------
Public Function GetTempDirectory() As String

    GetTempDirectory = Environ("TEMP")
End Function

'-------------------------------------------------------------------------------
' 関数名 : ExpandPathToEnvVar
' 概要   : 実パスを環境変数表記（%USERPROFILE% など）に変換する
'          指定されたパスが環境変数の値で始まっている場合のみ置換する
' 引数   : filePath - 変換対象のフルパス
' 戻り値 : 環境変数表記に変換されたパス
'-------------------------------------------------------------------------------
Public Function ExpandPathToEnvVar(ByVal filePath As String) As String
    Dim envNames    As Variant
    Dim envName     As Variant
    Dim envValue    As String
    
    envNames = Array("OneDrive", "TEMP", "LOCALAPPDATA", "APPDATA", "USERPROFILE", "ProgramData", "ProgramFiles", "ProgramFiles(x86)", "SystemRoot")
    
    For Each envName In envNames
    
        envValue = Environ(envName)
            
        If InStr(1, filePath, envValue, vbTextCompare) = 1 Then
        
            filePath = Replace(filePath, envValue, "%" & envName & "%", compare:=vbTextCompare)
            
            Exit For
        End If
    Next
    
    ExpandPathToEnvVar = filePath
End Function

'-------------------------------------------------------------------------------
' 関数名 : ExpandPathFromEnvVar
' 概要   : 環境変数表記（%USERPROFILE% など）を実パスに変換する
'          パスの先頭にある環境変数のみを展開する
' 引数   : filePath - 環境変数表記を含むパス
' 戻り値 : 実パスに変換されたパス
'-------------------------------------------------------------------------------
Public Function ExpandPathFromEnvVar(ByVal filePath As String) As String
    Dim envNames    As Variant
    Dim envName     As Variant
    Dim envValue    As String
    
    envNames = Array("OneDrive", "TEMP", "LOCALAPPDATA", "APPDATA", "USERPROFILE", "ProgramData", "ProgramFiles", "ProgramFiles(x86)", "SystemRoot")
    
    For Each envName In envNames
    
        If InStr(1, filePath, "%" & envName & "%", vbTextCompare) = 1 Then
        
            envValue = Environ(envName)
        
            filePath = Replace(filePath, "%" & envName & "%", envValue, compare:=vbTextCompare)
            
            Exit For
        End If
    Next
    
    ExpandPathFromEnvVar = filePath
End Function

'-------------------------------------------------------------------------------
' 関数名 : ShowOpenFileDialog
' 概要   : ファイル選択ダイアログを表示し、選択されたファイルパスを取得する
' 引数   : filePath - 選択されたファイルパス（参照渡し）
'        : filter   - ファイルフィルタ（省略可）
'        : caption  - ダイアログのタイトル（省略可）
' 戻り値 : ファイルが選択された場合 True、キャンセルされた場合 False
'-------------------------------------------------------------------------------
Public Function ShowOpenFileDialog(ByRef filePath As String, _
                                   Optional filter As String = "テキストファイル(*.txt),*.txt", _
                                   Optional caption As String = "ファイルを開く") As Boolean
    Dim result  As Boolean
    Dim ret     As String
    
    ret = Application.GetOpenFilename(FileFilter:=filter, Title:=caption)
    
    If LCase(ret) = "false" Then
        
        result = False
        filePath = ""
    Else
        
        result = True
        filePath = ret
    End If

    ShowOpenFileDialog = result
End Function

'-------------------------------------------------------------------------------
' 関数名 : GetProcessIdFromHwnd
' 概要   : 指定したウィンドウハンドル（hWnd）から
'          対応するプロセスID（PID）を取得する。
'          Win32 API GetWindowThreadProcessId のラッパー。
'
' 引数   : hWnd - ウィンドウハンドル
' 戻り値 : プロセスID
'          ・取得成功時 : PID
'          ・失敗時     : 0
'-------------------------------------------------------------------------------
Public Function GetProcessIdFromHwnd(ByVal hWnd As LongPtr) As Long
    Dim pid As Long
    Dim threadId As Long

    ' 不正なhWndチェック
    If hWnd = 0 Then
        GetProcessIdFromHwnd = 0
        Exit Function
    End If

    ' hWnd からプロセスID取得
    threadId = GetWindowThreadProcessId(hWnd, pid)

    ' 取得失敗
    If threadId = 0 Then
        GetProcessIdFromHwnd = 0
        Exit Function
    End If

    GetProcessIdFromHwnd = pid
End Function

'-------------------------------------------------------------------------------
' 関数名 : WaitMilliseconds
' 概要   : 指定したミリ秒数だけ待機する。
'          Sleep を短時間ループで呼び出し、
'          DoEvents を挟むことで UI フリーズを防ぐ。
'
' 引数   : milliseconds - 待機する時間（ミリ秒）
'-------------------------------------------------------------------------------
Public Sub WaitMilliseconds(ByVal milliseconds As Long)
    Dim remain As Long
    
    If milliseconds <= 0 Then Exit Sub
    
    remain = milliseconds
    
    Do While remain > 0
        Sleep 50
        DoEvents
        remain = remain - 50
    Loop
End Sub

'-------------------------------------------------------------------------------
' 関数名 : IsProcessRunning
' 概要   : 指定した PID のプロセスが実行中かどうかを判定する
' 引数   : pid - プロセス ID
' 戻り値 : 実行中の場合 True、存在しない場合 False
'-------------------------------------------------------------------------------
Public Function IsProcessRunning(ByVal pid As Long) As Boolean
    Dim wmi         As Object
    Dim process     As Object

    Set wmi = GetObject("winmgmts:")
    Set process = wmi.ExecQuery("SELECT * FROM Win32_Process WHERE ProcessId =" & pid)
    
    IsProcessRunning = (process.Count > 0)
    
    Set process = Nothing
    Set wmi = Nothing
End Function

'-------------------------------------------------------------------------------
' 関数名 : KillProcess
' 概要   : 指定した PID のプロセスを強制終了する
' 引数   : pid - プロセス ID
' 戻り値 : なし
' 備考   : taskkill を使用するため管理者権限が必要な場合あり
'-------------------------------------------------------------------------------
Public Sub KillProcess(ByVal pid As Long)

    Shell "taskkill /PID " & pid & " /F", vbHide
End Sub

'-------------------------------------------------------------------------------
' 関数名 : ReadAllLines
' 概要   : 指定したテキストファイルをすべて読み込み、
'          改行（CRLF）で分割した配列として返す
' 引数   : filePath - 読み込むファイルのフルパス
'        : enc      - 文字コード（省略可）
'                    ENCODING_UTF8 : UTF-8
'                    ENCODING_SJIS : Shift_JIS
' 戻り値 : 各行を要素とする文字列配列（Variant）
'-------------------------------------------------------------------------------
Public Function ReadAllLines(ByVal filePath As String, Optional ByVal enc As String = ENCODING_UTF8) As String()
    Dim stream  As Object
    Dim data    As String
    
    Const adReadAll As Long = -1

    Set stream = CreateObject("ADODB.Stream")
    With stream
        .Charset = IIf(enc = ENCODING_UTF8, "utf-8", "shift_jis")
        .Open
        .LoadFromFile filePath
        data = .ReadText(adReadAll)
        .Close
    End With
    
    Set stream = Nothing

    ReadAllLines = Split(data, vbCrLf)
End Function

'-------------------------------------------------------------------------------
' 関数名 : EscapePattern
' 概要   : 文字列を正規表現で安全に使用できるように、
'          特殊文字をエスケープする
'
'          エスケープ対象：
'          \ . + * ? ^ $ ( ) [ ] { } |
'
' 引数   : pattern - エスケープ対象の文字列
' 戻り値 : 正規表現用にエスケープされた文字列
'-------------------------------------------------------------------------------
Public Function EscapePattern(ByVal pattern As String) As String
    Dim specialChars As Variant
    Dim c As Variant
    
    specialChars = Array("\", ".", "+", "*", "?", "^", "$", "(", ")", "[", "]", "{", "}", "|")
    
    For Each c In specialChars
        
        pattern = Replace(pattern, c, "\" & c)
    Next

    EscapePattern = pattern
End Function

'-------------------------------------------------------------------------------
' 関数名 : GetLocalPathFromUrl
' 概要   : OneDrive 同期フォルダ内のファイルが
'          「https://d.docs.live.net/～」形式の URL として
'          取得された場合に、対応するローカルファイルパスへ変換する。
'
'          URL 形式でない場合は、そのまま引数を返す。
'
' 引数   : urlPath - OneDrive ファイルの URL または通常のローカルパス
'
' 戻り値 : ローカルファイルシステム上のパス
'          （例: C:\Users\xxx\OneDrive\...）
'
' 補足   : Excel で OneDrive 同期フォルダ内のファイルを参照した際、
'          ThisWorkbook.Path 等が URL 形式になる問題への対策用。
'          OneDrive の CID は GetOneDriveCID で取得する。
'-------------------------------------------------------------------------------
Public Function GetLocalPathFromUrl(ByVal urlPath As String) As String
    Dim oneDrivePath As String
    oneDrivePath = Environ("OneDrive")

    If Left(urlPath, 8) = "https://" Then
        ' URL部分を除去してローカルパスへ変換
        
        GetLocalPathFromUrl = Replace(urlPath, _
                                        "https://d.docs.live.net/" & GetOneDriveCID() & "/", _
                                         oneDrivePath & "\")
    Else
        
        GetLocalPathFromUrl = urlPath
    End If
End Function

'-------------------------------------------------------------------------------
' 関数名 : GetOneDriveCID
' 概要   : ローカル PC に設定されている OneDrive の
'          アカウント CID（Client ID）を取得する。
'
'          CID は OneDrive の URL
'          https://d.docs.live.net/{CID}/...
'          の識別子として使用される。
'
' 処理内容 : OneDrive の設定ファイル（global.ini）を読み込む
'            「cid=xxxxxx」の行を正規表現で検索
'
' 戻り値 : CID 文字列
'          取得できなかった場合は空文字 ("")
'
' 注意   : 個人用 OneDrive（Personal）を前提としている
'          設定ファイルが存在しない場合でもエラーにはならない
'-------------------------------------------------------------------------------
Private Function GetOneDriveCID() As String
    Dim fso     As Object
    Dim iniPath As String
    Dim ret     As String
    
    ret = ""
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    iniPath = Environ("LocalAppData") & "\Microsoft\OneDrive\settings\Personal\global.ini"

    If fso.FileExists(iniPath) Then
        
        Dim ts As Object
        Dim text As String

        Dim re As Object
        
        Set ts = fso.OpenTextFile(iniPath, 1, False, -1)
        Set re = CreateObject("VBScript.RegExp")
        
        re.pattern = "cid\s*=\s*(\w+)"
        re.IgnoreCase = True

        Do Until ts.AtEndOfStream
        
            text = ts.ReadLine ' 1行読み込み
        
            If re.Test(text) Then
                
                ret = re.Execute(text)(0).SubMatches(0)
                
                Exit Do
            End If
        Loop

        ts.Close
    End If
    
    GetOneDriveCID = ret
End Function

'-------------------------------------------------------------------------------
' 関数名 : FindWorksheet
' 概要   : 指定された名前のワークシートが、
'          対象の Workbook 内に存在するかを判定する。
'
' 引数   : wsName - 検索対象のワークシート名
'          wb     - 検索対象の Workbook（省略時は ThisWorkbook）
'
' 戻り値 : True  - 指定名のワークシートが存在する
'          False - 存在しない
'
' 備考   : Worksheets(wsName) 参照時にエラーが発生する可能性があるため、
'          On Error Resume Next を使用して存在チェックを行っている。
'-------------------------------------------------------------------------------
Public Function FindWorksheet(ByVal wsName As String, Optional ByVal wb As Workbook) As Boolean
    Dim ws As Worksheet
    
    If wb Is Nothing Then Set wb = ThisWorkbook
    
    On Error Resume Next
    Set ws = wb.Worksheets(wsName)
    On Error GoTo 0
    
    FindWorksheet = Not ws Is Nothing
    Set ws = Nothing
End Function
