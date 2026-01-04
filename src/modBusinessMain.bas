Attribute VB_Name = "modBusinessMain"
Option Explicit

'-------------------------------------------------------------------------------
' 関数名 : GenerateResuFilePath
'
' 概要   : grep結果を出力する Excel ファイルのフルパスを生成する。
'          ・検索キーワードをファイル名用に整形
'          ・長さ制限による省略
'          ・タイムスタンプを付与して一意な名前にする
'
' 引数   : keyword - grep時の検索キーワード
'
' 戻り値 : 出力用 Excel ファイルのフルパス
'-------------------------------------------------------------------------------
Private Function GenerateResuFilePath(ByVal pattern As String) As String
    Dim word            As String
    Dim file            As String
    Dim timeStamp       As String

    Const WORD_LENGTH   As Integer = 32
    
    timeStamp = Format(Now, "yyyyMMddHHmmss")
    word = Replace(StrConv(pattern, vbWide), "\", "＼")

    If Len(word) > WORD_LENGTH Then
    
        word = mid(word, 1, WORD_LENGTH - 1) & "…"
    End If

    file = "grep結果【" & word & "】_" & timeStamp & ".xlsm"
    
    GenerateResuFilePath = CombinePath(GetLocalPathFromUrl(ThisWorkbook.path), file)
End Function

'-------------------------------------------------------------------------------
' 関数名 : IsComment
'
' 概要   : 指定された文字列が「コメントのみで構成された行」かどうかを判定する。
'          拡張子ごとに定義されたコメント正規表現を使用し、
'          コメントを除去した結果が空文字になるかで判定する。
'
' 引数   : source          - 判定対象の1行分の文字列
'          ext             - ファイル拡張子（例: "c", "cpp", "vb" など）
'          CommentPattern - 拡張子ごとのコメント正規表現を格納した Dictionary
'
' 戻り値 : コメント行の場合 True、そうでない場合 False
'-------------------------------------------------------------------------------
Private Function IsComment(ByVal source As String, ByVal ext As String, ByVal commentPattern As Object) As Boolean
    Dim regexp          As Object
    Dim result          As Boolean
    Dim removeComment  As String
    
    result = False
    
    If commentPattern.Exists(ext) Then
    
        Set regexp = CreateObject("VBScript.RegExp")
        regexp.Global = False
        regexp.IgnoreCase = False
    
        regexp.pattern = commentPattern.Item(ext)
        removeComment = RTrim(regexp.Replace(Replace(source, vbTab, " "), ""))
        result = Len(removeComment) = 0
        
        Set regexp = Nothing
    End If
    
    IsComment = result
End Function

'-------------------------------------------------------------------------------
' 関数名 : IsBinary
'
' 概要   : 文字列にバイナリデータとみなせる制御文字が含まれているかを判定する。
'          NULL文字や改行を除く制御コード（0x00～0x08, 0x0A～0x1F, 0x7F）を
'          含む場合、その文字列はテキストではない可能性があると判断する。
'
' 引数   : source - 判定対象の文字列
'
' 戻り値 : 制御文字を含む場合 True、含まない場合 False
'-------------------------------------------------------------------------------
Private Function IsBinary(ByVal source As String) As Boolean
    Dim regexp As Object
    Dim result As Boolean
    
    Set regexp = CreateObject("VBScript.RegExp")
    regexp.Global = False
    regexp.IgnoreCase = False

    regexp.pattern = "[\x00-\x08\x0A-\x1F\x7F]"
  
    result = regexp.Test(source)
    
    Set regexp = Nothing
    
    IsBinary = result
End Function

'-------------------------------------------------------------------------------
' 関数名 : IsGarbled
'
' 概要   : 文字列に文字化けを示す代表的な文字が含まれているかを判定する。
'          ・0x19（EM / SUB など、エンコード失敗時に混入する制御文字）
'          ・U+FFFD（Replacement Character：?）
'          が含まれている場合、文字化けの可能性があると判断する。
'
' 引数   : source - 判定対象の文字列
'
' 戻り値 : 文字化けの疑いがある場合 True、問題なさそうな場合 False
'-------------------------------------------------------------------------------
Private Function IsGarbled(ByVal source As String) As Boolean
    Dim result As Boolean
    
    result = False
    
    If InStr(source, Chr(&H19)) > 0 Then
    
        result = True
    Else
    
        If InStr(source, ChrW(&HFFFD)) > 0 Then
        
            result = True
        End If
    End If
    
    IsGarbled = result
End Function

'-------------------------------------------------------------------------------
' 関数名 : TryParseGrepHeader
'
' 概要   : grep 結果ファイルの先頭ヘッダー部分を解析し、
'          検索条件（キーワード・対象ファイル・検索フォルダなど）を取得する
'          必要な情報がすべて取得できた場合のみ True を返す
'
' 引数   : filePath   - grep 結果ファイルのパス
'        : enc   - ファイルの文字コード（UTF-8 / Shift_JIS）
'        : grepHeader - 解析結果を格納する構造体（参照渡し）
'
' 戻り値 : 解析成功時 True、必要な情報が不足している場合 False
'-------------------------------------------------------------------------------
Private Function TryParseHeaderData(ByRef inData As InputData, _
                                    ByRef grepHeader As GrepHeaderData) As Boolean
    Dim buf         As String
    Dim rowIndex    As Long
    Dim result      As Boolean
    Dim regexp      As Object
    Dim matches     As Object
    Dim match       As Object
    
    Dim stream      As Object
    Dim dic         As Object
    
    Dim patterns    As Variant
    Dim pattern     As Variant
    Dim matchVal    As String
    
    Const adReadLine As Long = -2

    result = False

    Set regexp = CreateObject("VBScript.RegExp")
    regexp.Global = False
    regexp.IgnoreCase = False
    
    Set dic = CreateObject("Scripting.Dictionary")

    patterns = Array(GREP_HEADER_PATTERN_PATTERN, _
                     GREP_HEADER_PATTERN_FILENAME, _
                     GREP_HEADER_PATTERN_SEARCH_PATH, _
                     GREP_HEADER_PATTERN_ALL_DIRECTORIES, _
                     GREP_HEADER_PATTERN_IGNORECASE, _
                     GREP_HEADER_PATTERN_USE_REGEXP, _
                     GREP_HEADER_PATTERN_TEXTONLY)

    grepHeader.SearchPath = ""
    grepHeader.FileName = ""
    grepHeader.Keyword = ""
    grepHeader.AllDirectories = False
    grepHeader.UseRegExp = False
    grepHeader.IgnoreCase = False
    grepHeader.TextOnly = False
    grepHeader.UseHighLight = False
    
    rowIndex = 0
        
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2
    stream.Charset = IIf(inData.Encoding = ENCODING_UTF8, "utf-8", "shift_jis")
    stream.Open
    stream.LoadFromFile inData.ResultPath
    stream.Position = 0

    Do Until stream.Eos
    
        buf = stream.ReadText(adReadLine)
    
        rowIndex = rowIndex + 1
    
        If rowIndex >= GREP_HEADER_MAX_ROWS Then
                    
            Exit Do
        End If
    
        For Each pattern In patterns
        
            regexp.pattern = pattern
    
            If regexp.Test(buf) Then
                
                Set matches = regexp.Execute(buf)
                Set match = matches(0)
                
                If match.SubMatches.Count > 0 Then
                    
                    matchVal = match.SubMatches(0)
                Else
                
                    matchVal = match
                End If
                
                Set matches = Nothing
                Set match = Nothing
                
                If Not dic.Exists(pattern) Then
                
                    dic.Add pattern, matchVal
                    
                    Exit For
                End If
            End If
        Next
    Loop
    
    stream.Close
    
    Set stream = Nothing
    Set regexp = Nothing
    Set matches = Nothing
    Set match = Nothing
    
    grepHeader.Keyword = IIf(dic.Exists(GREP_HEADER_PATTERN_PATTERN), _
                                    dic.Item(GREP_HEADER_PATTERN_PATTERN), "")
    grepHeader.FileName = IIf(dic.Exists(GREP_HEADER_PATTERN_FILENAME), _
                                    dic.Item(GREP_HEADER_PATTERN_FILENAME), "")
    grepHeader.SearchPath = IIf(dic.Exists(GREP_HEADER_PATTERN_SEARCH_PATH), _
                                    RemovePathSeparator(dic.Item(GREP_HEADER_PATTERN_SEARCH_PATH)), "")
    grepHeader.AllDirectories = dic.Exists(GREP_HEADER_PATTERN_ALL_DIRECTORIES)
    grepHeader.UseRegExp = dic.Exists(GREP_HEADER_PATTERN_USE_REGEXP)
    grepHeader.IgnoreCase = dic.Exists(GREP_HEADER_PATTERN_IGNORECASE)
    grepHeader.TextOnly = dic.Exists(GREP_HEADER_PATTERN_TEXTONLY)
    grepHeader.UseHighLight = inData.UseHighLight

    If Len(grepHeader.Keyword) > 0 And _
       Len(grepHeader.FileName) > 0 And _
       Len(grepHeader.SearchPath) > 0 Then
    
        result = True
    End If
    
    TryParseHeaderData = result
End Function

'-------------------------------------------------------------------------------
' 処理名 : SetResultHeaderData
'
' 概要   : Result シートのヘッダ部を初期化する。
'          ・検索条件用セルにドロップダウン（入力規則）を設定
'          ・grepヘッダ解析結果を表示用セルへ反映
'
'
' 引数   : wsResult   - 出力先の Result ワークシート
'          grepHeader - grep結果ファイルから解析したヘッダ情報
'-------------------------------------------------------------------------------
Private Sub SetResultHeaderData(ByVal wsResult As Worksheet, ByRef headerData As GrepHeaderData)
    Dim rngResult   As Range
    Dim dic         As Object
    Dim addr        As Variant
    Dim csv         As String

    Set dic = CreateObject("Scripting.Dictionary")

    dic.Add ADDR_RESULT_ALL_DIRECTORIES, CSV_ALL_DIRECTORIES
    dic.Add ADDR_RESULT_USE_REGEXP, CSV_USE_REGEXP
    dic.Add ADDR_RESULT_TEXTONLY, CSV_TEXTONLY
    dic.Add ADDR_RESULT_IGNORECASE, CSV_IGNORECASE
    dic.Add ADDR_RESULT_USE_HIGHLIGHT, CSV_USE_HIGHLIGHT
    
    For Each addr In dic.Keys
    
        Set rngResult = wsResult.Range(addr)
        csv = dic.Item(addr)
        
        SetValidateList rngResult, csv
        Set rngResult = Nothing
    Next
    
    Set dic = Nothing
  
    wsResult.Range(ADDR_RESULT_SEARCH_PATH).Value = headerData.SearchPath
    wsResult.Range(ADDR_RESULT_FILE).Value = headerData.FileName
    wsResult.Range(ADDR_RESULT_PATTERN).Value = headerData.Keyword
    wsResult.Range(ADDR_RESULT_USE_REGEXP).Value = GetUseRegExpDisplayName(headerData.UseRegExp)
    wsResult.Range(ADDR_RESULT_TEXTONLY).Value = GetTextOnlyDisplayName(headerData.TextOnly)
    wsResult.Range(ADDR_RESULT_ALL_DIRECTORIES).Value = GetAllDirectoriesDisplayName(headerData.AllDirectories)
    wsResult.Range(ADDR_RESULT_IGNORECASE).Value = GetIgnoreCaseDisplayName(headerData.IgnoreCase)
    wsResult.Range(ADDR_RESULT_USE_HIGHLIGHT).Value = GetHighLightsDisplayName(headerData.UseHighLight)
End Sub

'-------------------------------------------------------------------------------
' 関数名 : ReadGrepDetails
'
' 概要   : grep結果ファイルを解析し、詳細情報をワークシート出力用の
'          2次元配列として生成する。
'
'          ・grep 出力 1 行を正規表現で分解
'            (ファイルパス / 行・列 / 文字コード / ソース)
'          ・バイナリ行 / コメント行 / 文字化け行の判定
'          ・結果を列定義に沿って配列へ格納
'
'
' 引数   : inData     - 入力設定情報（ファイルパス、判定マーク等）
'          grepHeader - grep ヘッダ情報（検索ディレクトリ等）
'
' 戻り値 : ワークシートへ一括出力可能な 2 次元文字列配列
'-------------------------------------------------------------------------------
Private Function ReadGrepDetails(ByRef inData As InputData, _
                                 ByRef grepHeader As GrepHeaderData) As String()
    Dim sourceData      As Variant

    Dim buf             As String
    Dim rowIndex        As Long
    
    Dim regexp          As Object
    Dim matches         As Object
    Dim match           As Object
    
    Dim parent          As String
    Dim path            As String
    Dim aliasPath       As String
    Dim folder          As String
    Dim file            As String
    Dim ext             As String
    Dim pos             As String
    Dim enc             As String
    Dim result          As String
    Dim src             As String
    
    Dim commentPattern  As Object
    
    Dim dataArray()     As Variant
    Dim capacity        As Long
    Dim rowData(COLIDX_RESULT_ROWIDX To COLIDX_RESULT_SOURCE) As String
        
    Dim i, j As Long
    
    Const INITIAL_SIZE  As Long = 5000
    Const CHUNK_SIZE    As Long = 5000

    rowIndex = 0
    
    sourceData = ReadAllLines(inData.ResultPath, inData.Encoding)

    capacity = INITIAL_SIZE
    ReDim dataArray(1 To capacity)
    
    Set commentPattern = CreateCommentExtensionPattern()
    
    Set regexp = CreateObject("VBScript.RegExp")
    regexp.Global = False
    regexp.IgnoreCase = False
    regexp.pattern = "^(.+)\((\d+,\d+)\)\s*\[([\w\d-]+)\]:\s*(.+)"
    
    For i = 0 To UBound(sourceData)
        
        buf = sourceData(i)
        
        If Len(buf) = 0 Then
            
            GoTo For_Continue
        End If
        
        If Not regexp.Test(buf) Then
            
            GoTo For_Continue
        End If
        
        rowIndex = rowIndex + 1
        
        Set matches = regexp.Execute(buf)
        Set match = matches(0)
        
        path = match.SubMatches(0)
        pos = match.SubMatches(1)
        enc = match.SubMatches(2)
        src = match.SubMatches(3)
        
        If Left(src, 1) = "'" Then
            
            src = "'" & src 'から始まる文字列の場合、'が画面に表示されないので先頭に'を追加する。
        End If
        
        src = Replace(src, Chr(0), vbNullString)        ' NULLがあるとワークシートへの貼り付けで途中で中断されてしまう。
        
        If InStr(grepHeader.SearchPath, ";") = 0 Then
        
            parent = GetParentPath(path)
            folder = Replace(AddPathSeparator(parent), _
                             AddPathSeparator(grepHeader.SearchPath), "")
        Else
        
            folder = GetParentPath(path)
        End If
        
        file = GetFileName(path)
        ext = LCase(GetExtension(path))
        
        Set matches = Nothing
        Set match = Nothing
        
        result = ""
        
        If Len(inData.BinaryMark) > 0 And Len(result) = 0 Then
                    
            If IsBinary(src) Then
                
                result = inData.BinaryMark
            End If
        End If
        
        If Len(inData.CommentMark) > 0 And Len(result) = 0 Then
                    
            If IsComment(src, ext, commentPattern) Then
                
                result = inData.CommentMark
            End If
        End If
        
        If Len(inData.GarbledMark) > 0 And Len(result) = 0 Then
                    
            If IsGarbled(src) Then
                
                result = inData.GarbledMark
            End If
        End If
        
        If rowIndex > capacity Then
            
            capacity = capacity + CHUNK_SIZE
            ReDim Preserve dataArray(1 To capacity)
        End If
        
        rowData(COLIDX_RESULT_ROWIDX) = rowIndex
        rowData(COLIDX_RESULT_FOLDER) = folder
        rowData(COLIDX_RESULT_FILE) = file
        rowData(COLIDX_RESULT_EXTENSION) = ext
        rowData(COLIDX_RESULT_POSITION) = pos
        rowData(COLIDX_RESULT_ENCODING) = enc
        rowData(COLIDX_RESULT_RESULT) = result
        rowData(COLIDX_RESULT_SOURCE) = src
        
        dataArray(rowIndex) = rowData    '1次元の配列を1行分として格納
        
For_Continue:
    Next
    
    If rowIndex > 0 Then

        Dim outputArray() As String
        ReDim outputArray(1 To rowIndex, 1 To COLIDX_RESULT_SOURCE)

        For i = 1 To rowIndex

            For j = COLIDX_RESULT_ROWIDX To COLIDX_RESULT_SOURCE

                outputArray(i, j) = dataArray(i)(j)
            Next
        Next
    End If

    Erase dataArray

    ReadGrepDetails = outputArray
End Function

'-------------------------------------------------------------------------------
' 関数名 : ValidateMainInput
'
' 概要   : メイン画面で入力された内容を検証する
'          ・grep 結果ファイルパスの未入力チェック
'          ・文字コードの未入力チェック
'          ・環境変数展開後のファイル存在チェック
'          エラーがある場合はメッセージを表示し False を返す
'
' 引数   : wsMain - 入力値が設定されているワークシート
'          inData - 入力値
'
' 戻り値 : すべての入力が正しい場合 True、それ以外は False
'-------------------------------------------------------------------------------
Private Function ValidateMainInput(ByVal wsMain As Worksheet, _
                                   ByRef inData As InputData) As Boolean
    Dim result          As Boolean
    
    result = False
    
    inData.ResultPath = RTrim(wsMain.Range(ADDR_MAIN_RESULT_PATH).Value)
    inData.Encoding = RTrim(wsMain.Range(ADDR_MAIN_ENCODING).Value)
    inData.UseHighLight = GetHighLightsFlag(wsMain.Range(ADDR_MAIN_USE_HIGHLIGHT).Value)
    inData.CommentMark = RTrim(wsMain.Range(ADDR_MAIN_COMMENT_MARK).Value)
    inData.BinaryMark = RTrim(wsMain.Range(ADDR_MAIN_BINARY_MARK).Value)
    inData.GarbledMark = RTrim(wsMain.Range(ADDR_MAIN_GARBLED_MARK).Value)

    If Len(inData.ResultPath) = 0 Then
    
        MsgBox "grep結果ファイルを入力してください", vbExclamation, vbOKOnly
        
        GoTo Exit_ValidateMainInput
    End If
    
    If Len(inData.Encoding) = 0 Then
    
        MsgBox "文字コードを入力してください", vbExclamation, vbOKOnly
        
        GoTo Exit_ValidateMainInput
    End If
    
    ' 環境変数を含むパスを実パスに展開
    inData.ResultPath = ExpandPathFromEnvVar(inData.ResultPath)
    
    If Len(Dir(inData.ResultPath, vbNormal)) = 0 Then
    
        MsgBox "grep結果ファイルが見つかりません", vbExclamation, vbOKOnly
        
        GoTo Exit_ValidateMainInput
    End If
    
    result = True
    
Exit_ValidateMainInput:

    ValidateMainInput = result
End Function

'-------------------------------------------------------------------------------
' 関数名 : BrowseAndSetFilePath
'
' 概要   : ファイル選択ダイアログを表示し、選択したファイルパスを
'          メインシートの指定セルに環境変数表記で設定する
'          （マクロとして表示されないようにFunctionとしている）
'-------------------------------------------------------------------------------
Public Function BrowseAndSetFilePath()
    Dim filePath As String
    Dim ws As Worksheet
    
    Set ws = ThisWorkbook.Sheets(WSNM_MAIN)
    
    If ShowOpenFileDialog(filePath, caption:="ファイルを選択してください") Then
    
        ws.Range(ADDR_MAIN_RESULT_PATH).Value = ExpandPathToEnvVar(filePath)
    End If
    
    Set ws = Nothing
End Function

'-------------------------------------------------------------------------------
' 関数名 : ImportData
'
' 概要   : grep結果ファイルを読み込み、結果専用ブックを生成して
'              整形済みの Result シートへ出力する。
'
'          ・入力チェック
'          ・grepヘッダ解析
'          ・結果ブックの複製＆別プロセス起動
'          ・検索条件／ドロップダウン設定
'          ・grep詳細結果の読み込みと貼り付け
'          ・書式・オートフィルタ設定
'          ・別Excelプロセスの後始末
'          ・マクロとして表示されないようにFunctionとしている
'-------------------------------------------------------------------------------
Public Function ImportData()
    Dim inData          As InputData
    Dim headerData      As GrepHeaderData
    Dim sourceData()    As String
    
    Dim resultFileName  As String
    Dim resultFilePath  As String
    
    Dim wsMain          As Worksheet
    
    Dim xlApp           As Application
    Dim wbResult        As Workbook
    Dim wsResult        As Worksheet
    Dim wsMainResult    As Worksheet
    
    Dim pid             As Long
    
    Dim rowIndex        As Long
    
    Dim startTime       As Double
    Dim endTime         As Double
    Dim duration        As Double
    
    Set wsMain = ThisWorkbook.Sheets(WSNM_MAIN)
    
    ' 入力チェック
    If Not ValidateMainInput(wsMain, inData) Then
    
        GoTo Exit_ImportButton_Click
    End If
    
    ' ヘッダー読み込み
    If Not TryParseHeaderData(inData, headerData) Then
            
        MsgBox "grep結果ファイルが正しくありません。" & vbCrLf & _
               "または文字コードが正しくありません", vbExclamation, vbOKOnly
            
        GoTo Exit_ImportButton_Click
    End If
    
    startTime = Timer
    
    ' 結果ファイルの設定
    resultFilePath = GenerateResuFilePath(headerData.Keyword)
    resultFileName = GetFileName(resultFilePath)
    
    ' EXCELの設定
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    ' ファイルコピー
    ThisWorkbook.SaveCopyAs resultFilePath
    
    ' 別プロセスでエクセル起動
    Set xlApp = New Excel.Application
    xlApp.Visible = False
    xlApp.DisplayAlerts = False
    xlApp.EnableEvents = False
    
    ' hWndからプロセスIDを取得
    pid = GetProcessIdFromHwnd(xlApp.hWnd)
    
    ' 結果ファイルの設定
    Set wbResult = xlApp.Workbooks.Open(resultFilePath)
    Set wsResult = wbResult.Sheets(WSNM_RESULT)
    Set wsMainResult = wbResult.Sheets(WSNM_MAIN)
    
    wsResult.Visible = xlSheetVisible
    wsResult.Select
    wsMainResult.Visible = xlSheetVeryHidden

    ' ヘッダーを設定する
    SetResultHeaderData wsResult, headerData
    
    ' 全件読み込み
    sourceData = ReadGrepDetails(inData, headerData)
    
    On Error Resume Next
    rowIndex = UBound(sourceData)
    On Error GoTo 0
    
    If rowIndex > 0 Then
            
        ' Resultシートに貼り付け
        wsResult.Range(wsResult.Cells(ROW_OFFSET_RESULT + 1, COLIDX_RESULT_ROWIDX), _
                       wsResult.Cells(ROW_OFFSET_RESULT + rowIndex, COLIDX_RESULT_SOURCE)).Value = sourceData
    End If
    
    ' 背景色を設定
    With wsResult.Range(wsResult.Cells(ROW_OFFSET_RESULT + 1, COLIDX_RESULT_ROWIDX), _
                        wsResult.Cells(ROW_OFFSET_RESULT + rowIndex, COLIDX_RESULT_SOURCE_FILLER4))
    
        .Interior.Color = vbWhite
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .Borders.Color = vbBlack
    End With

    With wsResult.Range(wsResult.Cells(ROW_OFFSET_RESULT + 1, COLIDX_RESULT_FOLDER), _
                        wsResult.Cells(ROW_OFFSET_RESULT + rowIndex, COLIDX_RESULT_FOLDER_FILLER1))
            
        .Borders(xlInsideVertical).LineStyle = xlNone
    End With

    With wsResult.Range(wsResult.Cells(ROW_OFFSET_RESULT + 1, COLIDX_RESULT_SOURCE), _
                        wsResult.Cells(ROW_OFFSET_RESULT + rowIndex, COLIDX_RESULT_SOURCE_FILLER4))
            
        .Borders(xlInsideVertical).LineStyle = xlNone
    End With
    
    xlApp.EnableEvents = True

    ' フィルタを設定
    wsResult.Rows(AUTO_FILTER_RESULT).EntireRow.AutoFilter

    wbResult.Close True
    
    Set wsResult = Nothing
    Set wbResult = Nothing
    Set wsMainResult = Nothing
    
    xlApp.DisplayAlerts = True
    xlApp.Quit
    
    Set xlApp = Nothing

    WaitMilliseconds 200
    
    If IsProcessRunning(pid) Then
            
        KillProcess pid
    End If
    
    Application.ScreenUpdating = True
    
    MsgBox "grep結果の取り込みが完了しました。" & vbCrLf & resultFileName, vbInformation + vbOKOnly
        
Exit_ImportButton_Click:

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
End Function

'-------------------------------------------------------------------------------
' 関数名 : ConvertFilePathToEnvVar
'
' 概要   : 指定セルのパスを環境変数表記に変換して更新する
'
' 引数   : rng - 変換対象のセル
'-------------------------------------------------------------------------------
Public Sub ConvertFilePathToEnvVar(ByVal rng As Range)
    Dim orgPath As Variant
    Dim newPath As Variant
    
    orgPath = rng.Value
    
    ' 配列が渡された場合は処理しない
    If IsArray(orgPath) Then Exit Sub
    
    newPath = ExpandPathToEnvVar(orgPath)
    
    If orgPath <> newPath Then
        rng.Value = newPath
    End If
End Sub
