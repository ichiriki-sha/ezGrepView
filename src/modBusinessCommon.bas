Attribute VB_Name = "modBusinessCommon"
Option Explicit

Private Const PATH_ALIAS_PREFIX As String = "Dir"

'-------------------------------------------------------------------------------
' 関数名 : SetValidateList
'
' 概要   : 指定したセル範囲にリスト形式の入力規則を設定する
'          既存の入力規則がある場合は削除してから再設定する
'
' 引数   : rng     - 入力規則を設定するセル範囲
'        : csvData - カンマ区切りのリストデータ
'                    （例: "UTF-8,Shift_JIS"）
'-------------------------------------------------------------------------------
Public Sub SetValidateList(ByVal rng As Range, ByVal csvData As String)

    Dim v As Validation   ' 入力規則オブジェクト

    ' 対象セルの入力規則を取得
    Set v = rng.Validation

    With v
        ' 既存の入力規則を削除
        .Delete

        ' リスト形式の入力規則を追加
        .Add _
            Type:=xlValidateList, _
            AlertStyle:=xlValidAlertStop, _
            Operator:=xlBetween, _
            Formula1:=csvData

        ' 空白入力を許可
        .IgnoreBlank = True

        ' セル内にドロップダウンを表示
        .InCellDropdown = True

        ' IME の制御は行わない
        .IMEMode = xlIMEModeNoControl

        ' 入力時のメッセージを表示
        .ShowInput = True

        ' 不正入力時のエラーメッセージを表示
        .ShowError = True
    End With

    ' オブジェクト解放
    Set v = Nothing
End Sub

'-------------------------------------------------------------------------------
' 関数名 : GetAllDirectoriesFlag
' 概要   : 「すべてのサブディレクトリを検索する」設定の
'          表示用文字列から Boolean 値を取得する。
'
' 引数   : displayName - 画面上の表示文字列
'                         （例: DISP_ALL_DIRECTORIES_YES / NO）
'
' 戻り値 : True  - 全ディレクトリを検索する
'          False - 指定ディレクトリのみ検索する
'-------------------------------------------------------------------------------
Public Function GetAllDirectoriesFlag(ByVal displayName As String) As Boolean

    GetAllDirectoriesFlag = (displayName = DISP_ALL_DIRECTORIES_YES)
End Function

'-------------------------------------------------------------------------------
' 関数名 : GetUseRegExpFlag
' 概要   : 「正規表現を使用する」設定の
'          表示用文字列から Boolean 値を取得する。
'
' 引数   : displayName - 画面上の表示文字列
'                         （例: DISP_USE_REGEXP_YES / NO）
'
' 戻り値 : True  - 正規表現を使用する
'          False - 通常文字列検索を行う
'-------------------------------------------------------------------------------
Public Function GetUseRegExpFlag(ByVal displayName As String) As Boolean

    GetUseRegExpFlag = (displayName = DISP_USE_REGEXP_YES)
End Function

'-------------------------------------------------------------------------------
' 関数名 : GetTextOnlyFlag
' 概要   : 「テキストのみ検索」設定の
'          表示用文字列から Boolean 値を取得する。
'
' 引数   : displayName - 画面上の表示文字列
'                         （例: DISP_TEXTONLY_YES / NO）
'
' 戻り値 : True  - テキストのみを対象に検索する
'          False - バイナリ等も含めて検索する
'-------------------------------------------------------------------------------
Public Function GetTextOnlyFlag(ByVal displayName As String) As Boolean

    GetTextOnlyFlag = (displayName = DISP_TEXTONLY_YES)
End Function

'-------------------------------------------------------------------------------
' 関数名 : GetIgnoreCaseFlag
' 概要   : 「大文字・小文字を区別しない」設定の
'          表示用文字列から Boolean 値を取得する。
'
' 引数   : displayName - 画面上の表示文字列
'                         （例: DISP_IGNORECASE_YES / NO）
'
' 戻り値 : True  - 大文字・小文字を区別しない
'          False - 区別する
'-------------------------------------------------------------------------------
Public Function GetIgnoreCaseFlag(ByVal displayName As String) As Boolean

    GetIgnoreCaseFlag = (displayName = DISP_IGNORECASE_YES)
End Function

'-------------------------------------------------------------------------------
' 関数名 : GetHighLightsFlag
' 概要   : 「ハイライト表示を行う」設定の
'          表示用文字列から Boolean 値を取得する。
'
' 引数   : displayName - 画面上の表示文字列
'                         （例: DISP_USE_HIGHLIGHT_YES / NO）
'
' 戻り値 : True  - 検索結果をハイライト表示する
'          False - ハイライト表示を行わない
'-------------------------------------------------------------------------------
Public Function GetHighLightsFlag(ByVal displayName As String) As Boolean

    GetHighLightsFlag = (displayName = DISP_USE_HIGHLIGHT_YES)
End Function

'-------------------------------------------------------------------------------
' 関数名 : GetAllDirectoriesDisplayName
'
' 概要   : 「サブフォルダーも検索」フラグの値に応じて、
'          表示用の文字列（Yes / No など）を返す
'
' 引数   : flag - サブフォルダーも検索する場合 True
'
' 戻り値 : 表示用文字列
'          True  : DISP_ALL_DIRECTORIES_YES
'          False : DISP_ALL_DIRECTORIES_NO
'-------------------------------------------------------------------------------
Public Function GetAllDirectoriesDisplayName(ByVal flag As Boolean) As String

    ' フラグに応じて表示名を切り替える
    GetAllDirectoriesDisplayName = IIf(flag, DISP_ALL_DIRECTORIES_YES, DISP_ALL_DIRECTORIES_NO)
End Function

'-------------------------------------------------------------------------------
' 関数名 : GetUseRegExpDisplayName
'
' 概要   : 「正規表現を使用する」フラグの値に応じて、
'          表示用の文字列（Yes / No など）を返す
'
' 引数   : flag - 正規表現を使用する場合 True
'
' 戻り値 : 表示用文字列
'          True  : DISP_USE_REGEXP_YES
'          False : DISP_USE_REGEXP_NO
'-------------------------------------------------------------------------------
Public Function GetUseRegExpDisplayName(ByVal flag As Boolean) As String

    ' フラグに応じて表示名を切り替える
    GetUseRegExpDisplayName = IIf(flag, DISP_USE_REGEXP_YES, DISP_USE_REGEXP_NO)
End Function

'-------------------------------------------------------------------------------
' 関数名 : GetTextOnlyDisplayName
'
' 概要   : 「テキストのみ表示」フラグの値に応じて、
'          表示用の文字列（Yes / No など）を返す。
'
' 引数   : flag - テキストのみ表示する場合 True
'
' 戻り値 : 表示用文字列
'          True  : DISP_TEXTONLY_YES
'          False : DISP_TEXTONLY_NO
'-------------------------------------------------------------------------------
Public Function GetTextOnlyDisplayName(ByVal flag As Boolean) As String

    ' フラグに応じて表示名を切り替える
    GetTextOnlyDisplayName = IIf(flag, DISP_TEXTONLY_YES, DISP_TEXTONLY_NO)
End Function

'-------------------------------------------------------------------------------
' 関数名 : GetIgnoreCaseDisplayName
'
' 概要   : 「大文字小文字を区別しない」フラグの値に応じて、
'          表示用の文字列（Yes / No など）を返す
'
' 引数   : flag - 大文字小文字を区別しない場合 True
'
' 戻り値 : 表示用文字列
'          True  : DISP_IGNORECASE_YES
'          False : DISP_IGNORECASE_NO
'-------------------------------------------------------------------------------
Public Function GetIgnoreCaseDisplayName(ByVal flag As Boolean) As String

    ' フラグに応じて表示名を切り替える
    GetIgnoreCaseDisplayName = IIf(flag, DISP_IGNORECASE_YES, DISP_IGNORECASE_NO)
End Function

'-------------------------------------------------------------------------------
' 関数名 : GetHighLightsDisplayName
'
' 概要   : 「ハイライト表示」フラグの値に応じて、
'          表示用の文字列（Yes / No など）を返す
'
' 引数   : flag - ハイライト表示を有効にする場合 True
'
' 戻り値 : 表示用文字列
'          True  : DISP_USE_HIGHLIGHT_YES
'          False : DISP_USE_HIGHLIGHT_NO
'-------------------------------------------------------------------------------
Public Function GetHighLightsDisplayName(ByVal flag As Boolean) As String

    ' フラグに応じて表示名を切り替える
    GetHighLightsDisplayName = IIf(flag, DISP_USE_HIGHLIGHT_YES, DISP_USE_HIGHLIGHT_NO)
End Function

'-------------------------------------------------------------------------------
' 関数名 : ReadCommentCore
' 概要   : コメント定義シートからコメント情報を読み込み、
'          CommentData 構造体の配列として返す共通処理。
'
'          シートの列構成は以下を想定：
'            Col1 : 拡張子（例: c, cpp, vb など）
'            Col2 : コメント種別表示文字列（例: ライン / ブロック）
'            Col3 : コメント開始文字列
'            Col4 : コメント終了文字列（ラインコメントの場合は空）
'
'          ext が空文字の場合は全件を対象とし、
'          ext が指定されている場合は拡張子一致の行のみを読み込む。
'
' 引数   : wsComment - コメント定義が記載されたワークシート
'          ext        - 対象とする拡張子（"" の場合は全件）
'          comData    - 読み込んだコメント定義を格納する配列（出力）
'
' 戻り値 : Boolean
'          True  - 1 件以上のコメント定義を読み込んだ場合
'          False - 対象となるコメント定義が存在しなかった場合
'
' 備考   : ヘッダ行は 1 行目に存在すると想定し、
'          2 行目以降を順に読み込み、拡張子列が空になるまで処理を続行する。
'-------------------------------------------------------------------------------
Private Function ReadCommentCore(ByVal wsComment As Worksheet, _
                                 ByVal ext As String, _
                                 ByRef comData() As CommentData) As Boolean
    Dim rowIndex    As Long
    Dim i           As Long

    Dim col1 As String, col2 As String, col3 As String, col4 As String
    Dim ret As Boolean

    ret = False

    ext = LCase(RTrim(ext))
    i = 0
    rowIndex = 1    ' ヘッダ行想定

    Do
        rowIndex = rowIndex + 1

        col1 = LCase(Trim(wsComment.Cells(rowIndex, 1).Value))
        If Len(col1) = 0 Then Exit Do

        col2 = LCase(Trim(wsComment.Cells(rowIndex, 2).Value))
        col3 = wsComment.Cells(rowIndex, 3).Value
        col4 = wsComment.Cells(rowIndex, 4).Value

        If Len(ext) = 0 Or col1 = ext Then

            ReDim Preserve comData(i)

            comData(i).Extension = col1

            If col2 = DISP_COMMENT_TYPE_LINE Then
            
                comData(i).CommentType = CommentType.Line
            Else
            
                comData(i).CommentType = CommentType.Block
            End If

            comData(i).CommentStart = col3
            comData(i).CommentEnd = col4

            ret = True
            
            i = i + 1
        End If
    Loop

    ReadCommentCore = ret
End Function

'-------------------------------------------------------------------------------
' 関数名 : ReadAllCommentFromSheet
' 概要   : コメント定義シートから、全拡張子を対象として
'          コメント定義を読み込み、CommentData 構造体配列として返す。
'
'          内部的には共通処理である ReadCommentCore を呼び出し、
'          拡張子条件を指定しない形で全行を読み込む。
'
' 引数   : wsComment - コメント定義が記載されたワークシート
'          comData    - 読み込んだコメント定義を格納する配列（出力）
'
' 戻り値 : Boolean
'          True  - 1 件以上のコメント定義を読み込んだ場合
'          False - コメント定義が 1 件も存在しなかった場合
'-------------------------------------------------------------------------------
Private Function ReadAllCommentFromSheet(ByVal wsComment As Worksheet, _
                                         ByRef comData() As CommentData) As Boolean

    ReadAllCommentFromSheet = ReadCommentCore(wsComment, "", comData)
End Function

'-------------------------------------------------------------------------------
' 関数名 : ReadCommentFromSheet
' 概要   : コメント定義シートから、指定された拡張子に一致する
'          コメント定義のみを読み込み、CommentData 構造体配列として返す。
'
'          内部的には共通処理である ReadCommentCore を呼び出し、
'          ext が指定されている場合は拡張子一致行のみを対象とする。
'
' 引数   : wsComment - コメント定義が記載されたワークシート
'          ext        - 対象とする拡張子（例: "c", "cpp", "vb"）
'          comData    - 読み込んだコメント定義を格納する配列（出力）
'
' 戻り値 : Boolean
'          True  - 1 件以上のコメント定義を読み込んだ場合
'          False - 指定拡張子に該当するコメント定義が存在しなかった場合
'-------------------------------------------------------------------------------
Private Function ReadCommentFromSheet(ByVal wsComment As Worksheet, _
                                      ByVal ext As String, _
                                      ByRef comData() As CommentData) As Boolean

    ReadCommentFromSheet = ReadCommentCore(wsComment, ext, comData)
End Function

'-------------------------------------------------------------------------------
' 関数名 : BuildCommentPatternDictionary
' 概要   : CommentData 配列からコメント検出用の正規表現パターンを生成し、
'          指定されたキー（拡張子またはコメント種別）ごとに
'          Dictionary としてまとめて返す。
'
'          各コメント定義は正規表現に変換され、
'          同一キーに属する複数定義は OR（|）で連結される。
'
' 引数   : comData     - コメント定義を格納した CommentData 配列
'          keySelector - Dictionary のキー種別
'                        0 : Extension（拡張子）をキーにする
'                        1 : CommentType（Line / Block）をキーにする
'
' 戻り値 : Object（Scripting.Dictionary）
'          Key   : 拡張子（String）または CommentType（Enum）
'          Value : コメント検出用の正規表現パターン文字列
'
' 用途例 :
'   - 拡張子ごとのコメントパターン管理
'   - コメント種別（Line / Block）ごとの一括正規表現生成
'-------------------------------------------------------------------------------
Private Function BuildCommentPatternDictionary(ByRef comData() As CommentData, _
                                               ByVal keySelector As Long) As Object
    ' keySelector:
    '   0 = Extension
    '   1 = CommentType

    Dim dic As Object
    Dim i As Long
    Dim pattern As String
    Dim key As Variant

    Set dic = CreateObject("Scripting.Dictionary")

    For i = LBound(comData) To UBound(comData)

        Select Case comData(i).CommentType
            Case CommentType.Line
            
                pattern = EscapePattern(comData(i).CommentStart) & ".*$"

            Case CommentType.Block
                
                pattern = EscapePattern(comData(i).CommentStart) & ".*?" & _
                          EscapePattern(comData(i).CommentEnd)
        End Select

        If keySelector = 0 Then
            
            key = comData(i).Extension
        Else
            
            key = comData(i).CommentType
        End If

        If Not dic.Exists(key) Then
            
            dic.Add key, pattern
        Else
            
            dic(key) = dic(key) & "|" & pattern
        End If
    Next

    Set BuildCommentPatternDictionary = dic
End Function

'-------------------------------------------------------------------------------
' 関数名 : CreateCommentExtensionPattern
' 概要   : コメント定義シートの内容をもとに、
'          拡張子ごとのコメント検出用正規表現パターンを生成する
'
'          ・拡張子をキー
'          ・コメント検出用の正規表現を値
'          とする Dictionary を返す
'
'          同一拡張子に複数のコメント定義がある場合は、
'          OR（|）で連結した正規表現を生成する
'
' 戻り値 : Scripting.Dictionary
'          Key   : 拡張子（例: "cs", "vb", "cpp"）
'          Value : コメント検出用の正規表現パターン
'-------------------------------------------------------------------------------
Public Function CreateCommentExtensionPattern() As Object
    Dim wsComment As Worksheet
    Dim comData() As CommentData
    
    Set wsComment = ThisWorkbook.Sheets(WSNM_COMMENT)

    If Not ReadAllCommentFromSheet(wsComment, comData) Then
    
        Set CreateCommentExtensionPattern = CreateObject("Scripting.Dictionary")
        
        Exit Function
    End If

    Set CreateCommentExtensionPattern = BuildCommentPatternDictionary(comData, 0)

    Set wsComment = Nothing
End Function

'-------------------------------------------------------------------------------
' 関数名 : CreateCommentTypePattern
' 概要   : コメント定義シートから指定拡張子のコメント定義を読み込み、
'          コメント種別（Line / Block）ごとに使用可能な
'          正規表現パターンを生成して Dictionary として返す。
'
'          - Line コメント  : 開始文字列 + 行末まで（^ ～ $）
'          - Block コメント : 開始文字列 ～ 終了文字列（最短一致）
'
'          同一コメント種別が複数定義されている場合は、
'          正規表現の OR（|）で連結したパターンを生成する。
'
' 引数   : ext - 対象ファイルの拡張子（例: "c", "cpp", "vb"）
'
' 戻り値 : Object（Scripting.Dictionary）
'          Key   : CommentType（CommentType.Line / CommentType.Block）
'          Value : 対応する正規表現パターン文字列
'
'          ※ コメント定義が存在しない場合は空の Dictionary を返す
'-------------------------------------------------------------------------------
Public Function CreateCommentTypePattern(ByVal ext As String) As Object
    Dim wsComment As Worksheet
    Dim comData() As CommentData

    Set wsComment = ThisWorkbook.Sheets(WSNM_COMMENT)

    If Not ReadCommentFromSheet(wsComment, ext, comData) Then
        
        Set CreateCommentTypePattern = CreateObject("Scripting.Dictionary")
        
        Exit Function
    End If

    Set CreateCommentTypePattern = BuildCommentPatternDictionary(comData, 1)

    Set wsComment = Nothing
End Function
