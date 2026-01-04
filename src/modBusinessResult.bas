Attribute VB_Name = "modBusinessResult"
Option Explicit

'-------------------------------------------------------------------------------
' 関数名 : GetResultHeaderData
'
' 概要   : Result シートのヘッダ部に入力されている検索条件を読み取り、
'          GrepHeaderData 構造体へ変換して返す。
'          表示用文字列（Yes / No 等）は内部用の Boolean 値に変換する。
'
'
' 引数   : wsResult - Result ワークシート
'
' 戻り値 : Result シートから取得した GrepHeaderData
'-------------------------------------------------------------------------------
Private Function GetResultHeaderData(ByVal wsResult As Worksheet) As GrepHeaderData
    Dim headerData As GrepHeaderData
    
    headerData.SearchPath = wsResult.Range(ADDR_RESULT_SEARCH_PATH).Value
    headerData.FileName = wsResult.Range(ADDR_RESULT_FILE)
    headerData.Keyword = wsResult.Range(ADDR_RESULT_PATTERN)
    
    headerData.UseRegExp = GetUseRegExpFlag(wsResult.Range(ADDR_RESULT_USE_REGEXP))
    headerData.TextOnly = GetTextOnlyFlag(wsResult.Range(ADDR_RESULT_TEXTONLY))
    headerData.AllDirectories = GetAllDirectoriesFlag(wsResult.Range(ADDR_RESULT_ALL_DIRECTORIES))
    headerData.IgnoreCase = GetIgnoreCaseFlag(wsResult.Range(ADDR_RESULT_IGNORECASE))
    headerData.UseHighLight = GetHighLightsFlag(wsResult.Range(ADDR_RESULT_USE_HIGHLIGHT))

    GetResultHeaderData = headerData
End Function

'-------------------------------------------------------------------------------
' 関数名 : QuickSortHighLights
'
' 概要   : 配列 arr を「開始位置 (arr()の0番目要素)」を基準に昇順でソートする
'                  （QuickSort アルゴリズムによる再帰ソート）
'
' 引数   : arr()  - ソート対象の配列（各要素は配列で、0番目がキー）
'          first  - ソート範囲の開始インデックス
'          last   - ソート範囲の終了インデックス
'-------------------------------------------------------------------------------
Private Sub QuickSortHighLightsCore(ByRef arr As Variant, _
                                    ByVal first As Long, _
                                    ByVal last As Long)

    Dim low As Long, high As Long
    Dim pivot As Variant
    Dim temp As Variant

    low = first
    high = last
    pivot = arr((first + last) \ 2)

    Do While low <= high

        Do While arr(low)(0) < pivot(0)
            
            low = low + 1
        Loop

        Do While arr(high)(0) > pivot(0)
            
            high = high - 1
        Loop

        If low <= high Then
            
            temp = arr(low)
            arr(low) = arr(high)
            arr(high) = temp

            low = low + 1
            high = high - 1
        End If
    Loop

    If first < high Then QuickSortHighLightsCore arr, first, high
    If low < last Then QuickSortHighLightsCore arr, low, last
End Sub

'-------------------------------------------------------------------------------
' 関数名 : QuickSortHighLights
'
' 概要   : ハイライト情報を保持した Collection を、
'          開始位置（配列要素(0)）を基準に昇順でソートする。
'
'          Collection は直接 QuickSort できないため、
'          一度 Variant 配列へ変換してソート後、
'          並び替えた結果を Collection に戻す。
'
' 引数   : src - ハイライト情報を格納した Collection
'                各要素は Array(StartPos, Length, Color, Bold)
'-------------------------------------------------------------------------------
Private Sub QuickSortHighLights(ByRef src As Collection)

    Dim arr() As Variant
    Dim i As Long

    If src.Count <= 1 Then Exit Sub

    ReDim arr(1 To src.Count)
    
    For i = 1 To src.Count
        
        arr(i) = src(i)
    Next

    Call QuickSortHighLightsCore(arr, LBound(arr), UBound(arr))

    ' Collection をクリア（後ろから削除）
    For i = src.Count To 1 Step -1
        
        src.remove i
    Next
    
    ' 並び替え結果を Collection に戻す
    For i = LBound(arr) To UBound(arr)
        
        src.Add arr(i)
    Next
End Sub

'-------------------------------------------------------------------------------
' 関数名 : MergeHighlights
'
' 概要   : キーワードハイライトとコメントハイライトを結合し、
'          1つの Collection として返す。
'
'          本関数では重なり判定や並び替えは行わず、
'          2つの Collection の要素をそのまま順に結合する。
'
' 引数   : keywordHL - キーワード一致部分のハイライト情報 Collection
'                      (Array(Start, Length, Color, Bold))
'          commentHL - コメント部分のハイライト情報 Collection
'                      (Array(Start, Length, Color, Bold))
'
' 戻り値 : 結合後のハイライト情報 Collection
'-------------------------------------------------------------------------------
Private Function MergeHighlights(ByVal keywordHL As Collection, _
                                 ByVal commentHL As Collection) As Collection

    Dim result As New Collection
    Dim v As Variant

    ' キーワードハイライトを追加
    For Each v In keywordHL
        
        result.Add v
    Next

    ' コメントハイライトを追加
    For Each v In commentHL
        
        result.Add v
    Next

    Set MergeHighlights = result
End Function

'-------------------------------------------------------------------------------
' 関数名 : RemoveKeywordByComment
'
' 概要   : コメント領域と重複するキーワードハイライトを除外する。
'
'           ・行コメントの場合：
'             コメント開始位置以降にあるキーワードをすべて除外
'           ・ブロックコメントの場合：
'             コメント範囲内に完全に含まれるキーワードを除外
'
'             コメント側のハイライトを優先し、
'             キーワードハイライトの Collection を再構築する。
'
' 引数   : keywordHL - キーワードハイライト情報（更新対象）
'                      Array(Start, Length, Color, Bold, …)
'          commentHL - コメントハイライト情報
'                      Array(Start, Length, Color, Bold, CommentType)
'
' 戻り値 : なし
'          keywordHL はコメント除外後の内容に置き換えられる
'-------------------------------------------------------------------------------
Private Sub RemoveKeywordByComment(ByRef keywordHL As Collection, _
                                   ByVal commentHL As Collection)

    Dim result As New Collection
    Dim k As Variant, c As Variant
    Dim ks As Long, ke As Long
    Dim cs As Long, ce As Long
    Dim remove As Boolean

    ' キーワードハイライトを1件ずつチェック
    For Each k In keywordHL

        ks = k(0)
        ke = ks + k(1) - 1
        remove = False

        ' コメントハイライトとの重なり判定
        For Each c In commentHL
            
            cs = c(0)
            ce = cs + c(1) - 1

            Select Case c(4) ' CommentType
                Case CommentType.Line
                    ' 行コメント：
                    ' コメント開始以降にあるキーワードはすべて除外
                    
                    If ks >= cs Then remove = True
                Case CommentType.Block
                    ' ブロックコメント：
                    ' コメント範囲内に完全に含まれるキーワードのみ除外
                    
                    If ks >= cs And ke <= ce Then remove = True
            End Select

            If remove Then Exit For
        Next

        ' コメントと重複していなければ結果に追加
        If Not remove Then result.Add k
    Next

    ' キーワードハイライトを置き換え
    Set keywordHL = result
End Sub

'-------------------------------------------------------------------------------
' 関数名 : CreateCommentHighlights
'
' 概要   : 対象ソース文字列からコメント部分を検出し、
'          コメント用ハイライト情報の Collection を生成して返す。
'
'          ・コメント定義は Comment シートから取得する
'          ・行コメント / ブロックコメントの両方に対応
'          ・コメント部分はキーワードより優先表示される前提
'
' 引数   : filePath - 対象ファイルのフルパス（拡張子判定用）
'          src      - 1行分のソースコード文字列
'
' 戻り値 : Collection
'          各要素は以下の配列形式
'          Array(
'              開始位置 (1-based),
'              文字数,
'              色 (RGB),
'              Bold フラグ,
'              CommentType
'          )
'-------------------------------------------------------------------------------
Private Function CreateCommentHighlights(ByVal filePath As String, _
                                         ByVal src As String) As Collection
    Dim result As New Collection
    Dim dic As Object
    Dim comType As Variant
    
    Dim regexp As Object, m As Object

    ' コメント定義が存在しない場合でも処理を止めない
    On Error Resume Next

    ' 正規表現オブジェクト生成（複数マッチ対応）
    Set regexp = CreateObject("VBScript.RegExp")
    regexp.Global = True

    Set dic = CreateCommentTypePattern(GetExtension(filePath))
    
    For Each comType In dic.Keys
    
        regexp.pattern = dic.Item(comType)
    
        ' ソース内のコメントをすべて検出
        For Each m In regexp.Execute(src)

            ' 開始位置（1-based）
            ' コメント長
            ' コメント色（緑）
            ' 太字なし
            ' コメント種別
            result.Add Array(m.FirstIndex + 1, m.length, RGB(0, 128, 0), False, comType)
        Next
    Next
    
    ' 結果を返却
    Set CreateCommentHighlights = result
End Function

'-------------------------------------------------------------------------------
' 関数名 : MergeContinuousHighlights
'
' 概要   : ハイライト情報の Collection を開始位置順に並び替え、
'          「連続しているハイライト」を1つにまとめた
'          新しい Collection を生成して返す。
'
'          連続判定条件：
'          ・次の開始位置 = 現在の開始位置 + 現在の長さ
'          ・色が同一
'          ・Bold フラグが同一
'
' 引数   : src - ハイライト情報の Collection
'                 各要素は以下形式の配列
'                 Array(開始位置, 長さ, 色, Bold, ...)
'
' 戻り値 : Collection
'          連続ハイライトをマージ済みの新しい Collection
'-------------------------------------------------------------------------------
Private Function MergeContinuousHighlights(ByVal src As Collection) As Collection
    Dim result As New Collection
    Dim cur As Variant, nxt As Variant
    Dim i As Long

    ' 空の場合は空コレクションを返す
    If src.Count = 0 Then
        
        Set MergeContinuousHighlights = result
        
        Exit Function
    End If

    ' 開始位置で昇順ソート
    QuickSortHighLights src

    ' 最初の要素を基準に設定
    cur = src(1)

    ' 2件目以降を順に比較
    For i = 2 To src.Count
        
        nxt = src(i)

       ' 開始位置が連続しており、色・Bold が同一ならマージ
        If nxt(0) = cur(0) + cur(1) And _
           nxt(2) = cur(2) And _
           nxt(3) = cur(3) Then

            ' 長さを拡張
            cur(1) = cur(1) + nxt(1)
        Else
            ' 連続していなければ確定して結果に追加

            result.Add cur
            cur = nxt
        End If
    Next

    ' 最後の要素を追加
    result.Add cur
    
    ' マージ済みハイライトを返却
    Set MergeContinuousHighlights = result
End Function

'-------------------------------------------------------------------------------
' 関数名 : CreateKeywordHighlights
'
' 概要   : 指定されたソース文字列から検索キーワードに一致する箇所を抽出し、
'          ハイライト情報のコレクションとして返す。
'
'          ・正規表現／通常検索の切り替えに対応
'          ・大文字／小文字の区別設定に対応
'
' 戻り値 : Collection
'          各要素は以下の形式の配列
'            [0] 開始位置（1-based）
'            [1] 文字列長
'            [2] 表示色（vbRed）
'            [3] 強調表示フラグ（True）
'
' 引数   : src        - 検索対象となる元文字列
'          headerData - 検索条件を保持する GrepHeaderData 構造体
'-------------------------------------------------------------------------------
Private Function CreateKeywordHighlights(ByVal src As String, _
                                         ByRef headerData As GrepHeaderData) As Collection

    Dim result As New Collection
    Dim regexp As Object, m As Object

    Set regexp = CreateObject("VBScript.RegExp")
    regexp.Global = True
    regexp.IgnoreCase = headerData.IgnoreCase
    regexp.pattern = IIf(headerData.UseRegExp, _
                         headerData.Keyword, _
                         EscapePattern(headerData.Keyword))

    For Each m In regexp.Execute(src)
        
        result.Add Array(m.FirstIndex + 1, m.length, vbRed, True)
    Next

    Set CreateKeywordHighlights = result
End Function

'-------------------------------------------------------------------------------
' 関数名 : CreateHighLights
'
' 概要   : ソース文字列全体を解析し、キーワードおよびコメントの
'          ハイライト情報を統合したコレクションを生成して返す。
'
'          処理の流れ：
'            1. キーワード検索結果からハイライト情報を作成
'            2. 隣接・連続するキーワードハイライトを結合
'            3. コメント定義に基づきコメント用ハイライトを作成
'            4. コメント領域と重複するキーワードハイライトを除外
'            5. キーワード／コメント両方のハイライトを統合
'
' 戻り値 : Collection
'          ハイライト情報の集合
'          各要素は以下の形式の配列を想定
'            [0] 開始位置（1-based）
'            [1] 文字列長
'            [2] 表示色
'            [3] 強調表示フラグ
'
' 引数   : filePath   - 対象ファイルのパス（拡張子判定等に使用）
'          src        - 検索・解析対象となる元文字列
'          headerData - 検索条件を保持する GrepHeaderData 構造体
'-------------------------------------------------------------------------------
Private Function CreateHighLights(ByVal filePath As String, _
                                  ByVal src As String, _
                                  ByRef headerData As GrepHeaderData) As Collection

    Dim keywordHL As Collection    ' キーワード用ハイライト情報
    Dim commentHL As Collection    ' コメント用ハイライト情報

    ' キーワードに一致する箇所のハイライト情報を作成
    Set keywordHL = CreateKeywordHighlights(src, headerData)

    ' 隣接・連続するキーワードハイライトを結合
    Set keywordHL = MergeContinuousHighlights(keywordHL)

    ' コメント定義に基づきコメント部分のハイライト情報を作成
    Set commentHL = CreateCommentHighlights(filePath, src)

    ' コメント領域に含まれるキーワードハイライトを削除
    RemoveKeywordByComment keywordHL, commentHL

    ' キーワード／コメント両方のハイライト情報を統合して返却
    Set CreateHighLights = MergeHighlights(keywordHL, commentHL)
End Function

'-------------------------------------------------------------------------------
' 関数名 : UpdateSourceCell
'
' 概要   : Result シート上で選択された行に基づき、
'          ソース文字列・ファイルパスを表示エリアへ反映し、
'          検索条件に応じたハイライト処理を行う。
'
'          主な処理内容：
'            1. ヘッダ行・前回選択行の場合は処理をスキップ
'            2. Result シートから検索結果（ソース／パス情報）を取得
'            3. 表示用セルへソース文字列とフルパスを反映
'            4. ハイライト有効時はキーワード／コメントの装飾を適用
'
' 引数   : Target - 選択されたセル（Worksheet_SelectionChange から渡される）
'
' 備考   : 直前に処理した行番号は Static 変数で保持し、
'          同一行の再処理を防止している。
'-------------------------------------------------------------------------------
Public Sub UpdateSourceCell(ByVal target As Range)
    Dim wsResult As Worksheet
    
    Dim headerData As GrepHeaderData
    
    Dim src As String, parent As String, file As String, ext As String, fullPth As String, path As String
    
    Static preRowIndex As Long
    
    ' ヘッダ行以前、または同一行が再選択された場合は処理しない
    If target.Row <= ROW_OFFSET_RESULT Or target.Row = preRowIndex Then Exit Sub
    
    Set wsResult = target.Worksheet
    
    ' Result シートのヘッダ情報（検索条件）を取得
    headerData = GetResultHeaderData(wsResult)
    
    ' 選択行から検索結果データを取得
    src = wsResult.Cells(target.Row, COLIDX_RESULT_SOURCE).Value
    parent = wsResult.Cells(target.Row, COLIDX_RESULT_FOLDER).Value
    file = wsResult.Cells(target.Row, COLIDX_RESULT_FILE).Value
    ext = LCase(wsResult.Cells(target.Row, COLIDX_RESULT_EXTENSION).Value)

    ' 先頭が「'」の場合、Excel 表示上消えるためエスケープ用に付与
    If Left(src, 1) = "'" Then

        src = "'" & src
    End If
    
    ' フォルダパスとファイル名を結合してフルパスを生成
    If InStr(headerData.SearchPath, ";") = 0 Then
    
        path = CombinePath(headerData.SearchPath, parent)
    Else
    
        path = parent
    End If
    
    fullPth = CombinePath(path, file)
    
    ' 画面更新・イベントを一時停止
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    ' ソース表示セルへ反映（書式を初期化）
    With wsResult.Range(ADDR_RESULT_SOURCE)
        
        .Value = src
        .Font.Color = vbBlack
        .Font.Bold = False
    End With
    
    ' パス表示セルへフルパスを反映
    wsResult.Range(ADDR_RESULT_PATH).Value = fullPth
    
    ' ハイライト設定が有効な場合のみ装飾を適用
    If headerData.UseHighLight Then
            
        Dim highLightsList As Collection
        Dim hl As Variant
        Dim rngSource As Range
        
        ' キーワード／コメントのハイライト情報を生成
        Set highLightsList = CreateHighLights(fullPth, src, headerData)
        Set rngSource = wsResult.Range(ADDR_RESULT_SOURCE)
        
        ' ハイライト情報に基づき文字装飾を適用
        For Each hl In highLightsList
        
            With rngSource.Characters(Start:=hl(0), length:=hl(1)).Font
                
                .Color = hl(2)
                .Bold = hl(3)
            End With
        Next
        
        Set rngSource = Nothing
    End If
    
    ' 処理済み行を記録
    preRowIndex = target.Row
    
    ' 画面更新・イベントを再開
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    Set wsResult = Nothing
End Sub
