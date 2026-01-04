Attribute VB_Name = "modBusinessWorkbook"
Option Explicit

'-------------------------------------------------------------------------------
' 関数名 : ValidateRequiredSheets
' 概要   : 処理に必要なワークシートがすべて存在するかを検証する
'
'          必須シート：
'            ・Main    （入力シート）
'            ・Result  （結果表示シート）
'            ・Comment （コメント定義シート）
'
' 引数   : wb - 対象となるブック
' 戻り値 : True  - すべての必須シートが存在する
'          False - いずれかのシートが存在しない
'-------------------------------------------------------------------------------
Public Function ValidateRequiredSheets(ByVal wb As Workbook) As Boolean

    ValidateRequiredSheets = _
        FindWorksheet(WSNM_MAIN, wb) And _
        FindWorksheet(WSNM_RESULT, wb) And _
        FindWorksheet(WSNM_COMMENT, wb)

End Function

'-------------------------------------------------------------------------------
' 関数名 : IsInputMode
' 概要   : 現在のブックが「入力モード」かどうかを判定する
'
'          判定基準：
'            ・Main シートが表示状態（Visible）の場合は入力モード
'
' 引数   : wb - 対象となるブック
' 戻り値 : True  - 入力モード
'          False - 結果モード
'-------------------------------------------------------------------------------
Public Function IsInputMode(ByVal wb As Workbook) As Boolean

    IsInputMode = _
        wb.Worksheets(WSNM_MAIN).Visible = xlSheetVisible
End Function

'-------------------------------------------------------------------------------
' 関数名 : InitializeInputMode
' 概要   : 入力モード用の初期化処理を行う
'
'         ・Result シートを非表示（VeryHidden）
'         ・Comment シートを表示
'         ・Main シートの入力項目に入力規則（ドロップダウン）を設定
'
' 引数   : wb - 対象となるブック
'-------------------------------------------------------------------------------
Public Sub InitializeInputMode(ByVal wb As Workbook)

    Dim wsMain As Worksheet
    Dim wsResult As Worksheet
    Dim wsComment As Worksheet

    Set wsMain = wb.Worksheets(WSNM_MAIN)
    Set wsResult = wb.Worksheets(WSNM_RESULT)
    Set wsComment = wb.Worksheets(WSNM_COMMENT)

    wsResult.Visible = xlSheetVeryHidden
    wsComment.Visible = xlSheetVisible

    SetValidateList wsMain.Range(ADDR_MAIN_ENCODING), CSV_ENCODING
    SetValidateList wsMain.Range(ADDR_MAIN_USE_HIGHLIGHT), CSV_USE_HIGHLIGHT

    Set wsMain = Nothing
    Set wsResult = Nothing
    Set wsComment = Nothing
End Sub

'-------------------------------------------------------------------------------
' 関数名 : InitializeResultMode
' 概要   : 結果モード用の初期化処理を行う
'
'         ・Result シートを表示
'         ・Comment シートを表示
'         ・Ctrl + J にサクラエディタ起動処理を割り当てる
'
' 引数   : wb - 対象となるブック
'-------------------------------------------------------------------------------
Public Sub InitializeResultMode(ByVal wb As Workbook)

    Dim wsResult As Worksheet
    Dim wsComment As Worksheet

    Set wsResult = wb.Worksheets(WSNM_RESULT)
    Set wsComment = wb.Worksheets(WSNM_COMMENT)

    wsResult.Visible = xlSheetVisible
    wsComment.Visible = xlSheetVisible

    Application.OnKey "^j", "OpenSakura"

    Set wsResult = Nothing
    Set wsComment = Nothing
End Sub
