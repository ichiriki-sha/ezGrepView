Attribute VB_Name = "modConstants"
Option Explicit

'-------------------------------------------------------------------------------
' シート名
'-------------------------------------------------------------------------------
Public Const WSNM_MAIN                              As String = "Main"
Public Const WSNM_RESULT                            As String = "Result"
Public Const WSNM_COMMENT                           As String = "Comment"

'-------------------------------------------------------------------------------
' アドレス
'-------------------------------------------------------------------------------
'メインシート
Public Const ADDR_MAIN_RESULT_PATH                  As String = "E3"
Public Const ADDR_MAIN_ENCODING                     As String = "E5"
Public Const ADDR_MAIN_USE_HIGHLIGHT                As String = "E7"
Public Const ADDR_MAIN_COMMENT_MARK                 As String = "E11"
Public Const ADDR_MAIN_BINARY_MARK                  As String = "E13"
Public Const ADDR_MAIN_GARBLED_MARK                 As String = "E15"

'結果シート
Public Const ADDR_RESULT_SEARCH_PATH                As String = "C1"
Public Const ADDR_RESULT_FILE                       As String = "C2"
Public Const ADDR_RESULT_PATTERN                    As String = "C3"
Public Const ADDR_RESULT_USE_REGEXP                 As String = "K2"
Public Const ADDR_RESULT_TEXTONLY                   As String = "K3"
Public Const ADDR_RESULT_ALL_DIRECTORIES            As String = "M1"
Public Const ADDR_RESULT_IGNORECASE                 As String = "M2"
Public Const ADDR_RESULT_USE_HIGHLIGHT              As String = "M3"
Public Const ADDR_RESULT_SOURCE                     As String = "A4"
Public Const ADDR_RESULT_PATH                       As String = "A13"

'-------------------------------------------------------------------------------
' 列番号
'-------------------------------------------------------------------------------
Public Const COLIDX_RESULT_ROWIDX                   As Integer = 1
Public Const COLIDX_RESULT_FOLDER                   As Integer = 2
Public Const COLIDX_RESULT_FOLDER_FILLER1           As Integer = 3
Public Const COLIDX_RESULT_FILE                     As Integer = 4
Public Const COLIDX_RESULT_EXTENSION                As Integer = 5
Public Const COLIDX_RESULT_POSITION                 As Integer = 6
Public Const COLIDX_RESULT_ENCODING                 As Integer = 7
Public Const COLIDX_RESULT_RESULT                   As Integer = 8
Public Const COLIDX_RESULT_SOURCE                   As Integer = 9
Public Const COLIDX_RESULT_SOURCE_FILLER1           As Integer = 10
Public Const COLIDX_RESULT_SOURCE_FILLER2           As Integer = 11
Public Const COLIDX_RESULT_SOURCE_FILLER3           As Integer = 12
Public Const COLIDX_RESULT_SOURCE_FILLER4           As Integer = 13

'-------------------------------------------------------------------------------
' 行オフセット
'-------------------------------------------------------------------------------
Public Const ROW_OFFSET_RESULT                      As Integer = 14

'-------------------------------------------------------------------------------
' オートフィルタ
'-------------------------------------------------------------------------------
Public Const AUTO_FILTER_RESULT                     As String = "14:14"

'-------------------------------------------------------------------------------
' Grep結果ヘッダー最大行数
'-------------------------------------------------------------------------------
Public Const GREP_HEADER_MAX_ROWS  As Long = 30

'-------------------------------------------------------------------------------
' Grep結果ヘッダーパターン
'-------------------------------------------------------------------------------
Public Const GREP_HEADER_PATTERN_SEARCH_PATH        As String = "^フォルダー?\s+(.*)$"
Public Const GREP_HEADER_PATTERN_FILENAME           As String = "^検索対象\s+(.*)$"
Public Const GREP_HEADER_PATTERN_PATTERN            As String = "^□検索条件\s+""(.*)""$"
Public Const GREP_HEADER_PATTERN_ALL_DIRECTORIES    As String = "^\s+\(サブフォルダー?も検索\)"
Public Const GREP_HEADER_PATTERN_USE_REGEXP         As String = "^\s+\(正規表現.+\)"
Public Const GREP_HEADER_PATTERN_IGNORECASE         As String = "^\s+\(英大文字小文字を区別しない\)"
Public Const GREP_HEADER_PATTERN_TEXTONLY           As String = "^\s+\(テキストのみ検索\)"

'-------------------------------------------------------------------------------
' 各種表示名
'-------------------------------------------------------------------------------
'サブフォルダ検索
Public Const DISP_ALL_DIRECTORIES_YES               As String = "検索する"
Public Const DISP_ALL_DIRECTORIES_NO                As String = "検索しない"
Public Const CSV_ALL_DIRECTORIES                    As String = DISP_ALL_DIRECTORIES_YES & "," & DISP_ALL_DIRECTORIES_NO

'正規表現
Public Const DISP_USE_REGEXP_YES                    As String = "使用する"
Public Const DISP_USE_REGEXP_NO                     As String = "使用しない"
Public Const CSV_USE_REGEXP                         As String = DISP_USE_REGEXP_YES & "," & DISP_USE_REGEXP_NO

'大文字小文字区別
Public Const DISP_IGNORECASE_YES                    As String = "区別する"
Public Const DISP_IGNORECASE_NO                     As String = "区別しない"
Public Const CSV_IGNORECASE                         As String = DISP_IGNORECASE_YES & "," & DISP_IGNORECASE_NO

'ハイライト
Public Const DISP_USE_HIGHLIGHT_YES                 As String = "する"
Public Const DISP_USE_HIGHLIGHT_NO                  As String = "しない"
Public Const CSV_USE_HIGHLIGHT                      As String = DISP_USE_HIGHLIGHT_YES & "," & DISP_USE_HIGHLIGHT_NO

'テキストのみ検索
Public Const DISP_TEXTONLY_YES                      As String = "する"
Public Const DISP_TEXTONLY_NO                       As String = "しない"
Public Const CSV_TEXTONLY                           As String = DISP_TEXTONLY_YES & "," & DISP_TEXTONLY_NO

'コメントタイプ
Public Const DISP_COMMENT_TYPE_LINE                 As String = "ライン"
Public Const DISP_COMMENT_TYPE_BLOCK                As String = "ブロック"
Public Const CSV_COMMENT_TYPE                       As String = DISP_COMMENT_TYPE_LINE & "," & DISP_COMMENT_TYPE_BLOCK

' 文字コード
Public Const CSV_ENCODING                           As String = ENCODING_UTF8 & "," & ENCODING_SJIS

'-------------------------------------------------------------------------------
' 各種列挙型
'-------------------------------------------------------------------------------
'Comment
Public Enum CommentType
    Line = 0
    Block = 1
End Enum

'-------------------------------------------------------------------------------
' 構造体
'-------------------------------------------------------------------------------
' Input情報
Public Type InputData
    ResultPath                                      As String
    Encoding                                        As String
    UseHighLight                                    As Boolean
    CommentMark                                     As String
    BinaryMark                                      As String
    GarbledMark                                     As String
End Type

' GrepHeader情報
Public Type GrepHeaderData
    SearchPath                                      As String
    FileName                                        As String
    Keyword                                         As String
    AllDirectories                                  As Boolean
    UseRegExp                                       As Boolean
    IgnoreCase                                      As Boolean
    TextOnly                                        As Boolean
    UseHighLight                                    As Boolean
End Type

' Comment情報
Public Type CommentData
    Extension                                       As String
    CommentType                                     As CommentType
    CommentStart                                    As String
    CommentEnd                                      As String
End Type
