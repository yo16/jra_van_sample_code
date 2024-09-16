Attribute VB_Name = "basEnum"
'
'   Enum宣言モジュール
'

Option Explicit


'
'   機能: JV-Link 取得モード
'
'   備考: なし
'
Public Enum ukJVLMode
    ukjUsual
    ukjThisWeek
End Enum


'
'   機能: 速報取得モード
'
'   備考: なし
'
Public Enum ukPromptMode
    ukpRA
    ukpOD
    ukpPALLET
End Enum


'
'   機能: データエクスポートモード
'
'   備考: なし
'
Public Enum ukExportMode
    ukeJVDATA
    ukeCSV
End Enum


'
'   機能: データエクスポートフィルタモード
'
'   備考: なし
'
Public Enum ukExportFilter
    ukfNone
    ukfDate
    ukfJyoCD
    ukfDateJyoCD
End Enum


'
'   機能: ctlPane モード
'
'   備考: なし
'
Public Enum ukCtlPaneMode
    ukcpNowFetching
    ukcpNoData
    ukcpShowControls
    ukcpHideControls
End Enum

