Attribute VB_Name = "basAPI"
'
'   API関数宣言モジュール
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   API関数宣言
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private Declare Function GetCursorPos_ Lib "user32" Alias "GetCursorPos" (lpPoint As POINTAPI) As Long
Private Declare Function GetWindowRect_ Lib "user32" Alias "GetWindowRect" (ByVal hwnd As Long, lpRect As RECT) As Long

'この先のAPIプロシッジャーはclsH1Iteratorが使用
Public Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Public Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long

Private Declare Function HtmlHelp_ Lib "hhctrl.ocx" _
  Alias "HtmlHelpA" ( _
  ByVal hwndCaller As Long, _
  ByVal pszFile As String, _
  ByVal uCommand As Long, _
  ByVal dwData As Any) As Long

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部定数（API関数用）
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Const HH_DISPLAY_TOPIC = &H0
Const HH_HELP_CONTEXT = &HF

'この先のConstantsはclsH1Iteratorが使用
Const MAX_PATH = 260
Const MAXDWORD = &HFFFF
Const INVALID_HANDLE_VALUE = -1
Const FILE_ATTRIBUTE_ARCHIVE = &H20
Const FILE_ATTRIBUTE_DIRECTORY = &H10
Const FILE_ATTRIBUTE_HIDDEN = &H2
Const FILE_ATTRIBUTE_NORMAL = &H80
Const FILE_ATTRIBUTE_READONLY = &H1
Const FILE_ATTRIBUTE_SYSTEM = &H4
Const FILE_ATTRIBUTE_TEMPORARY = &H100

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部型宣言（API関数用）
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Public Type POINTAPI
  X As Long
  Y As Long
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'この先のデータタイプはclsH1Iteratorが使用
Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Public Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: マウスカーソルの座標を取得
'
'   備考: なし
'
Public Sub GetCursorPos(ByRef X As Long, ByRef Y As Long)
    Dim p As POINTAPI
    
    If GetCursorPos_(p) <> 0 Then
        X = p.X
        Y = p.Y
    End If
End Sub


'
'   機能: ウィンドウの座標を取得
'
'   備考: なし
'
Public Sub GetWindowRect(ByVal hwnd As Long, ByRef X1 As Long, ByRef Y1 As Long, ByRef X2 As Long, ByRef Y2 As Long)
    Dim r As RECT
    
    If GetWindowRect_(hwnd, r) <> 0 Then
        X1 = r.Left
        Y1 = r.Top
        X2 = r.Right
        Y2 = r.Bottom
    End If
End Sub


'
'   機能: 先頭からNullまでの部分文字列を返す
'
'   備考: なし
'
Public Function StripNulls(OriginalStr As String) As String
    If (InStr(OriginalStr, Chr(0)) > 0) Then
        OriginalStr = Left$(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
    End If
    StripNulls = OriginalStr
End Function


'
'   機能: HTMLヘルプファイルを表示する
'
'   備考: なし
'
Public Sub ShowHtmlHelp()
  HtmlHelp_ 0, App.Path & "\" & cHelpFileName, HH_DISPLAY_TOPIC, ByVal "welcome.htm"
End Sub
