Attribute VB_Name = "basIni"
'
'   INIファイルに関するモジュール
'

Option Explicit


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   API関数宣言
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: 指定された初期化ファイル (.INI ファイル) の、指定されたセクション内にある、
'         指定されたキーに関連付けられている文字列を取得します。
'         関数が成功すると、バッファに格納された文字数が返ります (終端の NULL 文字は含まない) 。
'
'   備考: なし
'
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpDefault As String, _
    ByVal lpReturnedString As String, _
    ByVal nSize As Long, _
    ByVal lpFileName As String) As Long
    

'
'   機能: 指定された初期化ファイル（.INIファイル）の、指定されたセクション内に、
'         指定されたキーとそれに関連付けられた文字列のペアを複数個格納します。
'
'   備考: なし
'
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" ( _
    ByVal lpApplicationName As String, _
    ByVal lpKeyName As Any, _
    ByVal lpString As Any, _
    ByVal lpFileName As String) As Long


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: INIファイルデータの取得
'
'   備考: 引き数 AppName  - ルートキー
' 　　　         KeyName  - サブキー
' 　　　         Default  - 値の名前
'                FileName - INIファイル名
'
Public Function GetIniData(AppName As String, KeyName As String, Default As String, filename As String) As String
On Error GoTo ErrorHandler
    Dim str         As String * 1024    'バッファ
    Dim retuenValue As Long

    GetIniData = ""
    
    'INIファイルデータ取得 ( ByValの使用 )
    retuenValue = GetPrivateProfileString(AppName, KeyName, Default, ByVal str, 1024, filename)
    If retuenValue > 0 Then
        GetIniData = Left$(str, retuenValue)
    End If
    
    Exit Function
ErrorHandler:
    GetIniData = ""
End Function

'
'   機能: INIファイルデータの書き込み
'
'   備考: 引き数 AppName  - セクション名
' 　　　         KeyName  - キー名
' 　　　         Value    - 値
' 　　　         FileName - INIファイル名
'
Public Function SetIniData(AppName As String, KeyName As String, value As String, filename As String) As Boolean
On Error GoTo ErrorHandler
    SetIniData = False

    'データを書き込む
    If WritePrivateProfileString(AppName, KeyName, ByVal value, filename) <> 0 Then
        SetIniData = True
    End If
    
    gApp.Log "SetIniData : " & AppName & "." & KeyName & "=" & value & " : " & filename
    
    Exit Function
ErrorHandler:
    gApp.ErrLog
    SetIniData = False
End Function

