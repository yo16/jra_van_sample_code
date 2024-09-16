Attribute VB_Name = "basReg"
'
'   レジストリに関するモジュール
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部定数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
Public Const HKEY_CLASSES_ROOT = &H80000000     'ファイルの関連付け
Public Const HKEY_CURRENT_CONFIG = &H80000005   '
Public Const HKEY_CURRENT_USER = &H80000001     '現在使用しているユーザーの設定
Public Const HKEY_DYN_DATA = &H80000006         '
Public Const HKEY_LOCAL_MACHINE = &H80000002    '複数のユーザーに共通の設定
Public Const HKEY_PERFORMANCE_DATA = &H80000004 '
Public Const HKEY_USERS = &H80000003            '

Public Const REG_DWORD = 4  '文字列
Public Const REG_SZ = 1     '整数

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   API関数宣言
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: レジストリのキーを開く
'
'   備考: なし
'
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" ( _
    ByVal hKey As Long, _
    ByVal lpSubKey As String, _
    phkResult As Long) As Long


'
'   機能: レジストリのキーを作成
'
'   備考: なし
'
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" ( _
    ByVal hKey As Long, _
    ByVal lpSubKey As String, _
    phkResult As Long) As Long


'
'   機能: レジストリデータの取得
'
'   備考: なし
'
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" ( _
    ByVal hKey As Long, _
    ByVal lpValueName As String, _
    ByVal lpReserved As Long, _
    lpType As Long, _
    lpData As Any, _
    lpcbData As Long) As Long


'
'   機能: レジストリデータを書き込む
'
'   備考: なし
'
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" ( _
    ByVal hKey As Long, _
    ByVal lpValueName As String, _
    ByVal Reserved As Long, _
    ByVal dwType As Long, _
    lpData As Any, _
    ByVal cbData As Long) As Long


'
'   機能: レジストリキーを閉じる
'
'   備考: なし
'
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: レジストリデータの取得
'
'   備考: 引き数 RootKey - ルートキー
' 　　　         SubKey  - サブキー
' 　　　         ValueName   - 値の名前
'
Public Function GetRegData(RootKey As Long, SubKey As String, ValueName As String) As String
On Error GoTo ErrH

    Dim hnd         As Long             'キーハンドル
    Dim rtype       As Long             'タイプ
    Dim rlen        As Long             'バッファの長さ
    Dim str         As String * 1024    'バッファ
    Dim regret      As Long             '戻り値
    
    GetRegData = ""
    'レジストリキーを開く
    If RegOpenKey(RootKey, SubKey, hnd) = 0 Then
        rlen = 1024
        'レジストリデータ取得 ( ByValの使用 )
        If RegQueryValueEx(hnd, ValueName, 0, rtype, ByVal str, rlen) = 0 Then
            GetRegData = Left$(str, InStr(str, Chr(0)) - 1)
        End If
    End If
    'レジストリキーを閉じる
    regret = RegCloseKey(hnd)
    
    Exit Function
    
ErrH:
    GetRegData = ""
End Function


'
'   機能: レジストリデータの書き込み
'
'   備考: 引き数 RootKey - ルートキー
' 　　　         SubKey  - サブキー
' 　　　         ValueName   - 値の名前
' 　　　         Value - 書き込む値のデータ
'         返り値 True:正常終了, False:異常終了
'
Public Function SetRegData(RootKey As Long, SubKey As String, ValueName As String, value As String) As Boolean
On Error GoTo ErrH

    Dim hnd         As Long
    Dim regret      As Long
    
    SetRegData = False
    'レジストリキーを開く、キーがなければ作成
    If RegCreateKey(RootKey, SubKey, hnd) = 0 Then
        'データを書き込む
        If RegSetValueEx(hnd, ValueName, 0, REG_SZ, ByVal value, LenB(value)) = 0 Then
            SetRegData = True
        End If
    End If
    'レジストリキーを閉じる
    regret = RegCloseKey(hnd)
    
    Exit Function
    
ErrH:
    SetRegData = False
End Function

