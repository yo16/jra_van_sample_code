Option Explicit On 

Imports System.Text

Module basUtility

    Declare Function GetPrivateProfileString Lib "KERNEL32.DLL" Alias "GetPrivateProfileStringA" ( _
     ByVal lpAppName As String, _
     ByVal lpKeyName As String, ByVal lpDefault As String, _
     ByVal lpReturnedString As StringBuilder, ByVal nSize As Integer, _
     ByVal lpFileName As String) As Integer

    Public strConnectString As String

    ' @(f)
    '
    ' 機能　　　: プロファイルからのデータ取得
    '
    ' 返り値　　: プロファイルデータ
    '
    ' 引き数　　: strAppName - セクション名
    '            strKeyName - キー名
    '            strFileName - プロファイル名
    '
    Public Function GetProfileDataStr(ByVal strAppName As String, ByVal strKeyName As String, ByVal strFileName As String) As String

        Dim iReturnCode As String
        Dim sb As StringBuilder = New StringBuilder(1024)

        ' 文字列を読み出す
        iReturnCode = GetPrivateProfileString(strAppName, strKeyName, "", sb, sb.Capacity, strFileName)

        GetProfileDataStr = sb.ToString

    End Function

    ' @(f)
    '
    ' 機能　　　: データベース接続を行う
    '
    ' 返り値　　: 処理結果(True-正常終了, False-異常終了)
    '
    Public Function ConnectDB() As Boolean
        On Error GoTo ErrorHandler

        Dim bReturnCode As Boolean
        bReturnCode = False

        ' データベースとの接続を行う。
        gCon = New ADODB.Connection()
        Dim strPath As String

        ' データベースのオープン
        gCon.Open(strConnectString)

        bReturnCode = True

ExitHandler:
        ConnectDB = bReturnCode
        Exit Function

ErrorHandler:
        System.Diagnostics.Debug.WriteLine(Err.Description)
        bReturnCode = False

        MsgBox(Err.Description)
        Resume ExitHandler

    End Function

End Module
