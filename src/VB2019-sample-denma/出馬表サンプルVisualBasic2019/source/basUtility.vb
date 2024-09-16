Option Explicit On 

Imports System.Text

Module basUtility

    Declare Function GetPrivateProfileString Lib "KERNEL32.DLL" Alias "GetPrivateProfileStringA" ( _
     ByVal lpAppName As String, _
     ByVal lpKeyName As String, ByVal lpDefault As String, _
     ByVal lpReturnedString As StringBuilder, ByVal nSize As Integer, _
     ByVal lpFileName As String) As Integer

    Declare Function WritePrivateProfileString Lib "KERNEL32.DLL" Alias "WritePrivateProfileStringA" ( _
     ByVal lpAppName As String, _
     ByVal lpKeyName As String, _
     ByVal lpString As String, _
     ByVal lpFileName As String) As Integer

    Public ImportRA As clsImportRA
    Public ImportSE As clsImportSE
    Public ImportUM As clsImportUM

    Public objCDCv As clsCodeConv

    Public strConnectString As String

    ' @(f)
    '
    ' 機能　　　: 指定バイト数まで空白を「右に」付け加える
    '
    ' 返り値　　: 指定バイト数 文字列
    '
    ' 引き数　　: strpad - 対象文字列(byte指定特に無し)
    '          : totalBytes - 指定するバイト数
    ' 
    ' 備考     : 通常のPadRightは全角文字(2byte)も1文字と数えるので、
    '            バイト数で指定することで全角半角が混じった文字列も長さを合わせることが容易となる
    ' 
    Public Function bPadR(ByVal strPad As String, ByVal totalBytes As Integer) As String

        Dim strReturn As String
        Dim intTmp As Integer
        Dim bBuff As Byte()
        Dim bSize As Long

        If strPad Is Nothing Then
            strPad = ""
        End If
        bSize = Str2Byte(strPad).Length
        bBuff = New Byte(bSize) {}

        bBuff = Str2Byte(strPad)
        If bBuff.Length < totalBytes Then
            If bBuff.Length.Equals(strPad.Length) Then
                strReturn = strPad.PadRight(totalBytes)
            Else
                intTmp = totalBytes - (bBuff.Length - strPad.Length)
                strReturn = strPad.PadRight(intTmp)
            End If
        Else
            strReturn = strPad
        End If

        bPadR = strReturn
    End Function

    ' @(f)
    '
    ' 機能　　　: 指定バイト数まで空白を「左に」付け加える
    '
    ' 返り値　　: 指定バイト数 文字列
    '
    ' 引き数　　: strpad - 対象文字列(byte指定特に無し)
    '          : totalBytes - 指定するバイト数
    ' 
    ' 備考     : 通常のPadLeftは全角文字(2byte)も1文字と数えるので、
    '            バイト数で指定することで全角半角が混じった文字列も長さを合わせることが容易となる
    ' 
    Public Function bPadL(ByVal strPad As String, ByVal totalBytes As Integer) As String

        Dim strReturn As String
        Dim intTmp As Integer
        Dim bBuff As Byte()
        Dim bSize As Long

        If strPad Is Nothing Then
            strPad = ""
        End If
        bSize = Str2Byte(strPad).Length
        bBuff = New Byte(bSize) {}

        bBuff = Str2Byte(strPad)
        If bBuff.Length < totalBytes Then
            If bBuff.Length.Equals(strPad.Length) Then
                strReturn = strPad.PadLeft(totalBytes)
            Else
                intTmp = totalBytes - (bBuff.Length - strPad.Length)
                strReturn = strPad.PadLeft(intTmp)
            End If
        Else
            strReturn = strPad
        End If

        bPadL = strReturn
    End Function


    ' @(f)
    '
    ' 機能　　　: 文字列から空白部分を抜く
    '
    ' 返り値　　: 空白を抜く対象の文字列
    '
    ' 引き数　　: strTrim - 文字列(byte指定特に無し)
    '
    Public Function TrimSP(ByVal strTrim As String) As String

        Dim strReturn As String
        Dim i As Short ' ループカウンタ
        Dim strTmp As String
        strReturn = ""

        For i = 0 To strTrim.Length - 1
            ' 対象文字列の先頭から1文字ずつ調べる
            strTmp = strTrim.Substring(i, 1)
            ' 半角、全角の空白を除き、文字列に格納する
            If strTmp.Equals(" ") Or strTmp.Equals("　") Then
            Else
                strReturn = strReturn & strTmp
            End If
        Next i

        TrimSP = strReturn
    End Function


    ' @(f)
    '
    ' 機能　　　: グレードコード の 名称 を可読文字列にする
    '
    ' 返り値　　: 名称 文字列
    '
    ' 引き数　　: strCD - グレードコード文字列(1byte)
    '
    Public Function GRAD2(ByVal strCD As String) As String
        GRAD2 = String.Empty
        Select Case strCD
            Case "A"
                GRAD2 = "(ＧⅠ)"
            Case "B"
                GRAD2 = "(ＧⅡ)"
            Case "C"
                GRAD2 = "(ＧⅢ)"
            Case "D"
                GRAD2 = ""
            Case "E"
                GRAD2 = ""
            Case "F"
                GRAD2 = "(J･ＧⅠ)"
            Case "G"
                GRAD2 = "(J･ＧⅡ)"
            Case "H"
                GRAD2 = "(J･ＧⅢ)"
            Case " "
                GRAD2 = ""
        End Select
    End Function

    ' @(f)
    '
    ' 機能　　　: グレードコード の 名称 を可読文字列にする
    '
    ' 返り値　　: 名称 文字列
    '
    ' 引き数　　: strCD - グレードコード文字列(1byte)
    '
    Public Function GRAD3(ByVal strCD As String) As String
        GRAD3 = String.Empty
        Select Case strCD
            Case "A"
                GRAD3 = "(ＧⅠ)"
            Case "B"
                GRAD3 = "(ＧⅡ)"
            Case "C"
                GRAD3 = "(ＧⅢ)"
            Case "D"
                GRAD3 = ""
            Case "E"
                GRAD3 = ""
            Case "F"
                GRAD3 = "(JGⅠ)"
            Case "G"
                GRAD3 = "(JGⅡ)"
            Case "H"
                GRAD3 = "(JGⅢ)"
            Case " "
                GRAD3 = ""
        End Select
    End Function

    ' @(f)
    '
    ' 機能　　　: 競走種別コード の 名称 を可読文字列にする
    '
    ' 返り値　　: 名称 文字列
    '
    ' 引き数　　: strCD - 競走種別コード文字列(2byte)
    '
    Public Function KSSB6(ByVal strCD As String) As String
        KSSB6 = String.Empty
        Select Case strCD
            Case "0"
                KSSB6 = ""
            Case "11"
                KSSB6 = "２歳" '"サラブレッド系2歳"
            Case "12"
                KSSB6 = "３歳" '"サラブレッド系3歳"
            Case "13"
                KSSB6 = "３歳上" '"サラブレッド系3歳以上"
            Case "14"
                KSSB6 = "４歳上" '"サラブレッド系4歳以上"
            Case "18"
                KSSB6 = "３歳上" '"サラブレッド系障害3歳以上"
            Case "19"
                KSSB6 = "４歳上" '"サラブレッド系障害4歳以上"
            Case "21"
                KSSB6 = "２歳" '"アラブ系2歳"
            Case "22"
                KSSB6 = "３歳" '"アラブ系3歳"
            Case "23"
                KSSB6 = "３歳上" '"アラブ系3歳以上"
            Case "24"
                KSSB6 = "４歳上" '"アラブ系4歳以上"
        End Select
    End Function

    ' @(f)
    '
    ' 機能　　　: 競走種別コード の 名称 を可読文字列にする
    '
    ' 返り値　　: 名称 文字列
    '
    ' 引き数　　: strCD - 競走種別コード文字列(2byte)
    '
    Public Function KSSB7(ByVal strCD As String) As String
        KSSB7 = String.Empty
        Select Case strCD
            Case "0"
                KSSB7 = ""
            Case "11"
                KSSB7 = "２歳" '"サラブレッド系2歳"
            Case "12"
                KSSB7 = "３歳" '"サラブレッド系3歳"
            Case "13"
                KSSB7 = "３歳上" '"サラブレッド系3歳以上"
            Case "14"
                KSSB7 = "４歳上" '"サラブレッド系4歳以上"
            Case "18"
                KSSB7 = "障害３歳上" '"サラブレッド系障害3歳以上"
            Case "19"
                KSSB7 = "障害４歳上" '"サラブレッド系障害4歳以上"
            Case "21"
                KSSB7 = "２歳" '"アラブ系2歳"
            Case "22"
                KSSB7 = "３歳" '"アラブ系3歳"
            Case "23"
                KSSB7 = "３歳上" '"アラブ系3歳以上"
            Case "24"
                KSSB7 = "４歳上" '"アラブ系4歳以上"
        End Select
    End Function

    ' @(f)
    '
    ' 機能　　　: 競走条件コード の 名称 を可読文字列にする
    '
    ' 返り値　　: 名称 文字列
    '
    ' 引き数　　: strCD - 競走条件コード文字列(3byte)
    '
    Public Function KSJK4(ByVal strCD As String) As String
        KSJK4 = String.Empty
        Select Case Val(strCD)
            Case 0
                KSJK4 = ""
            Case 1 To 99
                KSJK4 = 100 * Val(strCD) & "万下"
            Case 100
                KSJK4 = "１億"
            Case "701"
                KSJK4 = "新馬"
            Case "702"
                KSJK4 = "未出走"
            Case "703"
                KSJK4 = "未勝利"
            Case "999"
                KSJK4 = "ｵｰﾌﾟﾝ"
        End Select
    End Function

    ' @(f)
    '
    ' 機能　　　: 性別コード の 名称 を可読文字列にする
    '
    ' 返り値　　: 略称 文字列
    '
    ' 引き数　　: strCD - 性別コード文字列(1byte)
    '
    Public Function SEIB4(ByVal strCD As String) As String
        SEIB4 = String.Empty
        Select Case strCD
            Case "0"
                SEIB4 = ""
            Case "1"
                SEIB4 = "牡"
            Case "2"
                SEIB4 = "牝"
            Case "3"
                SEIB4 = "騙"
        End Select
    End Function

    ' @(f)
    '
    ' 機能　　　: 枠背景色の指定
    '
    ' 返り値　　: 枠背景色 RGB値(&Hで16進数表記)
    '
    ' 引き数　　: strCD - 枠番文字列(1byte)
    '
    Public Function CELBK1(ByVal strCD As String) As String
        CELBK1 = String.Empty
        Select Case strCD
            Case "0"
                CELBK1 = "&HFFFFFF"
            Case "1"
                CELBK1 = "&HFFFFFF"
            Case "2"
                CELBK1 = "&H010000"
            Case "3"
                CELBK1 = "&HFF0000"
            Case "4"
                CELBK1 = "&H0000FF"
            Case "5"
                CELBK1 = "&HFFFF00"
            Case "6"
                CELBK1 = "&H00FF00"
            Case "7"
                CELBK1 = "&HFF8000"
            Case "8"
                CELBK1 = "&HFF8080"
        End Select
    End Function

    ' @(f)
    '
    ' 機能　　　: 着順背景色の指定
    '
    ' 返り値　　: 着順背景色 RGB値(&Hで16進数表記)
    '
    ' 引き数　　: strCD - 着順文字列(1byte)
    '
    Public Function CELBK2(ByVal strCD As String) As String
        CELBK2 = String.Empty
        Select Case strCD
            Case "01"
                CELBK2 = "&HFFCCCC"
            Case "02"
                CELBK2 = "&HFFCC80"
            Case "03"
                CELBK2 = "&HCCFFFF"
        End Select
    End Function

    ' @(f)
    '
    ' 機能　　　: 枠文字色の指定
    '
    ' 返り値　　: 枠文字色 文字列
    '
    ' 引き数　　: strCD - 枠番文字列(1byte)
    '
    Public Function CELFK(ByVal strCD As String) As String
        CELFK = String.Empty
        Select Case strCD
            Case "0"
                CELFK = ""
            Case "1"
                CELFK = ""
            Case "2"
                CELFK = "White"
            Case "3"
                CELFK = "White"
            Case "4"
                CELFK = "White"
            Case "5"
                CELFK = ""
            Case "6"
                CELFK = ""
            Case "7"
                CELFK = ""
            Case "8"
                CELFK = ""
        End Select
    End Function

    ' @(f)
    '
    ' 機能　　　: トラックコード の 名称 を可読文字列にする
    '
    ' 返り値　　: 名称 文字列
    '
    ' 引き数　　: strCD - トラックコード文字列(2byte)
    '
    Public Function TRCK4(ByVal strCD As String) As String
        TRCK4 = String.Empty
        Select Case strCD
            Case "00"
                TRCK4 = ""
            Case "10" To "22"
                TRCK4 = "芝"
            Case "23" To 26, "29"
                TRCK4 = "ダ"
            Case "27", "28"
                TRCK4 = "砂"
            Case "51" To "59"
                TRCK4 = "障"
        End Select
    End Function

    ' @(f)
    '
    ' 機能　　　: トラックコード の 名称 を可読文字列にする
    '
    ' 返り値　　: 名称 文字列
    '
    ' 引き数　　: strCD - トラックコード文字列(2byte)
    '
    Public Function TRCK5(ByVal strCD As String) As String
        TRCK5 = String.Empty
        Select Case strCD
            Case "00"
                TRCK5 = ""
            Case "10"
                TRCK5 = "芝直"
            Case "11" To "16"
                TRCK5 = "芝左"
            Case "17" To "22"
                TRCK5 = "芝右"
            Case "23", "25"
                TRCK5 = "ダ左"
            Case "24", "26"
                TRCK5 = "ダ右"
            Case "27"
                TRCK5 = "砂左"
            Case "28"
                TRCK5 = "砂右"
            Case "29"
                TRCK5 = "ダ直"
            Case "51" To "59"
                TRCK5 = "障害"
        End Select
    End Function

    ' @(f)
    '
    ' 機能　　　: 馬場状態コード の 名称 を可読文字列にする
    '
    ' 返り値　　: 名称 文字列
    '
    ' 引き数　　: strCD - 馬場状態コード文字列(1byte)
    '
    Public Function BBJT4(ByVal strCD As String) As String
        BBJT4 = String.Empty
        Select Case strCD
            Case "0"
                BBJT4 = ""
            Case "1"
                BBJT4 = "良"
            Case "2"
                BBJT4 = "稍"
            Case "3"
                BBJT4 = "重"
            Case "4"
                BBJT4 = "不"
        End Select
    End Function

    ' @(f)
    '
    ' 機能　　　: データ区分 を可読文字列にする
    '
    ' 返り値　　: 区分 文字列
    '
    ' 引き数　　: strCD - データ区分文字列(1byte)
    '
    Public Function DTKB1(ByVal strCD As String) As String
        DTKB1 = String.Empty
        Select Case strCD
            Case "1"
                DTKB1 = "出走馬名表(木曜)"
            Case "2"
                DTKB1 = "出馬表(金・土曜)"
            Case "3"
                DTKB1 = "速報成績(3着まで確定)"
            Case "4"
                DTKB1 = "速報成績(5着まで確定)"
            Case "5"
                DTKB1 = "速報成績(全馬着順確定)"
            Case "6"
                DTKB1 = "速報成績(全馬着順+ｺｰﾅｰ通過順)"
            Case "7"
                DTKB1 = "成績(月曜)"
            Case "A"
                DTKB1 = "地方競馬"
            Case "B"
                DTKB1 = "海外国際レース"
            Case "9"
                DTKB1 = "レース中止"
            Case "0"
                DTKB1 = "該当レコード削除"
        End Select
    End Function

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
        iReturnCode = GetPrivateProfileString(strAppName, strKeyName, "default", sb, sb.Capacity, strFileName)

        GetProfileDataStr = sb.ToString

    End Function

    ' @(f)
    '
    ' 機能　　　: プロファイルへのデータ設定
    '
    ' 引き数　　: strAppName - セクション名
    '            strKeyName - キー名
    '            strValue - 設定値
    '            strFileName - プロファイル名
    '
    Public Sub WriteProfileDataStr(ByVal strAppName As String, ByVal strKeyName As String, ByVal strValue As String, ByVal strFileName As String)

        WritePrivateProfileString(strAppName, strKeyName, strValue, strFileName)

    End Sub


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

        '接続文字列

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
