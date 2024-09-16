Attribute VB_Name = "basMain"
'
'   起動モジュール
'
'   いくつかのユーティリティーFunctionを含む
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

' Assertモード エラーログを書いた時点で停止する場合 1
Public ASSERTMODE As Long

Public gApp As clsApp               '' アプリケーションオブジェクト
Public gCC As clsCodeConverter      '' コード変換オブジェクト
Public gSC As clsStringConverter    '' 文字列変換オブジェクト

Public gJVLinkSID As String ' Main関数でExeヘッダから生成します。

Public gDebugCounter_clsGridData As Long
Public gDebugCounter_clsGridItem As Long

Public gColDarkBG As Long
Public gColBG     As Long

Public gstrMDBName(0 To 49) As String '' 直接接続用
Public gCallStack As New Collection

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部定数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Public Const cAppName As String = "馬吉オープンソース版 for DataLab."
Public Const cHelpFileName As String = "Umakichi.chm"

' 色定数
Public Const ColorMother As Long = &HEEEEFF     '母系列（血統）
Public Const ColorFather As Long = &HFFFFFF     '父系列（血統）
Public Const ColorODBack0 As Long = &HFFFFFF    '白（オッズ）
Public Const ColorODBack1 As Long = &H10101     '黒（オッズ）
Public Const ColorODFore0 As Long = &H1FFFF     '黄（オッズ）
Public Const ColorODForeH As Long = &H101FF     '赤（オッズ）
Public Const ColorODForeM As Long = &HFF0101    '青（オッズ）
Public Const ColorODForeL As Long = &H10101     '黒（オッズ）
Public Const ColorLinkExist As Long = &HFF0101  '青（全リンク）
Public Const ColorLinked As Long = &HFF00FF     'ピンク（全リンク）

' その他定数
Public Const cRegistrySubKey As String = "Umakichi5"
Public Const cFromtimeFN As String = "Fromtime.dat"                 ' Fromtime保存ファイル名
Public Const cFromtimeThisWeekFN As String = "FromtimeThisWeek.dat" ' FromtimeThisWeek保存ファイル名

Public Const ASCII_ZERO  As Byte = 48
Public Const ASCII_TWO   As Byte = 50
Public Const ASCII_SEVEN As Byte = 55


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: 起動
'
'   備考: 当VBプロジェクトの開始プロシージャ
'
Public Sub Main()
    Dim splashWindow As frmSplash
    
    ASSERTMODE = 0
    
    ' SIDを生成
    With App
    gJVLinkSID = "Umakichi/OpenSource"
    End With
    
    If App.PrevInstance Then
        End
    End If
    
    ' スプラッシュウインドウでコピーライトの表示
    Set splashWindow = New frmSplash
    splashWindow.Show
    splashWindow.Refresh
    
    Call init
    
    ' アプリケーションオブジェクトの生成
    Set gApp = New clsApp
    Set gCC = New clsCodeConverter
    Set gSC = New clsStringConverter
    ' 起動
    gApp.start
    
    ' スプラッシュウインドウを破棄する
    splashWindow.kill
End Sub


'
'   機能: 引き数のうち、大きいほうの値を返す
'
'   備考: なし
'
Public Function Bigger(a As Long, b As Long) As Long
    Bigger = IIf(a > b, a, b)
End Function


'
'   機能: 引き数のうち、小さいほうの値を返す
'
'   備考: なし
'
Public Function Smaller(a As Long, b As Long) As Long
    Smaller = IIf(a < b, a, b)
End Function


'
'   機能: 文字列中の連続した空白を" "に置き換える
'
'   備考: なし
'
Public Function ContractSpace(str As String) As String
    Dim i    As Long
    Dim p    As Long
    Dim out  As String
    Dim flag As Boolean
    Dim c    As String
    
    p = 1
    flag = False
    For i = 1 To Len(str)
        c = Mid$(str, i, 1)
        If c = " " Or c = "　" Then
            If Not flag Then
                out = out & IIf(p = 1, "", " ") & Mid$(str, p, i - p)
                flag = True
            End If
        Else
            If flag Then
                p = i
                flag = False
            End If
        End If
    Next i
    
    If Not flag Then
        out = out & " " & Mid$(str, p, i - p)
    Else
        out = out & " "
    End If
    
    ContractSpace = out
End Function


'
'   機能: 正規表現でスペースを削除
'
'   備考: なし
'
Public Function DelSpace(strString As String) As String
On Error GoTo Errorhandler
    Dim rx As New RegExp
    With rx
        .Global = True
        .Pattern = "\s|　"
        DelSpace = .Replace(strString, "")
    End With
    Exit Function
Errorhandler:
    gApp.ErrLog
    Resume Next
End Function


'
'   機能: グレイスケールを２値化する
'
'   備考: なし
'
Public Function Contrast(color As Long) As Long
    Dim r As Long
    Dim G As Long
    Dim b As Long
    Dim Gray As Long
    
    r = color Mod 256
    G = (color \ 256) Mod 256
    b = (color \ 65536) Mod 256
    
    Gray = 0.2126 * r ^ 2.2 + 0.7152 * G ^ 2.2 + 0.0724 * b ^ 2.2
    Gray = Gray ^ (1 / 2.2)
    
    If Gray < 128 Then
        Contrast = &HFFFFFF
    Else
        Contrast = &H0
    End If
End Function


'
'   機能: レコードセットを開く
'
'   備考: なし
'
Public Function OpenTableDirect(rs As ADODB.Recordset, cn As ADODB.Connection, TableName As String) As Boolean
On Error GoTo Errorhandler
    rs.CursorLocation = adUseServer
    rs.Index = "PrimaryKey"
    rs.Open TableName, cn, adOpenKeyset, adLockReadOnly, adCmdTableDirect
    OpenTableDirect = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    OpenTableDirect = False
End Function


'
'   機能: コネクションを開放する
'
'   備考: なし
'
Public Sub freecn(cn As ADODB.Connection)

    If Not cn Is Nothing Then
        Do While cn.State And adStateExecuting
            Call cn.Cancel
            gApp.Log "freecn Cancel"
        Loop
        Do While cn.State And adStateOpen
            cn.Close
            gApp.Log "freers Close"
        Loop
        Set cn = Nothing
    Else
        gApp.Log "freecn Nothing"
    End If
End Sub


'
'   機能: レコードセットを閉じるキャンセルユーティリティー
'
'   備考: なし
'
Public Sub freers(rs As ADODB.Recordset)
    If Not rs Is Nothing Then
        Do While rs.State And adStateExecuting
            Call rs.Cancel
            
            gApp.Log "freers Cancel"
        Loop
        Do While rs.State And adStateOpen
            rs.Close
            gApp.Log "freers Close"
        Loop
        Set rs = Nothing
    Else
        gApp.Log "freers Nothing"
    End If
End Sub


'
'   機能: 安全にSeekする
'
'   備考: 普通にSeekすると、Seek出来ていない場合がある為
'
Public Sub SafeSeek(ByRef rs As ADODB.Recordset, ByRef Fields As Variant, ByRef Values As Variant)
On Error GoTo Errorhandler
    Dim i As Long
    Dim c As Long
    Dim NG As Boolean

    If rs.EOF And rs.BOF Then
        Exit Sub
    End If

    rs.MoveFirst
    
    Do
        rs.Seek Values
        If rs.EOF Or rs.BOF Then
            Exit Do
        End If
        
        NG = False
        For i = 0 To UBound(Fields)
            NG = NG Or (rs(Fields(i)) <> Values(i))
        Next i
        
        If Not NG Then
            Exit Do
        End If
        
        gApp.Log "SafeSeek"
        For i = 0 To UBound(Fields)
            gApp.Log c & ":: " & Fields(i) & " : " & rs(Fields(i)) & " <-> " & Values(i)
        Next i
        c = c + 1
        If c > 10 Then
            gApp.Log "SafaSeek failed"
            Exit Do
        End If
    Loop
    
    Exit Sub
Errorhandler:
    gApp.ErrLog
    gApp.Log "SafeSeek Error"
    Resume Next
End Sub


'
'   機能: "&"を"&&"に置換する
'
'   備考: ラベルに"&"を代入すると"_"になる為
'
Public Function ReplaceAmpersand(str As String) As String
    ReplaceAmpersand = Replace(str, "&", "&&")
End Function


'
'   機能: ソートのために空白を"黒"に置き換える
'
'   備考: なし
'
Public Function FormatForSort(str As String) As String
    str = ContractSpace(str)
    If str = Space(1) Then
        str = "黑"          ' "黑" is Unicode's last 漢字 according to value
    Else
        str = Trim$(str)
    End If
    FormatForSort = str
End Function


'
'   機能: レコードセットから値を取得する
'
'   備考: なし
'
Public Function IfExist(rs As ADODB.Recordset, FieldName As String) As String
    If Not rs.EOF Then
        If Not IsNull(rs(FieldName).value) Then
            IfExist = rs(FieldName)
        End If
    End If
End Function


'
'   機能: 空データ(初期値)の判断
'
'   備考: " ","0"だけのデータは"":データなし
'
Public Function IfBe(str As String) As String
    If Space(Len(str)) = str Then
        IfBe = ""
    ElseIf String$(Len(str), "0") = str Then
        IfBe = ""
    Else
        IfBe = str
    End If
End Function


'
'   機能: バイト配列に値を挿入する
'
'   備考: なし
'
Public Sub ByteInsert(ByRef b() As Byte, pos As Long, width As Long, val() As Byte)
    Dim i As Long
    For i = 0 To width - 1
        If i <= UBound(val) Then
            b(pos + i) = val(i)
        End If
    Next i
End Sub


'
'   機能: JVOpenのエラーメッセージを変換する
'
'   備考: なし
'
Public Function ErrMsgJVOpen(lngRet As Long) As String
    Select Case lngRet
    Case 0
        ErrMsgJVOpen = "正常" & vbCrLf & ""
    Case -1
        ErrMsgJVOpen = "該当データ無し" & vbCrLf & "指定されたパラメータに合致する新しいデータがサーバーに存在しない｡又は､最新バージョンが公開され､ユーザーが最新バージョンのダウンロードを選択しました｡JVCloseを呼び出して取り込み処理を終了してください｡"
    Case -2
        ErrMsgJVOpen = "セットアップダイアログでキャンセルが押された" & vbCrLf & "セットアップ用データの取り込み時にユーザーがダイアログでキャンセルを押しました｡JVCloseを呼び出して取り込み処理を終了してください｡ "
    Case -111
        ErrMsgJVOpen = "dataspecパラメータが不正" & vbCrLf & "パラメータの渡し方かパラメータの内容に問題があると思われます｡サンプルプログラム等を参照し､正しくパラメータがJV -Linkに渡っているか確認してください｡ "
    Case -112
        ErrMsgJVOpen = "fromdateパラメータが不正" & vbCrLf & "パラメータの渡し方かパラメータの内容に問題があると思われます｡サンプルプログラム等を参照し､正しくパラメータがJV -Linkに渡っているか確認してください｡ "
    Case -114
        ErrMsgJVOpen = "keyパラメータが不正" & vbCrLf & "パラメータの渡し方かパラメータの内容に問題があると思われます｡サンプルプログラム等を参照し､正しくパラメータがJV -Linkに渡っているか確認してください｡ "
    Case -115
        ErrMsgJVOpen = "optionパラメータが不正" & vbCrLf & "パラメータの渡し方かパラメータの内容に問題があると思われます｡サンプルプログラム等を参照し､正しくパラメータがJV -Linkに渡っているか確認してください｡ "
    Case -116
        ErrMsgJVOpen = "dataspecとoptionの組み合わせが不正" & vbCrLf & "パラメータの渡し方かパラメータの内容に問題があると思われます｡サンプルプログラム等を参照し､正しくパラメータがJV -Linkに渡っているか確認してください｡ "
    Case -201
        ErrMsgJVOpen = "ＪＶＩｎｉｔが行なわれていない" & vbCrLf & "JVOpen/JVRTOpenに先立ってJVInitが呼ばれていないと思われます｡必ずJVInitを先に呼び出してください｡ "
    Case -202
        ErrMsgJVOpen = "前回のJVOpen/JVRTOpenに対してJVCloseが呼ばれていない（オープン中）" & vbCrLf & "前回呼び出したJVOpen/JVRTOpenがJVCloseによってクローズされていないと思われます｡JVOpen/JVRTOpenを呼び出した後は次に呼び出すまでの間にJVCloseを必ず呼び出してください｡ "
    Case -211
        ErrMsgJVOpen = "レジストリ内容が不正（レジストリ内容が不正に変更された）" & vbCrLf & "JV-Linkはレジストリに値をセットする際に値のチェックを行います（例えばサービスキーの桁数など）が、レジストリから値を読み出して使用する際に問題が発生するとこのエラーが発生します｡レジストリが直接書き換えられたなどの状況が考えられない場合にはJRA-VANへご連絡ください。"
    Case -301
        ErrMsgJVOpen = "認証エラー" & vbCrLf & "サービスキーが正しくない。あるいは複数のマシンで同一サービスキーを使用した場合に発生します。複数のマシンで同じサービスキーをしようした場合には、このエラーが発生したマシンのJV-Linkをアンインストールし、再インストール後、利用キーの再発行が必要となります。"
    Case -302
        ErrMsgJVOpen = "サービスキーの有効期限切れ" & vbCrLf & "Data Lab.サービスの有効期限が切れています。サービス権の自動延長が停止していると思われます｡解消するにはサービス権の再購入が必要です｡ "
    Case -303
        ErrMsgJVOpen = "サービスキーが設定されていない（サービスキーが空値）" & vbCrLf & "サービスキーを設定していないと思われます。JVLinkインストール直後はサービスキーが空なので必ず設定する必要があります｡ "
    Case -401
        ErrMsgJVOpen = "JV-Link内部エラー" & vbCrLf & "JV-Link内部でエラーが発生したと思われます。JRAVANへご連絡ください｡ "
    Case -411
        ErrMsgJVOpen = "サーバーエラー（ HTTP ステータス404NotFount）" & vbCrLf & "レジストリが直接変更されたか、Data Lab.用サーバーに問題が発生したと思われます｡JRA -VANのメンテナンス中でない場合で､このエラーが続く場合はJRA-VANへご連絡ください。"
    Case -412
        ErrMsgJVOpen = "サーバーエラー（ HTTP ステータス403Forbidden）" & vbCrLf & "Data Lab.用サーバーに問題が発生したと思われます｡このエラーが続く場合はJRA -VANへご連絡ください｡ "
    Case -413
        ErrMsgJVOpen = "サーバーエラー（HTTPステータス200,403,404以外）" & vbCrLf & "Data Lab.用サーバーに問題が発生したと思われます｡このエラーが続く場合はJRA -VANへご連絡ください｡ "
    Case -421
        ErrMsgJVOpen = "サーバーエラー（サーバーの応答が不正）" & vbCrLf & "Data Lab.用サーバーに問題が発生したと思われます｡このエラーが続く場合はJRA -VANへご連絡ください｡ "
    Case -431
        ErrMsgJVOpen = "サーバーエラー（サーバーアプリケーション内部エラー）" & vbCrLf & "Data Lab.用サーバーに問題が発生したと思われます｡このエラーが続く場合はJRA -VANへご連絡ください｡ "
    Case -501
        ErrMsgJVOpen = "セットアップ処理においてＣＤ－ＲＯＭが無効" & vbCrLf & "JRA-VANが提供した正しいCD-ROMをセットしていないと思われます｡正しいCD -ROMをセットしてください｡ "
    Case -504 '追加
        ErrMsgJVOpen = "サーバーメンテナンス中" & vbCrLf & "サーバーがメンテナンス中です。"
    Case Else
        ErrMsgJVOpen = "想定外のエラーが発生しました。" & vbCrLf & ""
    End Select
End Function



'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: データベース名初期化
'
'   備考: なし
'
Private Sub init()
    gstrMDBName(0) = "subBANUSI.mdb"
    gstrMDBName(1) = "subBATAIJYU.mdb"
    gstrMDBName(2) = "subCHOKYO.mdb"
    gstrMDBName(3) = "subCHOKYO_SEISEKI.mdb"
    gstrMDBName(4) = "subHANRO.mdb"
    gstrMDBName(5) = "subHANSYOKU.mdb"
    gstrMDBName(6) = "subHARAI.mdb"
    gstrMDBName(7) = "subKISHU.mdb"
    gstrMDBName(8) = "subKISHU_CHANGE.mdb"
    gstrMDBName(9) = "subKISHU_SEISEKI.mdb"
    gstrMDBName(10) = "subMINING.mdb"
    gstrMDBName(11) = "subODDS_SANREN0.mdb"
    gstrMDBName(12) = "subODDS_SANREN1.mdb"
    gstrMDBName(13) = "subODDS_SANREN2.mdb"
    gstrMDBName(14) = "subODDS_SANREN3.mdb"
    gstrMDBName(15) = "subODDS_SANREN4.mdb"
    gstrMDBName(16) = "subODDS_SANREN5.mdb"
    gstrMDBName(17) = "subODDS_SANREN6.mdb"
    gstrMDBName(18) = "subODDS_SANREN7.mdb"
    gstrMDBName(19) = "subODDS_SANREN8.mdb"
    gstrMDBName(20) = "subODDS_SANREN9.mdb"
    gstrMDBName(21) = "subODDS_TANPUKUWAKU.mdb"
    gstrMDBName(22) = "subODDS_UMAREN.mdb"
    gstrMDBName(23) = "subODDS_UMATAN0.mdb"
    gstrMDBName(24) = "subODDS_UMATAN1.mdb"
    gstrMDBName(25) = "subODDS_UMATAN2.mdb"
    gstrMDBName(26) = "subODDS_UMATAN3.mdb"
    gstrMDBName(27) = "subODDS_UMATAN4.mdb"
    gstrMDBName(28) = "subODDS_UMATAN5.mdb"
    gstrMDBName(29) = "subODDS_UMATAN6.mdb"
    gstrMDBName(30) = "subODDS_UMATAN7.mdb"
    gstrMDBName(31) = "subODDS_UMATAN8.mdb"
    gstrMDBName(32) = "subODDS_UMATAN9.mdb"
    gstrMDBName(33) = "subODDS_WIDE.mdb"
    gstrMDBName(34) = "subRACE.mdb"
    gstrMDBName(35) = "subRECORD.mdb"
    gstrMDBName(36) = "subSANKU.mdb"
    gstrMDBName(37) = "subSCHEDULE.mdb"
    gstrMDBName(38) = "subSEISAN.mdb"
    gstrMDBName(39) = "subTENKO_BABA.mdb"
    gstrMDBName(40) = "subTOKU.mdb"
    gstrMDBName(41) = "subTOKU_RACE.mdb"
    gstrMDBName(42) = "subTORIKESI_JYOGAI.mdb"
    gstrMDBName(43) = "subUMA.mdb"
    gstrMDBName(44) = "subUMA_RACE_A.mdb"
    gstrMDBName(45) = "subUMA_RACE_B.mdb"
    gstrMDBName(46) = "LinkTables.mdb"
    gstrMDBName(47) = "subRAKaiSel.mdb"
    gstrMDBName(48) = "subHASSOU_CHANGE.mdb"
    gstrMDBName(49) = "subCOURSE_CHANGE.mdb"
End Sub


