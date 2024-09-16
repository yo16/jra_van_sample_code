Option Explicit On 

Public Class frmMenu
    Inherits System.Windows.Forms.Form

#Region " Windows フォーム デザイナで生成されたコード "

    Public Sub New()
        MyBase.New()

        ' この呼び出しは Windows フォーム デザイナで必要です。
        InitializeComponent()

        ' InitializeComponent() 呼び出しの後に初期化を追加します。

    End Sub

    ' Form は dispose をオーバーライドしてコンポーネント一覧を消去します。
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    ' Windows フォーム デザイナで必要です。
    Private components As System.ComponentModel.IContainer

    ' メモ : 以下のプロシージャは、Windows フォーム デザイナで必要です。
    ' Windows フォーム デザイナを使って変更してください。  
    ' コード エディタは使用しないでください。
    Friend WithEvents btnGetJVData As System.Windows.Forms.Button
    Friend WithEvents cmbYear As System.Windows.Forms.ComboBox
    Friend WithEvents btnInitDB As System.Windows.Forms.Button
    Friend WithEvents btnSettingJVLink As System.Windows.Forms.Button
    Friend WithEvents btnViewDenmaList As System.Windows.Forms.Button
    Friend WithEvents btnStopJVData As System.Windows.Forms.Button
    Friend WithEvents barFileCount As System.Windows.Forms.ProgressBar
    Friend WithEvents barReadSize As System.Windows.Forms.ProgressBar
    Friend WithEvents JVLink As AxJVDTLabLib.AxJVLink
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmMenu))
        Me.btnGetJVData = New System.Windows.Forms.Button()
        Me.barFileCount = New System.Windows.Forms.ProgressBar()
        Me.cmbYear = New System.Windows.Forms.ComboBox()
        Me.btnInitDB = New System.Windows.Forms.Button()
        Me.btnViewDenmaList = New System.Windows.Forms.Button()
        Me.btnSettingJVLink = New System.Windows.Forms.Button()
        Me.btnStopJVData = New System.Windows.Forms.Button()
        Me.barReadSize = New System.Windows.Forms.ProgressBar()
        Me.JVLink = New AxJVDTLabLib.AxJVLink()
        CType(Me.JVLink, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnGetJVData
        '
        Me.btnGetJVData.Location = New System.Drawing.Point(8, 8)
        Me.btnGetJVData.Name = "btnGetJVData"
        Me.btnGetJVData.Size = New System.Drawing.Size(88, 40)
        Me.btnGetJVData.TabIndex = 1
        Me.btnGetJVData.Text = "開催情報取得"
        '
        'barFileCount
        '
        Me.barFileCount.Location = New System.Drawing.Point(0, 264)
        Me.barFileCount.Name = "barFileCount"
        Me.barFileCount.Size = New System.Drawing.Size(200, 24)
        Me.barFileCount.TabIndex = 2
        '
        'cmbYear
        '
        Me.cmbYear.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbYear.Location = New System.Drawing.Point(8, 64)
        Me.cmbYear.Name = "cmbYear"
        Me.cmbYear.Size = New System.Drawing.Size(184, 20)
        Me.cmbYear.TabIndex = 3
        '
        'btnInitDB
        '
        Me.btnInitDB.Location = New System.Drawing.Point(8, 144)
        Me.btnInitDB.Name = "btnInitDB"
        Me.btnInitDB.Size = New System.Drawing.Size(88, 40)
        Me.btnInitDB.TabIndex = 4
        Me.btnInitDB.Text = "ＤＢ初期化"
        '
        'btnViewDenmaList
        '
        Me.btnViewDenmaList.Location = New System.Drawing.Point(8, 88)
        Me.btnViewDenmaList.Name = "btnViewDenmaList"
        Me.btnViewDenmaList.Size = New System.Drawing.Size(184, 40)
        Me.btnViewDenmaList.TabIndex = 5
        Me.btnViewDenmaList.Text = "出馬表表示"
        '
        'btnSettingJVLink
        '
        Me.btnSettingJVLink.Location = New System.Drawing.Point(104, 144)
        Me.btnSettingJVLink.Name = "btnSettingJVLink"
        Me.btnSettingJVLink.Size = New System.Drawing.Size(88, 40)
        Me.btnSettingJVLink.TabIndex = 6
        Me.btnSettingJVLink.Text = "JV-Link設定"
        '
        'btnStopJVData
        '
        Me.btnStopJVData.Enabled = False
        Me.btnStopJVData.Location = New System.Drawing.Point(104, 8)
        Me.btnStopJVData.Name = "btnStopJVData"
        Me.btnStopJVData.Size = New System.Drawing.Size(88, 40)
        Me.btnStopJVData.TabIndex = 7
        Me.btnStopJVData.Text = "キャンセル"
        '
        'barReadSize
        '
        Me.barReadSize.Location = New System.Drawing.Point(0, 288)
        Me.barReadSize.Name = "barReadSize"
        Me.barReadSize.Size = New System.Drawing.Size(200, 24)
        Me.barReadSize.TabIndex = 8
        '
        'JVLink
        '
        Me.JVLink.Enabled = True
        Me.JVLink.Location = New System.Drawing.Point(104, 200)
        Me.JVLink.Name = "JVLink"
        Me.JVLink.OcxState = CType(resources.GetObject("JVLink.OcxState"), System.Windows.Forms.AxHost.State)
        Me.JVLink.Size = New System.Drawing.Size(88, 40)
        Me.JVLink.TabIndex = 9
        '
        'frmMenu
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 12)
        Me.ClientSize = New System.Drawing.Size(200, 309)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.JVLink, Me.barReadSize, Me.btnStopJVData, Me.btnSettingJVLink, Me.btnViewDenmaList, Me.btnInitDB, Me.cmbYear, Me.barFileCount, Me.btnGetJVData})
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.Name = "frmMenu"
        Me.Text = "サンプルプログラム － メニュー"
        CType(Me.JVLink, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Declare Function GetActiveWindow Lib "USER32" () As Integer

    ' キャンセルフラグ
    Private bCancelFlag As Boolean

    ' カレントパス
    Dim strCurPath As String

    ' FromTime
    Dim strFromTime As String

    'エラーメッセージ
    Dim strErrMsg As String

    Private Sub btnGetJVData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGetJVData.Click

        Try
            ' リターンコード
            Dim lReturnCode As Long

            ' JVOpen用(DataSpec)
            Dim strDataSpec As String
            ' JVOpen用(Option)
            Dim lOption As Long
            ' JVOpen用(ReadCount)
            Dim lReadCount As Long
            ' JVOpen用(DownloadCount)
            Dim lDownloadCount As Long
            ' JVOpen用(LastFileTimestamp)
            Dim strLastFileTimestamp As String = String.Empty

            ' JVGets用(バッファポインタ)
            Dim szBuff(0) As Byte
            ' JVGets用(バッファ)
            Dim strBuff As String
            ' JVGets用(バッファサイズ)
            Dim lBuffSize As Long
            ' JVGets用(ファイル名)
            Dim strFileName As String = String.Empty

            ' データ区分
            Dim strRecID As String

            ' キャンセル検知フラグの初期化
            bCancelFlag = False

            Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

            ' プログレスバーの初期化
            Me.barFileCount.Value = 0
            Me.barReadSize.Value = 0

            ' JVOpenの呼び出し
            strDataSpec = "RACERCVN"
            lOption = "2"
            'strDataSpec = "RACE"
            'lOption = "4"
            lBuffSize = 110000
            lReturnCode = Me.JVLink.JVOpen(strDataSpec, strFromTime, lOption, lReadCount, lDownloadCount, strLastFileTimestamp)
            If lReturnCode = ST_READ_EOF Then
                Me.Cursor = System.Windows.Forms.Cursors.Default
                MsgBox("データベースは最新の状態です。")
                Exit Sub
            End If
            If lReturnCode < 0 Then
                Me.Cursor = System.Windows.Forms.Cursors.Default
                MsgBox("JVOpenエラー コード：" & lReturnCode & "：", MessageBoxIcon.Error)
                Exit Sub
            End If

            ' ボタンの抑止
            Me.btnGetJVData.Enabled = False
            Me.btnViewDenmaList.Enabled = False
            Me.btnInitDB.Enabled = False
            Me.btnSettingJVLink.Enabled = False
            Me.cmbYear.Enabled = False
            Me.btnStopJVData.Enabled = True

            ' 合計ファイル数のプログレスバーの上限を設定
            Me.barFileCount.Maximum = lReadCount

            ' プログレスバー用変数
            Dim lTotalFileCount As Long
            Dim lTotalReadSize As Long
            lTotalReadSize = 0
            lTotalFileCount = 0

            ' JVSkip制御フラグ
            Dim bSkipFlg As Boolean

            Do
                ' バックグラウンドでの処理
                System.Windows.Forms.Application.DoEvents()

                'キャンセルが押されたら処理(ループ)を抜ける
                If bCancelFlag = True Then Exit Do

                ' JVGetsの呼び出し
#Disable Warning BC41999
                lReturnCode = Me.JVLink.JVGets(szBuff, lBuffSize, strFileName)
#Enable Warning BC41999

                ' エラー判定
                Select Case lReturnCode

                    Case Is > ST_READ_SUCCESS
                        ' 正常

                        ' 文字コード変換(SJIS→UNICODE)
                        strBuff = System.Text.Encoding.GetEncoding(932).GetString(szBuff)

                        ' データ区分の取得
                        strRecID = strBuff.Substring(0, 2)

                        bSkipFlg = False

                        ' 処理対象データのみデータベースへ登録
                        If strRecID = ID_RACE Then
                            ' レース情詳細
                            ImportRA.Add(strBuff, lBuffSize)
                        ElseIf strRecID = ID_RACE_UMA Then
                            ' 馬毎レース情報
                            ImportSE.Add(strBuff, lBuffSize)
                        ElseIf strRecID = ID_UMA Then
                            ' 競走馬マスタ
                            ImportUM.Add(strBuff, lBuffSize)
                        Else
                            '対象外ファイルはスキップ(フラグを設定)
                            bSkipFlg = True
                        End If

                        If bSkipFlg = True Then
                            '対象外ファイルはスキップ
                            Me.JVLink.JVSkip()

                            ' カレントファイルのプログレスバーを更新
                            Me.barReadSize.Value = Me.barReadSize.Maximum

                            ' 合計ファイル数のプログレスバーを更新
                            lTotalFileCount = lTotalFileCount + 1
                            Me.barFileCount.Value = lTotalFileCount

                        Else
                            ' カレントファイルのプログレスバーを更新
                            Me.barReadSize.Maximum = Me.JVLink.m_CurrentReadFilesize
                            lTotalReadSize = lTotalReadSize + szBuff.Length - 1
                            Me.barReadSize.Value = lTotalReadSize
                        End If
                        ReDim szBuff(0)

                    Case ST_READ_EOF
                        ' ファイルの区切れ

                        ' 合計ファイル数のプログレスバーを更新
                        lTotalFileCount = lTotalFileCount + 1
                        Me.barFileCount.Value = lTotalFileCount

                        ' カレントファイルのプログレスバーを初期化
                        lTotalReadSize = 0

                        ' FromTimeを退避
                        strFromTime = Me.JVLink.m_CurrentFileTimeStamp

                    Case ST_READ_EOL
                        ' 全レコード読込み終了(EOF)
                        Exit Do

                    Case ST_READ_DOWNLOAD_NOW
                        ' ダウンロード中の場合、1秒スリープしダウンロード待ち
                        System.Threading.Thread.Sleep(1000)

                    Case Is <= ST_READ_ERR
                        ' エラー
                        MsgBox("JVGetsエラー コード：" & lReturnCode & "：", MessageBoxIcon.Error)
                        Exit Do

                End Select
            Loop

            ' FromTimeをiniファイルに保存
            WriteProfileDataStr("Setting", "FromTime", strFromTime, strCurPath)

            ' JVCloseの呼出
            lReturnCode = Me.JVLink.JVClose()
            If lReturnCode <> 0 Then
                MsgBox("JVCloseエラー コード：" & lReturnCode & "：", MessageBoxIcon.Error)
            End If

            If bCancelFlag = False Then
                MsgBox("開催情報の取得が終了しました。", MsgBoxStyle.ApplicationModal)
            Else
                MsgBox("開催情報の取得を中止しました。", MsgBoxStyle.ApplicationModal)
            End If

            ' 開催年月日選択コンボボックスの表示
            getRaceYMDList(Me.cmbYear)

            ' ボタンの抑止を解除
            Me.btnGetJVData.Enabled = True
            Me.btnViewDenmaList.Enabled = True
            Me.btnInitDB.Enabled = True
            Me.btnSettingJVLink.Enabled = True
            Me.cmbYear.Enabled = True
            Me.btnStopJVData.Enabled = False
            ' プログレスバーを元に戻す
            Me.barFileCount.Value = 0
            Me.barReadSize.Value = 0


        Catch
            Debug.WriteLine(Err.Description)
        End Try

        Me.Cursor = System.Windows.Forms.Cursors.Default

    End Sub

    Private Sub btnInitDB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInitDB.Click

        ' リターンコード
        Dim lReturnCode As Long

        lReturnCode = MsgBox("ＤＢを初期化しますか？", MsgBoxStyle.YesNo + MsgBoxStyle.Question, "確認")
        If lReturnCode = DialogResult.No Then
            Exit Sub
        End If

        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        ' プログレスバーの初期化
        Me.barReadSize.Value = 0
        Me.barFileCount.Value = 0

        Me.barFileCount.Maximum = 100

        'データベース登録クラスのクローズ
        If ImportRA Is Nothing = False Then
            ImportRA.Close()
            ImportRA = Nothing
        End If
        If ImportSE Is Nothing = False Then
            ImportSE.Close()
            ImportSE = Nothing
        End If
        If ImportUM Is Nothing = False Then
            ImportUM.Close()
            ImportUM = Nothing
        End If

        ' テーブルの全レコードをクリア
        gCon.Execute("DELETE FROM RACE")
        Me.barFileCount.Value = 30

        gCon.Execute("DELETE FROM UMA_RACE")
        Me.barFileCount.Value = 60

        gCon.Execute("DELETE FROM UMA")
        Me.barFileCount.Value = 90

        'データベース登録クラスの生成
        ImportRA = New clsImportRA()
        ImportSE = New clsImportSE()
        ImportUM = New clsImportUM()


        ' FromTimeを初期化しiniファイルに保存
        strFromTime = "00000000000000"
        WriteProfileDataStr("Setting", "FromTime", "00000000000000", strCurPath)

        ' 開催年月日選択コンボボックスのクリア
        getRaceYMDList(Me.cmbYear)

        Me.btnViewDenmaList.Enabled = False
        Me.cmbYear.Enabled = False

        Me.barFileCount.Value = 100

        Me.Cursor = System.Windows.Forms.Cursors.Default

        MsgBox("ＤＢの初期化が終了しました。")

    End Sub

    Private Sub frmMenu_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim lReturnCode As Long
        Dim iHWnd As Integer

        strCurPath = CurDir() + "\sample.ini"

        ' 設定ファイル存在チェック
        If Dir(strCurPath) = "" Then
            MsgBox("初期設定ファイル(sample.ini)が見つかりません")
            Exit Sub
        End If

        ' -----設定ファイルより、各種情報の読み込み-----
        ' DB接続文字列の取得
        strConnectString = GetProfileDataStr("Setting", "DBConnectString", strCurPath)
        If strConnectString = "" Then
            MsgBox("データベース接続文字列の取得に失敗しました。", MessageBoxIcon.Error)
            Exit Sub
        End If

        ' DBモードの取得
        Dim strDBMode As String
        strDBMode = GetProfileDataStr("Setting", "DBMode", strCurPath)
        If strDBMode = "" Then
            MsgBox("データベースモードの取得に失敗しました。", MessageBoxIcon.Error)
            Exit Sub
        End If
        If strDBMode = "0" Then
            SS = "["
            SE = "]"
        Else
            SS = ""
            SE = ""
        End If

        ' FROMTIME
        strFromTime = GetProfileDataStr("Setting", "FromTime", strCurPath)
        If strFromTime = "" Then
            strFromTime = "00000000000000"
        End If

        ' インスタンス生成
        objCDCv = New clsCodeConv()

        ' パスを指定し、コードファイルを読込む
        Dim strPath As String
        strPath = System.Reflection.Assembly.GetExecutingAssembly.Location()
        strPath = System.IO.Path.GetDirectoryName(strPath)
        objCDCv.FileName = strPath & "\CodeTable.csv"

        ' JVInitの呼び出し
        lReturnCode = Me.JVLink.JVInit("JVLinkSDKSampleAPP1")
        If lReturnCode <> 0 Then
            MsgBox("JVInitエラー コード：" & lReturnCode & "：", MessageBoxIcon.Error)
            Me.Cursor = System.Windows.Forms.Cursors.Default
            Exit Sub
        End If

        ' JV-Linkへのウィンドウハンドル登録
        iHWnd = GetActiveWindow()
        Me.JVLink.ParentHWnd = iHWnd

        ' データベースとの接続を行う。
        If ConnectDB() = True Then
            ' データベース登録クラスの生成
            ImportRA = New clsImportRA()
            ImportSE = New clsImportSE()
            ImportUM = New clsImportUM()

            ' 開催年月日選択コンボボックスの表示
            getRaceYMDList(Me.cmbYear)

            ' データベース関連機能ボタンを使用可に設定
            Me.btnGetJVData.Enabled = True
            Me.btnInitDB.Enabled = True

            If Me.cmbYear.Items.Count > 0 Then
                Me.btnViewDenmaList.Enabled = True
                Me.cmbYear.Enabled = True
            Else
                Me.btnViewDenmaList.Enabled = False
                Me.cmbYear.Enabled = False
            End If


        Else
            ' ADODBオブジェクトの開放
            gCon = Nothing

            ' 接続失敗時にデータベース関連機能ボタンを使用不可に設定
            Me.btnGetJVData.Enabled = False
            Me.btnViewDenmaList.Enabled = False
            Me.btnInitDB.Enabled = False

        End If

    End Sub

    Private Sub frmMenu_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed

        'データベース登録クラスのクローズ
        If ImportRA Is Nothing = False Then
            ImportRA.Close()
            ImportRA = Nothing
        End If
        If ImportSE Is Nothing = False Then
            ImportSE.Close()
            ImportSE = Nothing
        End If
        If ImportUM Is Nothing = False Then
            ImportUM.Close()
            ImportUM = Nothing
        End If

        'データベースとの切断を行う。
        If gCon Is Nothing = False Then
            gCon.Close()
            gCon = Nothing
        End If

        System.Diagnostics.Debug.WriteLine("gCon.Close")

    End Sub

    Private Sub btnViewDenmaList_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnViewDenmaList.Click

        Dim frmSubForm As New frmDenmaList()

        ' パラメータの設定
        frmSubForm.txtParam.Text = cmbYear.Text()

        'モードレスフォームとして表示
        frmSubForm.Show()

    End Sub

    Private Sub btnStopJVData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStopJVData.Click

        ' キャンセルフラグの設定
        bCancelFlag = True

    End Sub

    Private Sub btnSettingJVLink_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSettingJVLink.Click

        Try

            ' リターンコード
            Dim lReturnCode As Long

            ' 設定画面表示
            lReturnCode = JVLink.JVSetUIProperties()
            If lReturnCode <> 0 Then
                MsgBox("JVSetUIPropertiesエラー コード：" & lReturnCode & "：", MessageBoxIcon.Error)
            End If

        Catch
            Debug.WriteLine(Err.Description)
        End Try

    End Sub

    Public Function getRaceYMDList(ByVal cmbYMD As ComboBox) As Boolean
        On Error GoTo ErrorHandler

        Dim dbRS As ADODB.Recordset
        Dim dbFld As ADODB.Fields

        Dim strSQL As String
        'strSQL = "SELECT distinct Year, MonthDay FROM RACE ORDER BY Year desc, MonthDay desc"
        ' 地方・海外レース（データ区分"A","B"）を除外
        strSQL = "SELECT distinct Year, MonthDay FROM RACE WHERE not DataKubun in ('A','B') ORDER BY Year desc, MonthDay desc"

        ' コンボボックスのクリア
        cmbYMD.Text = ""
        cmbYMD.Items.Clear()

        ' レコードセットのオープン
        dbRS = New ADODB.Recordset()
        dbRS.Open(strSQL, gCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockPessimistic)

        While Not dbRS.EOF
            ' フィールドの取得
            dbFld = dbRS.Fields

            cmbYMD.Items.Add(dbFld("Year").Value() + dbFld("MonthDay").Value())

            dbRS.MoveNext()

        End While

        ' 直近日付を初期表示
        If cmbYMD.Items.Count > 0 Then
            cmbYMD.SelectedIndex() = 0
        End If

ExitHandler:
        ' レコードセットのクローズ
        dbRS.Close()
        dbRS = Nothing

        getRaceYMDList = True

        Exit Function

ErrorHandler:
        'System.Diagnostics.Debug.WriteLine(Err.Description)
        Resume ExitHandler

    End Function

End Class
