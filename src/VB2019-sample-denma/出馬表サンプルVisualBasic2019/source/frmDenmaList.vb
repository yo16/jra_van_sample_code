Option Explicit On 

Public Class frmDenmaList
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
    Friend WithEvents txtParam As System.Windows.Forms.TextBox
    Friend WithEvents grdDenmaList As AxMSFlexGridLib.AxMSFlexGrid
    Friend WithEvents TabDenmaList As System.Windows.Forms.TabControl
    Friend WithEvents lblDenmaList As System.Windows.Forms.Label
    Friend WithEvents TabKaisaiInfo As System.Windows.Forms.TabPage
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmDenmaList))
        Me.grdDenmaList = New AxMSFlexGridLib.AxMSFlexGrid()
        Me.txtParam = New System.Windows.Forms.TextBox()
        Me.TabDenmaList = New System.Windows.Forms.TabControl()
        Me.TabKaisaiInfo = New System.Windows.Forms.TabPage()
        Me.lblDenmaList = New System.Windows.Forms.Label()
        CType(Me.grdDenmaList, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabDenmaList.SuspendLayout()
        Me.SuspendLayout()
        '
        'grdDenmaList
        '
        Me.grdDenmaList.Location = New System.Drawing.Point(16, 72)
        Me.grdDenmaList.Name = "grdDenmaList"
        Me.grdDenmaList.OcxState = CType(resources.GetObject("grdDenmaList.OcxState"), System.Windows.Forms.AxHost.State)
        Me.grdDenmaList.Size = New System.Drawing.Size(848, 344)
        Me.grdDenmaList.TabIndex = 0
        '
        'txtParam
        '
        Me.txtParam.Enabled = False
        Me.txtParam.Location = New System.Drawing.Point(752, 48)
        Me.txtParam.Name = "txtParam"
        Me.txtParam.Size = New System.Drawing.Size(120, 19)
        Me.txtParam.TabIndex = 1
        Me.txtParam.Text = ""
        Me.txtParam.Visible = False
        '
        'TabDenmaList
        '
        Me.TabDenmaList.Controls.AddRange(New System.Windows.Forms.Control() {Me.TabKaisaiInfo})
        Me.TabDenmaList.Location = New System.Drawing.Point(8, 48)
        Me.TabDenmaList.Name = "TabDenmaList"
        Me.TabDenmaList.SelectedIndex = 0
        Me.TabDenmaList.Size = New System.Drawing.Size(864, 376)
        Me.TabDenmaList.TabIndex = 2
        '
        'TabKaisaiInfo
        '
        Me.TabKaisaiInfo.Location = New System.Drawing.Point(4, 21)
        Me.TabKaisaiInfo.Name = "TabKaisaiInfo"
        Me.TabKaisaiInfo.Size = New System.Drawing.Size(856, 351)
        Me.TabKaisaiInfo.TabIndex = 0
        Me.TabKaisaiInfo.Text = "開催情報"
        '
        'lblDenmaList
        '
        Me.lblDenmaList.BackColor = System.Drawing.SystemColors.ControlDark
        Me.lblDenmaList.Font = New System.Drawing.Font("MS UI Gothic", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblDenmaList.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.lblDenmaList.Location = New System.Drawing.Point(8, 8)
        Me.lblDenmaList.Name = "lblDenmaList"
        Me.lblDenmaList.Size = New System.Drawing.Size(864, 32)
        Me.lblDenmaList.TabIndex = 3
        Me.lblDenmaList.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.lblDenmaList.UseMnemonic = False
        '
        'frmDenmaList
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 12)
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(882, 431)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.txtParam, Me.lblDenmaList, Me.grdDenmaList, Me.TabDenmaList})
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.Name = "frmDenmaList"
        Me.Text = "サンプルプログラム － 出馬表選択"
        CType(Me.grdDenmaList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabDenmaList.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
    Dim structRA As JV_RA_RACE()
    Dim index(2, 11) As String

    Private Sub frmDenmaList_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        ' 開催年
        Dim strYYYY As String
        ' 開催月日
        Dim strMMDD As String

        ' RACEデータ取得SQL
        Dim strSQL_SELECT As String
        Dim strSQL_WHERE As String
        Dim strSQL_ORDER As String

        Dim strDenmaList As String

        ' 開催年の取得
        strYYYY = Me.txtParam.Text.Substring(0, 4)

        ' 開催月日の取得
        strMMDD = Me.txtParam.Text.Substring(4, 4)

        'SQL文字列の作成
        strSQL_SELECT = "SELECT * FROM RACE WHERE "

        strSQL_WHERE = SS + "Year" + SE + "='" + strYYYY + "' AND "
        strSQL_WHERE = strSQL_WHERE + SS + "MonthDay" + SE + "='" + strMMDD + "' AND "
        strSQL_WHERE = strSQL_WHERE + SS + "DataKubun" + SE + " not in ('A', 'B') "

        strSQL_ORDER = "ORDER BY " + SS + "JyoCD" + SE + " ASC, "
        strSQL_ORDER = strSQL_ORDER + SS + "RaceNum" + SE + " ASC "

        structRA = ImportRA.SelectDB(strSQL_SELECT + strSQL_WHERE + strSQL_ORDER)

        Me.grdDenmaList.Cols = 3
        Me.grdDenmaList.Rows = 13
        Me.grdDenmaList.set_ColWidth(-1, 4200)
        Me.grdDenmaList.set_RowHeight(-1, 400)
        Me.grdDenmaList.set_RowHeight(0, 200)

        Dim iCol As Integer ' 列番号
        iCol = 0
        Dim iRaceNum As Integer ' レース番号(行番号)
        Dim iLoopCnt As Integer  'ループカウンタ
        Dim iTmp1 As Integer
        Dim iTmp2 As Integer
        Dim strTmp As String

        '' ラベルの表示
        Dim strDenmaListdate As String
        strDenmaListdate = "  " & txtParam.Text.Substring(0, 4) & "年" & txtParam.Text.Substring(4, 2) & "月" & txtParam.Text.Substring(6, 2) & "日(" & objCDCv.GetCodeName(CV_WD_CD, structRA(0).RaceInfo.YoubiCD, 2) & ")"
        Me.lblDenmaList.Text = strDenmaListdate

        ' 出馬表選択一覧の表示
        '
        For iLoopCnt = 0 To structRA.Length - 1
            ' タイトル行の表示
            ' 
            ' 競馬場コードが変わったら次の列をカレント列とする
            If iLoopCnt <> 0 Then
                If structRA(iLoopCnt).id.JyoCD.Equals(structRA(iLoopCnt - 1).id.JyoCD) = False Then
                    iCol = iCol + 1
                End If
            End If
            Me.grdDenmaList.Col = iCol
            ' 競馬場コードの変換
            strTmp = objCDCv.GetCodeName(CV_JO_CD, structRA(iLoopCnt).id.JyoCD, 4)
            ' 開催回、開催日数ゼロサプレス
            iTmp1 = structRA(iLoopCnt).id.Kaiji
            iTmp2 = structRA(iLoopCnt).id.Nichiji
            ' 表示
            Me.grdDenmaList.set_TextArray(iCol, strTmp & iTmp1.ToString.PadLeft(2) & "回" & iTmp2.ToString.PadLeft(2) & "日")

            iRaceNum = structRA(iLoopCnt).id.RaceNum
            ' 上から[レース番号]番目の行をカレント行とする
            Me.grdDenmaList.Row = iRaceNum
            ' [レース番号]
            strDenmaList = iRaceNum.ToString.PadLeft(4) & "R "
            ' [競走名略称６文字][重賞]
            strTmp = TrimSP(structRA(iLoopCnt).RaceInfo.Ryakusyo6) & GRAD3(structRA(iLoopCnt).GradeCD)
            strDenmaList = strDenmaList & bPadR(strTmp, 18) & " "
            ' [競走種別]
            strTmp = KSSB7(structRA(iLoopCnt).JyokenInfo.SyubetuCD)
            strDenmaList = strDenmaList & strTmp & " "
            ' [競争条件]
            strDenmaList = strDenmaList & KSJK4(structRA(iLoopCnt).JyokenInfo.JyokenCD(4))
            ' [改行]
            strDenmaList = strDenmaList & vbCrLf
            ' [発走時刻]
            iTmp1 = structRA(iLoopCnt).HassoTime.Substring(0, 2)
            strDenmaList = strDenmaList & iTmp1.ToString.PadLeft(2) & ":"
            strDenmaList = strDenmaList & structRA(iLoopCnt).HassoTime.Substring(2, 2) & " "
            ' [トラックコードの変換]
            strDenmaList = strDenmaList & bPadR(objCDCv.GetCodeName(CV_TR_CD, structRA(iLoopCnt).TrackCD, 2), 18) & " "
            ' [出走頭数](成績確定までは登録頭数)
            Select Case structRA(iLoopCnt).head.DataKubun
                Case KB_THU
                    iTmp1 = structRA(iLoopCnt).TorokuTosu
                Case KB_FRI
                    iTmp1 = structRA(iLoopCnt).TorokuTosu
                Case KB_S3
                    iTmp1 = structRA(iLoopCnt).TorokuTosu
                Case KB_S5
                    iTmp1 = structRA(iLoopCnt).TorokuTosu
                Case KB_SALL
                    iTmp1 = structRA(iLoopCnt).TorokuTosu
                Case KB_SCOR
                    iTmp1 = structRA(iLoopCnt).SyussoTosu
                Case KB_MON
                    iTmp1 = structRA(iLoopCnt).SyussoTosu
            End Select
            strDenmaList = strDenmaList & structRA(iLoopCnt).Kyori & "m " & iTmp1.ToString.PadLeft(2) & "頭"
            ' 表示
            Me.grdDenmaList.Text = strDenmaList
            ' index(次画面移行の際に渡すパラメータ)保持
            index(iCol, iRaceNum - 1) = txtParam.Text & structRA(iLoopCnt).id.JyoCD & structRA(iLoopCnt).id.RaceNum
        Next iLoopCnt

    End Sub


    Private Sub grdDenmaList_DblClickEvent(ByVal sender As Object, ByVal e As System.EventArgs) Handles grdDenmaList.DblClick
        Dim frmSubForm As New frmRaceInfo()

        Dim iCol As Integer
        Dim iRow As Integer

        ' 選択されたグリッドの列、行を取得
        iCol = Me.grdDenmaList.Col
        iRow = Me.grdDenmaList.Row

        ' グリッドが空でない場合、次のフォームを開く
        If Me.grdDenmaList.get_TextMatrix(iRow, iCol).Length <> 0 Then
            frmSubForm.txtParam.Text = index(iCol, iRow - 1)
            'モードレスフォームとして表示
            frmSubForm.Show()
        End If

    End Sub

End Class
