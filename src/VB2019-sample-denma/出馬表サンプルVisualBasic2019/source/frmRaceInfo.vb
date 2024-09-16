Option Explicit On 

Public Class frmRaceInfo
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
    Friend WithEvents lblRaceInfo1 As System.Windows.Forms.Label
    Friend WithEvents lblRaceInfo2 As System.Windows.Forms.Label
    Friend WithEvents txtParam As System.Windows.Forms.TextBox
    Friend WithEvents TabRaceInfo As System.Windows.Forms.TabControl
    Friend WithEvents TabDenmaList1 As System.Windows.Forms.TabPage
    Friend WithEvents grdDenmaList As AxMSFlexGridLib.AxMSFlexGrid
    Friend WithEvents lblRaceInfo3 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmRaceInfo))
        Me.lblRaceInfo1 = New System.Windows.Forms.Label()
        Me.lblRaceInfo2 = New System.Windows.Forms.Label()
        Me.txtParam = New System.Windows.Forms.TextBox()
        Me.TabRaceInfo = New System.Windows.Forms.TabControl()
        Me.TabDenmaList1 = New System.Windows.Forms.TabPage()
        Me.grdDenmaList = New AxMSFlexGridLib.AxMSFlexGrid()
        Me.lblRaceInfo3 = New System.Windows.Forms.Label()
        Me.TabRaceInfo.SuspendLayout()
        Me.TabDenmaList1.SuspendLayout()
        CType(Me.grdDenmaList, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblRaceInfo1
        '
        Me.lblRaceInfo1.BackColor = System.Drawing.SystemColors.ControlDark
        Me.lblRaceInfo1.Font = New System.Drawing.Font("MS UI Gothic", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblRaceInfo1.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.lblRaceInfo1.Location = New System.Drawing.Point(8, 8)
        Me.lblRaceInfo1.Name = "lblRaceInfo1"
        Me.lblRaceInfo1.Size = New System.Drawing.Size(768, 32)
        Me.lblRaceInfo1.TabIndex = 0
        Me.lblRaceInfo1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.lblRaceInfo1.UseMnemonic = False
        '
        'lblRaceInfo2
        '
        Me.lblRaceInfo2.Font = New System.Drawing.Font("ＭＳ ゴシック", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblRaceInfo2.Location = New System.Drawing.Point(8, 48)
        Me.lblRaceInfo2.Name = "lblRaceInfo2"
        Me.lblRaceInfo2.Size = New System.Drawing.Size(664, 40)
        Me.lblRaceInfo2.TabIndex = 1
        Me.lblRaceInfo2.UseMnemonic = False
        '
        'txtParam
        '
        Me.txtParam.Enabled = False
        Me.txtParam.Location = New System.Drawing.Point(680, 56)
        Me.txtParam.Name = "txtParam"
        Me.txtParam.TabIndex = 2
        Me.txtParam.Text = ""
        Me.txtParam.Visible = False
        '
        'TabRaceInfo
        '
        Me.TabRaceInfo.Controls.AddRange(New System.Windows.Forms.Control() {Me.TabDenmaList1})
        Me.TabRaceInfo.Location = New System.Drawing.Point(8, 88)
        Me.TabRaceInfo.Name = "TabRaceInfo"
        Me.TabRaceInfo.SelectedIndex = 0
        Me.TabRaceInfo.Size = New System.Drawing.Size(768, 336)
        Me.TabRaceInfo.TabIndex = 3
        '
        'TabDenmaList1
        '
        Me.TabDenmaList1.Controls.AddRange(New System.Windows.Forms.Control() {Me.grdDenmaList})
        Me.TabDenmaList1.Location = New System.Drawing.Point(4, 21)
        Me.TabDenmaList1.Name = "TabDenmaList1"
        Me.TabDenmaList1.Size = New System.Drawing.Size(760, 311)
        Me.TabDenmaList1.TabIndex = 0
        Me.TabDenmaList1.Text = "基本情報"
        '
        'grdDenmaList
        '
        Me.grdDenmaList.ContainingControl = Me
        Me.grdDenmaList.Name = "grdDenmaList"
        Me.grdDenmaList.OcxState = CType(resources.GetObject("grdDenmaList.OcxState"), System.Windows.Forms.AxHost.State)
        Me.grdDenmaList.Size = New System.Drawing.Size(760, 312)
        Me.grdDenmaList.TabIndex = 0
        '
        'lblRaceInfo3
        '
        Me.lblRaceInfo3.BackColor = System.Drawing.SystemColors.Control
        Me.lblRaceInfo3.Font = New System.Drawing.Font("ＭＳ ゴシック", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblRaceInfo3.Location = New System.Drawing.Point(576, 40)
        Me.lblRaceInfo3.Name = "lblRaceInfo3"
        Me.lblRaceInfo3.Size = New System.Drawing.Size(200, 16)
        Me.lblRaceInfo3.TabIndex = 6
        Me.lblRaceInfo3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblRaceInfo3.UseMnemonic = False
        '
        'frmRaceInfo
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 12)
        Me.ClientSize = New System.Drawing.Size(786, 431)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.lblRaceInfo3, Me.TabRaceInfo, Me.txtParam, Me.lblRaceInfo2, Me.lblRaceInfo1})
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.Name = "frmRaceInfo"
        Me.Text = "サンプルプログラム － 出馬表"
        Me.TabRaceInfo.ResumeLayout(False)
        Me.TabDenmaList1.ResumeLayout(False)
        CType(Me.grdDenmaList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region
    Dim structRA As JV_RA_RACE()
    Dim structSE As JV_SE_RACE_UMA()
    Dim structUM As JV_UM_UMA()
    Dim index As String

    Private Sub frmRaceInfo_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        ' 開催年
        Dim strYYYY As String
        ' 開催月日
        Dim strMMDD As String
        ' 競馬場コード
        Dim strJyo As String
        ' レース番号
        Dim strRaceNum As String

        ' RACEデータ取得SQL
        Dim strSQL_SELECT As String
        Dim strSQL_SELECT_SE As String
        Dim strSQL_SELECT_UM As String
        Dim strSQL_WHERE As String
        Dim strSQL_WHERE_UM As String
        Dim strSQL_ORDER As String
        Dim strKettoNum As String
        Dim iLoopCnt1 As Integer ' ループカウンタ
        Dim iLoopCnt2 As Integer ' ループカウンタ

        ' 開催年の取得
        strYYYY = Me.txtParam.Text.Substring(0, 4)

        ' 開催月日の取得
        strMMDD = Me.txtParam.Text.Substring(4, 4)

        ' 競馬場コードの取得
        strJyo = Me.txtParam.Text.Substring(8, 2)

        ' レース番号の取得
        strRaceNum = Me.txtParam.Text.Substring(10, 2)

        'SQL文字列の作成
        strSQL_SELECT = "SELECT * FROM RACE WHERE "
        strSQL_SELECT_SE = "SELECT * FROM UMA_RACE WHERE "
        strSQL_SELECT_UM = "SELECT * FROM UMA WHERE "

        strSQL_WHERE = SS + "Year" + SE + "='" + strYYYY + "' AND "
        strSQL_WHERE = strSQL_WHERE + SS + "MonthDay" + SE + "='" + strMMDD + "' AND "
        strSQL_WHERE = strSQL_WHERE + SS + "JyoCD" + SE + "='" + strJyo + "' AND "
        strSQL_WHERE = strSQL_WHERE + SS + "RaceNum" + SE + "='" + strRaceNum + "' "

        strSQL_ORDER = "ORDER BY " + SS + "Umaban" + SE + " ASC, "
        strSQL_ORDER = strSQL_ORDER + SS + "Bamei" + SE + " ASC "

        structRA = ImportRA.SelectDB(strSQL_SELECT + strSQL_WHERE)

        ' 出走馬名表→出馬表でいなくなった馬を除外
        If structRA(0).head.DataKubun.Equals("KB_THU") = False Then
            strSQL_WHERE = strSQL_WHERE + "AND " + SS + "Umaban" + SE + "<>'00' "
        End If
        structSE = ImportSE.SelectDB(strSQL_SELECT_SE + strSQL_WHERE + strSQL_ORDER)

        ' SEが存在する場合、SEの血統登録番号に対応したUMを持ってくる
        ' 一件も存在しない場合、メッセージを表示しCloseする
        If structSE Is Nothing = False Then
            strKettoNum = "'" & structSE(0).KettoNum & "'"
            For iLoopCnt1 = 1 To structSE.Length - 1
                strKettoNum = strKettoNum & ", '" & structSE(iLoopCnt1).KettoNum & "'"
            Next iLoopCnt1
            strSQL_WHERE_UM = SS + "KettoNum" + SE + " in (" + strKettoNum + ") "

            structUM = ImportUM.SelectDB(strSQL_SELECT_UM + strSQL_WHERE_UM)
        Else
            GoTo ErrorHandler
        End If

        Dim strTmp1 As String
        Dim strTmp2 As String
        Dim iTmp1 As Integer
        Dim iTmp2 As Integer
        Dim iColIdx As Integer
        Dim iIndexUM As Integer
        Dim flg As Boolean


        '' ラベル表示（場、レース番号）
        '
        ' 競馬場コードの変換
        strTmp1 = " " & objCDCv.GetCodeName(CV_JO_CD, structRA(0).id.JyoCD, 4)
        ' レース番号を文字列に格納
        iTmp1 = structRA(0).id.RaceNum
        iTmp2 = structRA(0).RaceInfo.Nkai
        ' 場、レース番号、本題（＋重賞の場合は回次、グレード）
        strTmp1 = strTmp1 & iTmp1 & "R"
        If iTmp2 <> 0 Then
            strTmp1 = strTmp1 & " 第" & iTmp2 & "回 " & TrimSP(structRA(0).RaceInfo.Hondai) & GRAD2(structRA(0).GradeCD)
        Else
            If TrimSP(structRA(0).RaceInfo.Hondai).Equals("") = False Then
                strTmp1 = strTmp1 & " " & TrimSP(structRA(0).RaceInfo.Hondai)
            End If
        End If
        ' 表示
        Me.lblRaceInfo1.Text = strTmp1

        '' ラベル表示（レース詳細）
        '
        strTmp1 = structRA(0).id.Year & structRA(0).id.MonthDay
        ' [年月日]、[曜日]
        iTmp1 = strTmp1.Substring(4, 2)
        iTmp2 = strTmp1.Substring(6, 2)
        strTmp2 = " " & strTmp1.Substring(0, 4) & "/" & iTmp1.ToString.PadLeft(2) & "/" & iTmp2.ToString.PadLeft(2) & "(" & objCDCv.GetCodeName(CV_WD_CD, structRA(0).RaceInfo.YoubiCD, 2) & ")"
        ' [発走時刻]
        iTmp1 = structRA(0).HassoTime.Substring(0, 2)
        strTmp2 = strTmp2 & "  発走 " & iTmp1.ToString.PadLeft(2) & ":" & structRA(0).HassoTime.Substring(2, 2) & " "
        ' [競走種別]
        strTmp2 = strTmp2 & objCDCv.GetCodeName(CV_RS_CD, structRA(0).JyokenInfo.SyubetuCD, 3) & " "
        ' [競争条件]
        strTmp2 = strTmp2 & objCDCv.GetCodeName(CV_RJ_CD, structRA(0).JyokenInfo.JyokenCD(4), 1) & " "
        ' [競走記号]
        strTmp2 = bPadR(strTmp2, 58) & objCDCv.GetCodeName(CV_RK_CD, structRA(0).JyokenInfo.KigoCD, 1) & "   "
        ' [重量種別]、[改行]
        strTmp2 = strTmp2 & objCDCv.GetCodeName(CV_WH_CD, structRA(0).JyokenInfo.JyuryoCD, 1) & vbCrLf
        ' [コース区分]
        strTmp2 = strTmp2 & Space(17)
        If structRA(0).CourseKubunCD.Equals("  ") Then
            strTmp2 = strTmp2 & structRA(0).CourseKubunCD & "        "
        Else
            strTmp2 = strTmp2 & structRA(0).CourseKubunCD & "コース  "
        End If
        ' [トラックコード]、[距離]
        strTmp1 = objCDCv.GetCodeName(CV_TR_CD, structRA(0).TrackCD, 2) & structRA(0).Kyori & "m "
        strTmp2 = strTmp2 & bPadR(strTmp1, 16)
        ' [出走頭数]/[登録頭数]
        Select Case structRA(0).head.DataKubun
            Case KB_THU
                iTmp1 = structRA(0).TorokuTosu
                strTmp1 = "登録頭数 "
            Case KB_FRI
                iTmp1 = structRA(0).TorokuTosu
                strTmp1 = "登録頭数 "
            Case KB_S3
                iTmp1 = structRA(0).TorokuTosu
                strTmp1 = "登録頭数 "
            Case KB_S5
                iTmp1 = structRA(0).TorokuTosu
                strTmp1 = "登録頭数 "
            Case KB_SALL
                iTmp1 = structRA(0).TorokuTosu
                strTmp1 = "登録頭数 "
            Case KB_SCOR
                iTmp1 = structRA(0).SyussoTosu
                strTmp1 = "出走頭数 "
            Case KB_MON
                iTmp1 = structRA(0).SyussoTosu
                strTmp1 = "出走頭数 "
        End Select
        strTmp2 = strTmp2 & strTmp1 & iTmp1.ToString.PadLeft(2) & "頭  "
        ' [本賞金]（1着～5着）、[改行]
        strTmp2 = strTmp2 & "本賞金   "
        For iLoopCnt1 = 0 To 4
            iTmp1 = structRA(0).Honsyokin(iLoopCnt1).Substring(0, 6)
            strTmp2 = strTmp2 & iTmp1.ToString.PadLeft(6)
        Next iLoopCnt1
        strTmp2 = strTmp2 & " 万円" & vbCrLf
        ' [天候]
        strTmp1 = bPadR(objCDCv.GetCodeName(CV_WE_CD, structRA(0).TenkoBaba.TenkoCD, 1), 5)
        strTmp2 = strTmp2 & Space(17) & strTmp1
        ' [馬場状態]
        If structRA(0).TenkoBaba.SibaBabaCD.Equals("0") Then
            strTmp1 = ""
        Else
            strTmp1 = "芝:" & objCDCv.GetCodeName(CV_BC_CD, structRA(0).TenkoBaba.SibaBabaCD, 1) & " "
        End If
        If structRA(0).TenkoBaba.DirtBabaCD.Equals("0") Then
            strTmp1 = strTmp1 & ""
        Else
            strTmp1 = strTmp1 & "ダート:" & objCDCv.GetCodeName(CV_BC_CD, structRA(0).TenkoBaba.DirtBabaCD, 1)
        End If
        strTmp1 = bPadR(strTmp1, 21)
        strTmp2 = strTmp2 & strTmp1
        ' [入線頭数]
        iTmp1 = structRA(0).NyusenTosu
        If iTmp1 = 0 Then
            strTmp2 = strTmp2 & Space(15)
        Else
            strTmp2 = strTmp2 & "入線頭数 " & iTmp1.ToString.PadLeft(2) & "頭  "
        End If
        ' [付加賞金]（1着～3着）
        If structRA(0).Fukasyokin(0).Equals("00000000") = False Then
            strTmp1 = "付加賞金 "
            For iLoopCnt1 = 0 To 2
                iTmp1 = structRA(0).Fukasyokin(iLoopCnt1).Substring(0, 6)
                iTmp2 = structRA(0).Fukasyokin(iLoopCnt1).Substring(6, 1)
                If iTmp2 = 0 Then
                    strTmp1 = strTmp1 & iTmp1.ToString.PadLeft(6)
                Else
                    strTmp1 = strTmp1 & (iTmp1.ToString & "." & structRA(0).Fukasyokin(iLoopCnt1).Substring(6, 1)).PadLeft(6)
                End If
            Next iLoopCnt1
            strTmp1 = strTmp1 & " 万円"
            strTmp2 = strTmp2 & strTmp1
        End If
        ' 表示
        Me.lblRaceInfo2.Text = strTmp2

        '' ラベル表示（データ作成日）
        '
        ' [データ作成年月日]
        iTmp1 = structRA(0).head.MakeDate.Month
        iTmp2 = structRA(0).head.MakeDate.Day
        strTmp1 = structRA(0).head.MakeDate.Year & "/" & iTmp1.ToString.PadLeft(2) & "/" & iTmp2.ToString.PadLeft(2) & " 作成データ"
        ' 表示
        Me.lblRaceInfo3.Text = strTmp1


        '' グリッド内表示
        '
        ' 行・列数、高さ指定
        Me.grdDenmaList.Cols = 22
        Me.grdDenmaList.Rows = 1 + structSE.Length
        Me.grdDenmaList.set_RowHeight(-1, 220)

        ' 文字の表示位置（1:左寄せ　7:右寄せ）
        Me.grdDenmaList.set_ColAlignment(3, 1)
        Me.grdDenmaList.set_ColAlignment(9, 7)
        Me.grdDenmaList.set_ColAlignment(10, 7)
        Me.grdDenmaList.set_ColAlignment(11, 7)
        Me.grdDenmaList.set_ColAlignment(12, 7)
        Me.grdDenmaList.set_ColAlignment(13, 7)
        Me.grdDenmaList.set_ColAlignment(14, 1)


        'タイトル行の表示
        Dim strTitle() As String = {"枠", "番", "B", "馬記号", "馬名", "性齢", "毛", "習", "騎手", "負担", "馬体重", "増減", "本賞金累計", "収得賞金", "調教師", "馬主", "生産者", "逃げ回数", "先行回数", "差し回数", "追込回数", "服色"}
        For iLoopCnt1 = 0 To strTitle.Length - 1
            Me.grdDenmaList.set_TextArray(iLoopCnt1, strTitle(iLoopCnt1))
        Next iLoopCnt1

        ' 出馬表表示
        For iLoopCnt1 = 0 To structSE.Length - 1
            If structUM Is Nothing = False Then
                flg = False
                ' 血統登録番号よりその馬の競走馬マスタを探す
                For iLoopCnt2 = 0 To structUM.Length - 1
                    If structSE(iLoopCnt1).KettoNum.Equals(structUM(iLoopCnt2).KettoNum) Then
                        ' 該当馬がいた場合、フラグを立て、iIndexUMに保持
                        flg = True
                        iIndexUM = iLoopCnt2
                        ' 次画面に渡すパラメータ
                        index = index & structSE(iLoopCnt1).KettoNum
                    Else
                        ' 該当馬が見つかるまではiIndexUMは-1
                        If flg = False Then
                            iIndexUM = -1
                        End If
                    End If
                Next iLoopCnt2
            Else
                ' UMが存在しない場合もiIndexUMは-1
                iIndexUM = -1
            End If

            ' カレント行
            Me.grdDenmaList.Row = iLoopCnt1 + 1
            ' カレント列
            iColIdx = 0
            ' 表示[枠番]
            Me.grdDenmaList.Col = iColIdx
            If structSE(iLoopCnt1).Wakuban.Equals("0") = False Then
                Me.grdDenmaList.Text = structSE(iLoopCnt1).Wakuban
                ' 枠番によって色分け
                Me.grdDenmaList.CellBackColor = Color.FromArgb(CELBK1(structSE(iLoopCnt1).Wakuban))
                Me.grdDenmaList.CellForeColor = Color.FromName(CELFK(structSE(iLoopCnt1).Wakuban))
            End If
            iColIdx = iColIdx + 1
            ' 表示[馬番]
            Me.grdDenmaList.Col = iColIdx
            iTmp1 = structSE(iLoopCnt1).Umaban
            If iTmp1 <> 0 Then
                Me.grdDenmaList.Text = iTmp1
            End If
            iColIdx = iColIdx + 1
            ' 表示[ブリンカー]
            Me.grdDenmaList.Col = iColIdx
            If structSE(iLoopCnt1).Blinker.Equals("1") Then
                strTmp1 = "B"
            Else
                strTmp1 = ""
            End If
            Me.grdDenmaList.Text = strTmp1
            iColIdx = iColIdx + 1
            ' 表示[馬記号]
            Me.grdDenmaList.Col = iColIdx
            Me.grdDenmaList.Text = objCDCv.GetCodeName(CV_UK_CD, structSE(iLoopCnt1).UmaKigoCD, 1)
            iColIdx = iColIdx + 1
            ' 表示[馬名]
            Me.grdDenmaList.Col = iColIdx
            Me.grdDenmaList.Text = TrimSP(structSE(iLoopCnt1).Bamei)
            iColIdx = iColIdx + 1
            ' 表示[性別][馬齢]
            Me.grdDenmaList.Col = iColIdx
            iTmp1 = structSE(iLoopCnt1).Barei
            Me.grdDenmaList.Text = SEIB4(structSE(iLoopCnt1).SexCD) & iTmp1
            iColIdx = iColIdx + 1
            ' 表示[毛色]
            Me.grdDenmaList.Col = iColIdx
            Me.grdDenmaList.Text = objCDCv.GetCodeName(CV_FC_CD, structSE(iLoopCnt1).KeiroCD, 1)
            iColIdx = iColIdx + 1
            ' 表示[騎手見習い区分]
            Me.grdDenmaList.Col = iColIdx
            Me.grdDenmaList.Text = objCDCv.GetCodeName(CV_KM_CD, structSE(iLoopCnt1).MinaraiCD, 1)
            iColIdx = iColIdx + 1
            ' 表示[騎手]
            Me.grdDenmaList.Col = iColIdx
            Me.grdDenmaList.Text = structSE(iLoopCnt1).KisyuRyakusyo
            iColIdx = iColIdx + 1
            ' 表示[負担]
            Me.grdDenmaList.Col = iColIdx
            Me.grdDenmaList.Text = structSE(iLoopCnt1).Futan.Substring(0, 2) & "." & structSE(iLoopCnt1).Futan.Substring(2, 1)
            iColIdx = iColIdx + 1
            ' 表示[馬体重]
            Me.grdDenmaList.Col = iColIdx
            If TrimSP(structSE(iLoopCnt1).BaTaijyu).Length <> 0 Then
                iTmp1 = structSE(iLoopCnt1).BaTaijyu
                If iTmp1 <> 0 And iTmp1 <> 999 Then
                    Me.grdDenmaList.Text = iTmp1.ToString & "kg"
                End If
            End If
            iColIdx = iColIdx + 1
            ' 表示[増減]
            Me.grdDenmaList.Col = iColIdx
            If structSE(iLoopCnt1).ZogenFugo.Equals(" ") Then
                Select Case structSE(iLoopCnt1).ZogenSa
                    Case "000"
                        strTmp1 = "±0"
                    Case "999"
                        strTmp1 = "----"
                    Case "   "
                        strTmp1 = "    "
                End Select
            Else
                iTmp1 = structSE(iLoopCnt1).ZogenSa
                strTmp1 = structSE(iLoopCnt1).ZogenFugo & iTmp1
            End If
            Me.grdDenmaList.Text = strTmp1
            iColIdx = iColIdx + 1
            ' 表示[本賞金累計]
            Me.grdDenmaList.Col = iColIdx
            If iIndexUM <> -1 Then ' structUMが存在する場合
                iTmp1 = structUM(iIndexUM).RuikeiHonsyoHeiti & "00"
                If iTmp1 <> 0 Then
                    Me.grdDenmaList.Text = Format(iTmp1, "#,#")
                Else
                    Me.grdDenmaList.Text = iTmp1
                End If
            End If
            iColIdx = iColIdx + 1
            ' 表示[収得賞金累計]
            Me.grdDenmaList.Col = iColIdx
            If iIndexUM <> -1 Then ' structUMが存在する場合
                iTmp1 = structUM(iIndexUM).RuikeiSyutokuHeichi & "00"
                If iTmp1 <> 0 Then
                    Me.grdDenmaList.Text = Format(iTmp1, "#,#")
                Else
                    Me.grdDenmaList.Text = iTmp1
                End If
            End If
            iColIdx = iColIdx + 1
            ' 表示[調教師]
            Me.grdDenmaList.Col = iColIdx
            Me.grdDenmaList.Text = structSE(iLoopCnt1).ChokyosiRyakusyo
            iColIdx = iColIdx + 1
            ' 表示[馬主]
            Me.grdDenmaList.Col = iColIdx
            Me.grdDenmaList.Text = TrimSP(structSE(iLoopCnt1).BanusiName)
            iColIdx = iColIdx + 1
            ' 表示[生産者]
            Me.grdDenmaList.Col = iColIdx
            If iIndexUM <> -1 Then ' structUMが存在する場合
                Me.grdDenmaList.Text = TrimSP(structUM(iIndexUM).BreederName)
            End If
            iColIdx = iColIdx + 1
            ' 表示[逃げ回数]
            Me.grdDenmaList.Col = iColIdx
            If iIndexUM <> -1 Then ' structUMが存在する場合
                iTmp1 = structUM(iIndexUM).Kyakusitu(0)
                Me.grdDenmaList.Text = iTmp1
            End If
            iColIdx = iColIdx + 1
            ' 表示[先行回数]
            Me.grdDenmaList.Col = iColIdx
            If iIndexUM <> -1 Then ' structUMが存在する場合
                iTmp1 = structUM(iIndexUM).Kyakusitu(1)
                Me.grdDenmaList.Text = iTmp1
            End If
            iColIdx = iColIdx + 1
            ' 表示[差し回数]
            Me.grdDenmaList.Col = iColIdx
            If iIndexUM <> -1 Then ' structUMが存在する場合
                iTmp1 = structUM(iIndexUM).Kyakusitu(2)
                Me.grdDenmaList.Text = iTmp1
            End If
            iColIdx = iColIdx + 1
            ' 表示[追込回数]
            Me.grdDenmaList.Col = iColIdx
            If iIndexUM <> -1 Then ' structUMが存在する場合
                iTmp1 = structUM(iIndexUM).Kyakusitu(3)
                Me.grdDenmaList.Text = iTmp1
            End If
            iColIdx = iColIdx + 1
            ' 表示[服色]
            Me.grdDenmaList.Col = iColIdx
            Me.grdDenmaList.Text = TrimSP(structSE(iLoopCnt1).Fukusyoku)

        Next iLoopCnt1


        '' セル幅の決定
        ' 
        ' 幅を保持する配列
        Dim strWidth(Me.grdDenmaList.Cols - 1) As Integer
        ' 列単位でループ
        For iLoopCnt1 = 0 To strWidth.Length - 1
            Me.grdDenmaList.Col = iLoopCnt1
            iTmp1 = 0
            iTmp2 = 0
            ' 一行ずつ検証
            For iLoopCnt2 = 0 To structSE.Length
                Me.grdDenmaList.Row = iLoopCnt2
                iTmp1 = Str2Byte(Me.grdDenmaList.get_TextMatrix(iLoopCnt2, iLoopCnt1)).Length
                ' その列の最大幅(byte単位)をstrWidthに格納
                If iTmp1 > iTmp2 Then
                    strWidth(iLoopCnt1) = iTmp1
                    iTmp2 = iTmp1
                End If
            Next iLoopCnt2
        Next iLoopCnt1

        ' strWidthに格納された幅を元にグリッドのセル幅を指定
        For iLoopCnt1 = 0 To strWidth.Length - 1
            Me.grdDenmaList.set_ColWidth(iLoopCnt1, 100 + strWidth(iLoopCnt1) * 100)
        Next iLoopCnt1

ExitHandler:
        Exit Sub

ErrorHandler:
        Me.Close()
        MsgBox("該当データは未取得です", MsgBoxStyle.Information)
        Exit Sub

    End Sub

    Private Sub grdDenmaList_DblClickEvent(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles grdDenmaList.DblClick
        Dim frmSubForm As New frmUmaProfile()

        Dim iCol As Integer
        Dim iRow As Integer

        ' 選択されたグリッドの列、行を取得
        iCol = Me.grdDenmaList.Col
        iRow = Me.grdDenmaList.Row
        ' グリッドが空でない場合、次のフォームを開く
        If Me.grdDenmaList.get_TextMatrix(iRow, iCol).Length <> 0 Then
            frmSubForm.txtParam.Text = index.Substring((iRow - 1) * 10, 10)
            'モードレスフォームとして表示
            frmSubForm.Show()
        End If

    End Sub
End Class
