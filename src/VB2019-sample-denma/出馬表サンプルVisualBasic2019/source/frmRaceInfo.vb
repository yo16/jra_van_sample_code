Option Explicit On 

Public Class frmRaceInfo
    Inherits System.Windows.Forms.Form

#Region " Windows �t�H�[�� �f�U�C�i�Ő������ꂽ�R�[�h "

    Public Sub New()
        MyBase.New()

        ' ���̌Ăяo���� Windows �t�H�[�� �f�U�C�i�ŕK�v�ł��B
        InitializeComponent()

        ' InitializeComponent() �Ăяo���̌�ɏ�������ǉ����܂��B

    End Sub

    ' Form �� dispose ���I�[�o�[���C�h���ăR���|�[�l���g�ꗗ���������܂��B
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    ' Windows �t�H�[�� �f�U�C�i�ŕK�v�ł��B
    Private components As System.ComponentModel.IContainer

    ' ���� : �ȉ��̃v���V�[�W���́AWindows �t�H�[�� �f�U�C�i�ŕK�v�ł��B
    ' Windows �t�H�[�� �f�U�C�i���g���ĕύX���Ă��������B  
    ' �R�[�h �G�f�B�^�͎g�p���Ȃ��ł��������B
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
        Me.lblRaceInfo2.Font = New System.Drawing.Font("�l�r �S�V�b�N", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
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
        Me.TabDenmaList1.Text = "��{���"
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
        Me.lblRaceInfo3.Font = New System.Drawing.Font("�l�r �S�V�b�N", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
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
        Me.Text = "�T���v���v���O���� �| �o�n�\"
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

        ' �J�ÔN
        Dim strYYYY As String
        ' �J�Ì���
        Dim strMMDD As String
        ' ���n��R�[�h
        Dim strJyo As String
        ' ���[�X�ԍ�
        Dim strRaceNum As String

        ' RACE�f�[�^�擾SQL
        Dim strSQL_SELECT As String
        Dim strSQL_SELECT_SE As String
        Dim strSQL_SELECT_UM As String
        Dim strSQL_WHERE As String
        Dim strSQL_WHERE_UM As String
        Dim strSQL_ORDER As String
        Dim strKettoNum As String
        Dim iLoopCnt1 As Integer ' ���[�v�J�E���^
        Dim iLoopCnt2 As Integer ' ���[�v�J�E���^

        ' �J�ÔN�̎擾
        strYYYY = Me.txtParam.Text.Substring(0, 4)

        ' �J�Ì����̎擾
        strMMDD = Me.txtParam.Text.Substring(4, 4)

        ' ���n��R�[�h�̎擾
        strJyo = Me.txtParam.Text.Substring(8, 2)

        ' ���[�X�ԍ��̎擾
        strRaceNum = Me.txtParam.Text.Substring(10, 2)

        'SQL������̍쐬
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

        ' �o���n���\���o�n�\�ł��Ȃ��Ȃ����n�����O
        If structRA(0).head.DataKubun.Equals("KB_THU") = False Then
            strSQL_WHERE = strSQL_WHERE + "AND " + SS + "Umaban" + SE + "<>'00' "
        End If
        structSE = ImportSE.SelectDB(strSQL_SELECT_SE + strSQL_WHERE + strSQL_ORDER)

        ' SE�����݂���ꍇ�ASE�̌����o�^�ԍ��ɑΉ�����UM�������Ă���
        ' �ꌏ�����݂��Ȃ��ꍇ�A���b�Z�[�W��\����Close����
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


        '' ���x���\���i��A���[�X�ԍ��j
        '
        ' ���n��R�[�h�̕ϊ�
        strTmp1 = " " & objCDCv.GetCodeName(CV_JO_CD, structRA(0).id.JyoCD, 4)
        ' ���[�X�ԍ��𕶎���Ɋi�[
        iTmp1 = structRA(0).id.RaceNum
        iTmp2 = structRA(0).RaceInfo.Nkai
        ' ��A���[�X�ԍ��A�{��i�{�d�܂̏ꍇ�͉񎟁A�O���[�h�j
        strTmp1 = strTmp1 & iTmp1 & "R"
        If iTmp2 <> 0 Then
            strTmp1 = strTmp1 & " ��" & iTmp2 & "�� " & TrimSP(structRA(0).RaceInfo.Hondai) & GRAD2(structRA(0).GradeCD)
        Else
            If TrimSP(structRA(0).RaceInfo.Hondai).Equals("") = False Then
                strTmp1 = strTmp1 & " " & TrimSP(structRA(0).RaceInfo.Hondai)
            End If
        End If
        ' �\��
        Me.lblRaceInfo1.Text = strTmp1

        '' ���x���\���i���[�X�ڍׁj
        '
        strTmp1 = structRA(0).id.Year & structRA(0).id.MonthDay
        ' [�N����]�A[�j��]
        iTmp1 = strTmp1.Substring(4, 2)
        iTmp2 = strTmp1.Substring(6, 2)
        strTmp2 = " " & strTmp1.Substring(0, 4) & "/" & iTmp1.ToString.PadLeft(2) & "/" & iTmp2.ToString.PadLeft(2) & "(" & objCDCv.GetCodeName(CV_WD_CD, structRA(0).RaceInfo.YoubiCD, 2) & ")"
        ' [��������]
        iTmp1 = structRA(0).HassoTime.Substring(0, 2)
        strTmp2 = strTmp2 & "  ���� " & iTmp1.ToString.PadLeft(2) & ":" & structRA(0).HassoTime.Substring(2, 2) & " "
        ' [�������]
        strTmp2 = strTmp2 & objCDCv.GetCodeName(CV_RS_CD, structRA(0).JyokenInfo.SyubetuCD, 3) & " "
        ' [��������]
        strTmp2 = strTmp2 & objCDCv.GetCodeName(CV_RJ_CD, structRA(0).JyokenInfo.JyokenCD(4), 1) & " "
        ' [�����L��]
        strTmp2 = bPadR(strTmp2, 58) & objCDCv.GetCodeName(CV_RK_CD, structRA(0).JyokenInfo.KigoCD, 1) & "   "
        ' [�d�ʎ��]�A[���s]
        strTmp2 = strTmp2 & objCDCv.GetCodeName(CV_WH_CD, structRA(0).JyokenInfo.JyuryoCD, 1) & vbCrLf
        ' [�R�[�X�敪]
        strTmp2 = strTmp2 & Space(17)
        If structRA(0).CourseKubunCD.Equals("  ") Then
            strTmp2 = strTmp2 & structRA(0).CourseKubunCD & "        "
        Else
            strTmp2 = strTmp2 & structRA(0).CourseKubunCD & "�R�[�X  "
        End If
        ' [�g���b�N�R�[�h]�A[����]
        strTmp1 = objCDCv.GetCodeName(CV_TR_CD, structRA(0).TrackCD, 2) & structRA(0).Kyori & "m "
        strTmp2 = strTmp2 & bPadR(strTmp1, 16)
        ' [�o������]/[�o�^����]
        Select Case structRA(0).head.DataKubun
            Case KB_THU
                iTmp1 = structRA(0).TorokuTosu
                strTmp1 = "�o�^���� "
            Case KB_FRI
                iTmp1 = structRA(0).TorokuTosu
                strTmp1 = "�o�^���� "
            Case KB_S3
                iTmp1 = structRA(0).TorokuTosu
                strTmp1 = "�o�^���� "
            Case KB_S5
                iTmp1 = structRA(0).TorokuTosu
                strTmp1 = "�o�^���� "
            Case KB_SALL
                iTmp1 = structRA(0).TorokuTosu
                strTmp1 = "�o�^���� "
            Case KB_SCOR
                iTmp1 = structRA(0).SyussoTosu
                strTmp1 = "�o������ "
            Case KB_MON
                iTmp1 = structRA(0).SyussoTosu
                strTmp1 = "�o������ "
        End Select
        strTmp2 = strTmp2 & strTmp1 & iTmp1.ToString.PadLeft(2) & "��  "
        ' [�{�܋�]�i1���`5���j�A[���s]
        strTmp2 = strTmp2 & "�{�܋�   "
        For iLoopCnt1 = 0 To 4
            iTmp1 = structRA(0).Honsyokin(iLoopCnt1).Substring(0, 6)
            strTmp2 = strTmp2 & iTmp1.ToString.PadLeft(6)
        Next iLoopCnt1
        strTmp2 = strTmp2 & " ���~" & vbCrLf
        ' [�V��]
        strTmp1 = bPadR(objCDCv.GetCodeName(CV_WE_CD, structRA(0).TenkoBaba.TenkoCD, 1), 5)
        strTmp2 = strTmp2 & Space(17) & strTmp1
        ' [�n����]
        If structRA(0).TenkoBaba.SibaBabaCD.Equals("0") Then
            strTmp1 = ""
        Else
            strTmp1 = "��:" & objCDCv.GetCodeName(CV_BC_CD, structRA(0).TenkoBaba.SibaBabaCD, 1) & " "
        End If
        If structRA(0).TenkoBaba.DirtBabaCD.Equals("0") Then
            strTmp1 = strTmp1 & ""
        Else
            strTmp1 = strTmp1 & "�_�[�g:" & objCDCv.GetCodeName(CV_BC_CD, structRA(0).TenkoBaba.DirtBabaCD, 1)
        End If
        strTmp1 = bPadR(strTmp1, 21)
        strTmp2 = strTmp2 & strTmp1
        ' [��������]
        iTmp1 = structRA(0).NyusenTosu
        If iTmp1 = 0 Then
            strTmp2 = strTmp2 & Space(15)
        Else
            strTmp2 = strTmp2 & "�������� " & iTmp1.ToString.PadLeft(2) & "��  "
        End If
        ' [�t���܋�]�i1���`3���j
        If structRA(0).Fukasyokin(0).Equals("00000000") = False Then
            strTmp1 = "�t���܋� "
            For iLoopCnt1 = 0 To 2
                iTmp1 = structRA(0).Fukasyokin(iLoopCnt1).Substring(0, 6)
                iTmp2 = structRA(0).Fukasyokin(iLoopCnt1).Substring(6, 1)
                If iTmp2 = 0 Then
                    strTmp1 = strTmp1 & iTmp1.ToString.PadLeft(6)
                Else
                    strTmp1 = strTmp1 & (iTmp1.ToString & "." & structRA(0).Fukasyokin(iLoopCnt1).Substring(6, 1)).PadLeft(6)
                End If
            Next iLoopCnt1
            strTmp1 = strTmp1 & " ���~"
            strTmp2 = strTmp2 & strTmp1
        End If
        ' �\��
        Me.lblRaceInfo2.Text = strTmp2

        '' ���x���\���i�f�[�^�쐬���j
        '
        ' [�f�[�^�쐬�N����]
        iTmp1 = structRA(0).head.MakeDate.Month
        iTmp2 = structRA(0).head.MakeDate.Day
        strTmp1 = structRA(0).head.MakeDate.Year & "/" & iTmp1.ToString.PadLeft(2) & "/" & iTmp2.ToString.PadLeft(2) & " �쐬�f�[�^"
        ' �\��
        Me.lblRaceInfo3.Text = strTmp1


        '' �O���b�h���\��
        '
        ' �s�E�񐔁A�����w��
        Me.grdDenmaList.Cols = 22
        Me.grdDenmaList.Rows = 1 + structSE.Length
        Me.grdDenmaList.set_RowHeight(-1, 220)

        ' �����̕\���ʒu�i1:���񂹁@7:�E�񂹁j
        Me.grdDenmaList.set_ColAlignment(3, 1)
        Me.grdDenmaList.set_ColAlignment(9, 7)
        Me.grdDenmaList.set_ColAlignment(10, 7)
        Me.grdDenmaList.set_ColAlignment(11, 7)
        Me.grdDenmaList.set_ColAlignment(12, 7)
        Me.grdDenmaList.set_ColAlignment(13, 7)
        Me.grdDenmaList.set_ColAlignment(14, 1)


        '�^�C�g���s�̕\��
        Dim strTitle() As String = {"�g", "��", "B", "�n�L��", "�n��", "����", "��", "�K", "�R��", "���S", "�n�̏d", "����", "�{�܋��݌v", "�����܋�", "�����t", "�n��", "���Y��", "������", "��s��", "������", "�Ǎ���", "���F"}
        For iLoopCnt1 = 0 To strTitle.Length - 1
            Me.grdDenmaList.set_TextArray(iLoopCnt1, strTitle(iLoopCnt1))
        Next iLoopCnt1

        ' �o�n�\�\��
        For iLoopCnt1 = 0 To structSE.Length - 1
            If structUM Is Nothing = False Then
                flg = False
                ' �����o�^�ԍ���肻�̔n�̋����n�}�X�^��T��
                For iLoopCnt2 = 0 To structUM.Length - 1
                    If structSE(iLoopCnt1).KettoNum.Equals(structUM(iLoopCnt2).KettoNum) Then
                        ' �Y���n�������ꍇ�A�t���O�𗧂āAiIndexUM�ɕێ�
                        flg = True
                        iIndexUM = iLoopCnt2
                        ' ����ʂɓn���p�����[�^
                        index = index & structSE(iLoopCnt1).KettoNum
                    Else
                        ' �Y���n��������܂ł�iIndexUM��-1
                        If flg = False Then
                            iIndexUM = -1
                        End If
                    End If
                Next iLoopCnt2
            Else
                ' UM�����݂��Ȃ��ꍇ��iIndexUM��-1
                iIndexUM = -1
            End If

            ' �J�����g�s
            Me.grdDenmaList.Row = iLoopCnt1 + 1
            ' �J�����g��
            iColIdx = 0
            ' �\��[�g��]
            Me.grdDenmaList.Col = iColIdx
            If structSE(iLoopCnt1).Wakuban.Equals("0") = False Then
                Me.grdDenmaList.Text = structSE(iLoopCnt1).Wakuban
                ' �g�Ԃɂ���ĐF����
                Me.grdDenmaList.CellBackColor = Color.FromArgb(CELBK1(structSE(iLoopCnt1).Wakuban))
                Me.grdDenmaList.CellForeColor = Color.FromName(CELFK(structSE(iLoopCnt1).Wakuban))
            End If
            iColIdx = iColIdx + 1
            ' �\��[�n��]
            Me.grdDenmaList.Col = iColIdx
            iTmp1 = structSE(iLoopCnt1).Umaban
            If iTmp1 <> 0 Then
                Me.grdDenmaList.Text = iTmp1
            End If
            iColIdx = iColIdx + 1
            ' �\��[�u�����J�[]
            Me.grdDenmaList.Col = iColIdx
            If structSE(iLoopCnt1).Blinker.Equals("1") Then
                strTmp1 = "B"
            Else
                strTmp1 = ""
            End If
            Me.grdDenmaList.Text = strTmp1
            iColIdx = iColIdx + 1
            ' �\��[�n�L��]
            Me.grdDenmaList.Col = iColIdx
            Me.grdDenmaList.Text = objCDCv.GetCodeName(CV_UK_CD, structSE(iLoopCnt1).UmaKigoCD, 1)
            iColIdx = iColIdx + 1
            ' �\��[�n��]
            Me.grdDenmaList.Col = iColIdx
            Me.grdDenmaList.Text = TrimSP(structSE(iLoopCnt1).Bamei)
            iColIdx = iColIdx + 1
            ' �\��[����][�n��]
            Me.grdDenmaList.Col = iColIdx
            iTmp1 = structSE(iLoopCnt1).Barei
            Me.grdDenmaList.Text = SEIB4(structSE(iLoopCnt1).SexCD) & iTmp1
            iColIdx = iColIdx + 1
            ' �\��[�ѐF]
            Me.grdDenmaList.Col = iColIdx
            Me.grdDenmaList.Text = objCDCv.GetCodeName(CV_FC_CD, structSE(iLoopCnt1).KeiroCD, 1)
            iColIdx = iColIdx + 1
            ' �\��[�R�茩�K���敪]
            Me.grdDenmaList.Col = iColIdx
            Me.grdDenmaList.Text = objCDCv.GetCodeName(CV_KM_CD, structSE(iLoopCnt1).MinaraiCD, 1)
            iColIdx = iColIdx + 1
            ' �\��[�R��]
            Me.grdDenmaList.Col = iColIdx
            Me.grdDenmaList.Text = structSE(iLoopCnt1).KisyuRyakusyo
            iColIdx = iColIdx + 1
            ' �\��[���S]
            Me.grdDenmaList.Col = iColIdx
            Me.grdDenmaList.Text = structSE(iLoopCnt1).Futan.Substring(0, 2) & "." & structSE(iLoopCnt1).Futan.Substring(2, 1)
            iColIdx = iColIdx + 1
            ' �\��[�n�̏d]
            Me.grdDenmaList.Col = iColIdx
            If TrimSP(structSE(iLoopCnt1).BaTaijyu).Length <> 0 Then
                iTmp1 = structSE(iLoopCnt1).BaTaijyu
                If iTmp1 <> 0 And iTmp1 <> 999 Then
                    Me.grdDenmaList.Text = iTmp1.ToString & "kg"
                End If
            End If
            iColIdx = iColIdx + 1
            ' �\��[����]
            Me.grdDenmaList.Col = iColIdx
            If structSE(iLoopCnt1).ZogenFugo.Equals(" ") Then
                Select Case structSE(iLoopCnt1).ZogenSa
                    Case "000"
                        strTmp1 = "�}0"
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
            ' �\��[�{�܋��݌v]
            Me.grdDenmaList.Col = iColIdx
            If iIndexUM <> -1 Then ' structUM�����݂���ꍇ
                iTmp1 = structUM(iIndexUM).RuikeiHonsyoHeiti & "00"
                If iTmp1 <> 0 Then
                    Me.grdDenmaList.Text = Format(iTmp1, "#,#")
                Else
                    Me.grdDenmaList.Text = iTmp1
                End If
            End If
            iColIdx = iColIdx + 1
            ' �\��[�����܋��݌v]
            Me.grdDenmaList.Col = iColIdx
            If iIndexUM <> -1 Then ' structUM�����݂���ꍇ
                iTmp1 = structUM(iIndexUM).RuikeiSyutokuHeichi & "00"
                If iTmp1 <> 0 Then
                    Me.grdDenmaList.Text = Format(iTmp1, "#,#")
                Else
                    Me.grdDenmaList.Text = iTmp1
                End If
            End If
            iColIdx = iColIdx + 1
            ' �\��[�����t]
            Me.grdDenmaList.Col = iColIdx
            Me.grdDenmaList.Text = structSE(iLoopCnt1).ChokyosiRyakusyo
            iColIdx = iColIdx + 1
            ' �\��[�n��]
            Me.grdDenmaList.Col = iColIdx
            Me.grdDenmaList.Text = TrimSP(structSE(iLoopCnt1).BanusiName)
            iColIdx = iColIdx + 1
            ' �\��[���Y��]
            Me.grdDenmaList.Col = iColIdx
            If iIndexUM <> -1 Then ' structUM�����݂���ꍇ
                Me.grdDenmaList.Text = TrimSP(structUM(iIndexUM).BreederName)
            End If
            iColIdx = iColIdx + 1
            ' �\��[������]
            Me.grdDenmaList.Col = iColIdx
            If iIndexUM <> -1 Then ' structUM�����݂���ꍇ
                iTmp1 = structUM(iIndexUM).Kyakusitu(0)
                Me.grdDenmaList.Text = iTmp1
            End If
            iColIdx = iColIdx + 1
            ' �\��[��s��]
            Me.grdDenmaList.Col = iColIdx
            If iIndexUM <> -1 Then ' structUM�����݂���ꍇ
                iTmp1 = structUM(iIndexUM).Kyakusitu(1)
                Me.grdDenmaList.Text = iTmp1
            End If
            iColIdx = iColIdx + 1
            ' �\��[������]
            Me.grdDenmaList.Col = iColIdx
            If iIndexUM <> -1 Then ' structUM�����݂���ꍇ
                iTmp1 = structUM(iIndexUM).Kyakusitu(2)
                Me.grdDenmaList.Text = iTmp1
            End If
            iColIdx = iColIdx + 1
            ' �\��[�Ǎ���]
            Me.grdDenmaList.Col = iColIdx
            If iIndexUM <> -1 Then ' structUM�����݂���ꍇ
                iTmp1 = structUM(iIndexUM).Kyakusitu(3)
                Me.grdDenmaList.Text = iTmp1
            End If
            iColIdx = iColIdx + 1
            ' �\��[���F]
            Me.grdDenmaList.Col = iColIdx
            Me.grdDenmaList.Text = TrimSP(structSE(iLoopCnt1).Fukusyoku)

        Next iLoopCnt1


        '' �Z�����̌���
        ' 
        ' ����ێ�����z��
        Dim strWidth(Me.grdDenmaList.Cols - 1) As Integer
        ' ��P�ʂŃ��[�v
        For iLoopCnt1 = 0 To strWidth.Length - 1
            Me.grdDenmaList.Col = iLoopCnt1
            iTmp1 = 0
            iTmp2 = 0
            ' ��s������
            For iLoopCnt2 = 0 To structSE.Length
                Me.grdDenmaList.Row = iLoopCnt2
                iTmp1 = Str2Byte(Me.grdDenmaList.get_TextMatrix(iLoopCnt2, iLoopCnt1)).Length
                ' ���̗�̍ő啝(byte�P��)��strWidth�Ɋi�[
                If iTmp1 > iTmp2 Then
                    strWidth(iLoopCnt1) = iTmp1
                    iTmp2 = iTmp1
                End If
            Next iLoopCnt2
        Next iLoopCnt1

        ' strWidth�Ɋi�[���ꂽ�������ɃO���b�h�̃Z�������w��
        For iLoopCnt1 = 0 To strWidth.Length - 1
            Me.grdDenmaList.set_ColWidth(iLoopCnt1, 100 + strWidth(iLoopCnt1) * 100)
        Next iLoopCnt1

ExitHandler:
        Exit Sub

ErrorHandler:
        Me.Close()
        MsgBox("�Y���f�[�^�͖��擾�ł�", MsgBoxStyle.Information)
        Exit Sub

    End Sub

    Private Sub grdDenmaList_DblClickEvent(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles grdDenmaList.DblClick
        Dim frmSubForm As New frmUmaProfile()

        Dim iCol As Integer
        Dim iRow As Integer

        ' �I�����ꂽ�O���b�h�̗�A�s���擾
        iCol = Me.grdDenmaList.Col
        iRow = Me.grdDenmaList.Row
        ' �O���b�h����łȂ��ꍇ�A���̃t�H�[�����J��
        If Me.grdDenmaList.get_TextMatrix(iRow, iCol).Length <> 0 Then
            frmSubForm.txtParam.Text = index.Substring((iRow - 1) * 10, 10)
            '���[�h���X�t�H�[���Ƃ��ĕ\��
            frmSubForm.Show()
        End If

    End Sub
End Class
