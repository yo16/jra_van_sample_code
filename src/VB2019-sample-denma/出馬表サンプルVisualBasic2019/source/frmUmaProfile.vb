Option Explicit On 

Public Class frmUmaProfile
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
    Friend WithEvents txtParam As System.Windows.Forms.TextBox
    Friend WithEvents lblUmaProfile1 As System.Windows.Forms.Label
    Friend WithEvents lblUmaProfile2 As System.Windows.Forms.Label
    Friend WithEvents lblUmaProfile4 As System.Windows.Forms.Label
    Friend WithEvents lblUmaProfile3 As System.Windows.Forms.Label
    Friend WithEvents grdUmaProfile As AxMSFlexGridLib.AxMSFlexGrid
    Friend WithEvents TabUmaProfile As System.Windows.Forms.TabControl
    Friend WithEvents TabUmaProf As System.Windows.Forms.TabPage
    Friend WithEvents lblUmaProfile6 As System.Windows.Forms.Label
    Friend WithEvents lblUmaProfile5 As System.Windows.Forms.Label
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmUmaProfile))
        Me.lblUmaProfile1 = New System.Windows.Forms.Label()
        Me.lblUmaProfile2 = New System.Windows.Forms.Label()
        Me.lblUmaProfile4 = New System.Windows.Forms.Label()
        Me.txtParam = New System.Windows.Forms.TextBox()
        Me.lblUmaProfile3 = New System.Windows.Forms.Label()
        Me.grdUmaProfile = New AxMSFlexGridLib.AxMSFlexGrid()
        Me.TabUmaProfile = New System.Windows.Forms.TabControl()
        Me.TabUmaProf = New System.Windows.Forms.TabPage()
        Me.lblUmaProfile6 = New System.Windows.Forms.Label()
        Me.lblUmaProfile5 = New System.Windows.Forms.Label()
        CType(Me.grdUmaProfile, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabUmaProfile.SuspendLayout()
        Me.TabUmaProf.SuspendLayout()
        Me.SuspendLayout()
        '
        'lblUmaProfile1
        '
        Me.lblUmaProfile1.BackColor = System.Drawing.SystemColors.ControlDark
        Me.lblUmaProfile1.Font = New System.Drawing.Font("�l�r �S�V�b�N", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblUmaProfile1.ForeColor = System.Drawing.SystemColors.ControlLightLight
        Me.lblUmaProfile1.Location = New System.Drawing.Point(8, 8)
        Me.lblUmaProfile1.Name = "lblUmaProfile1"
        Me.lblUmaProfile1.Size = New System.Drawing.Size(768, 32)
        Me.lblUmaProfile1.TabIndex = 0
        Me.lblUmaProfile1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.lblUmaProfile1.UseMnemonic = False
        '
        'lblUmaProfile2
        '
        Me.lblUmaProfile2.BackColor = System.Drawing.SystemColors.Control
        Me.lblUmaProfile2.Font = New System.Drawing.Font("�l�r �S�V�b�N", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblUmaProfile2.Location = New System.Drawing.Point(8, 48)
        Me.lblUmaProfile2.Name = "lblUmaProfile2"
        Me.lblUmaProfile2.Size = New System.Drawing.Size(768, 56)
        Me.lblUmaProfile2.TabIndex = 1
        Me.lblUmaProfile2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.lblUmaProfile2.UseMnemonic = False
        '
        'lblUmaProfile4
        '
        Me.lblUmaProfile4.BackColor = System.Drawing.SystemColors.Control
        Me.lblUmaProfile4.Font = New System.Drawing.Font("�l�r �S�V�b�N", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblUmaProfile4.Location = New System.Drawing.Point(8, 112)
        Me.lblUmaProfile4.Name = "lblUmaProfile4"
        Me.lblUmaProfile4.Size = New System.Drawing.Size(416, 56)
        Me.lblUmaProfile4.TabIndex = 3
        Me.lblUmaProfile4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.lblUmaProfile4.UseMnemonic = False
        '
        'txtParam
        '
        Me.txtParam.Enabled = False
        Me.txtParam.Location = New System.Drawing.Point(680, 112)
        Me.txtParam.Name = "txtParam"
        Me.txtParam.TabIndex = 4
        Me.txtParam.Text = ""
        Me.txtParam.Visible = False
        '
        'lblUmaProfile3
        '
        Me.lblUmaProfile3.BackColor = System.Drawing.SystemColors.Control
        Me.lblUmaProfile3.Font = New System.Drawing.Font("�l�r �S�V�b�N", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblUmaProfile3.Location = New System.Drawing.Point(576, 48)
        Me.lblUmaProfile3.Name = "lblUmaProfile3"
        Me.lblUmaProfile3.Size = New System.Drawing.Size(200, 16)
        Me.lblUmaProfile3.TabIndex = 5
        Me.lblUmaProfile3.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.lblUmaProfile3.UseMnemonic = False
        '
        'grdUmaProfile
        '
        Me.grdUmaProfile.ContainingControl = Me
        Me.grdUmaProfile.Location = New System.Drawing.Point(0, 16)
        Me.grdUmaProfile.Name = "grdUmaProfile"
        Me.grdUmaProfile.OcxState = CType(resources.GetObject("grdUmaProfile.OcxState"), System.Windows.Forms.AxHost.State)
        Me.grdUmaProfile.Size = New System.Drawing.Size(760, 272)
        Me.grdUmaProfile.TabIndex = 0
        '
        'TabUmaProfile
        '
        Me.TabUmaProfile.Controls.AddRange(New System.Windows.Forms.Control() {Me.TabUmaProf})
        Me.TabUmaProfile.Location = New System.Drawing.Point(8, 176)
        Me.TabUmaProfile.Name = "TabUmaProfile"
        Me.TabUmaProfile.SelectedIndex = 0
        Me.TabUmaProfile.Size = New System.Drawing.Size(768, 312)
        Me.TabUmaProfile.TabIndex = 6
        '
        'TabUmaProf
        '
        Me.TabUmaProf.Controls.AddRange(New System.Windows.Forms.Control() {Me.lblUmaProfile6, Me.grdUmaProfile})
        Me.TabUmaProf.Location = New System.Drawing.Point(4, 21)
        Me.TabUmaProf.Name = "TabUmaProf"
        Me.TabUmaProf.Size = New System.Drawing.Size(760, 287)
        Me.TabUmaProf.TabIndex = 0
        Me.TabUmaProf.Text = "��������"
        '
        'lblUmaProfile6
        '
        Me.lblUmaProfile6.Font = New System.Drawing.Font("�l�r �S�V�b�N", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblUmaProfile6.Name = "lblUmaProfile6"
        Me.lblUmaProfile6.Size = New System.Drawing.Size(760, 16)
        Me.lblUmaProfile6.TabIndex = 1
        Me.lblUmaProfile6.Text = "��Q���[�X�ɂ��ẮA[��3�n����]��""��3F�^�C��""�łȂ��A""���Y���[�X���j�^�C����1F���σ^�C��""��\�����Ă��܂��B"
        Me.lblUmaProfile6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.lblUmaProfile6.UseMnemonic = False
        '
        'lblUmaProfile5
        '
        Me.lblUmaProfile5.BackColor = System.Drawing.SystemColors.Control
        Me.lblUmaProfile5.Font = New System.Drawing.Font("�l�r �S�V�b�N", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.lblUmaProfile5.Location = New System.Drawing.Point(424, 112)
        Me.lblUmaProfile5.Name = "lblUmaProfile5"
        Me.lblUmaProfile5.Size = New System.Drawing.Size(352, 64)
        Me.lblUmaProfile5.TabIndex = 7
        Me.lblUmaProfile5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.lblUmaProfile5.UseMnemonic = False
        '
        'frmUmaProfile
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 12)
        Me.ClientSize = New System.Drawing.Size(784, 495)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.txtParam, Me.lblUmaProfile5, Me.TabUmaProfile, Me.lblUmaProfile3, Me.lblUmaProfile4, Me.lblUmaProfile2, Me.lblUmaProfile1})
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.Name = "frmUmaProfile"
        Me.Text = "�T���v���v���O���� �| �����n�v���t�B�[��"
        CType(Me.grdUmaProfile, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabUmaProfile.ResumeLayout(False)
        Me.TabUmaProf.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
    Dim structRA As JV_RA_RACE()
    Dim structSE As JV_SE_RACE_UMA()
    Dim structUM As JV_UM_UMA()

    Private Sub frmUmaProfile_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' �����o�^�ԍ�
        Dim strKettoNum As String

        ' RACE�f�[�^�擾SQL
        Dim strSQL_SELECT As String
        Dim strSQL_SELECT_SE As String
        Dim strSQL_SELECT_UM As String
        Dim strSQL_WHERE As String
        Dim strSQL_WHERE_UM As String
        Dim strSQL_ORDER As String
        Dim iLoopCnt As Integer ' ���[�v�J�E���^
        Dim jLoopCnt As Integer ' ���[�v�J�E���^

        ' �����o�^�ԍ��̎擾
        strKettoNum = Me.txtParam.Text

        'SQL������̍쐬
        strSQL_SELECT = "SELECT * FROM RACE WHERE "
        strSQL_SELECT_SE = "SELECT * FROM UMA_RACE WHERE "
        strSQL_SELECT_UM = "SELECT * FROM UMA WHERE "

        strSQL_WHERE = SS + "KettoNum" + SE + "='" + strKettoNum + "' AND "
        strSQL_WHERE = strSQL_WHERE + SS + "DataKubun" + SE + " in ('7', '9', 'A', 'B') "
        strSQL_WHERE_UM = SS + "KettoNum" + SE + "='" + strKettoNum + "'"

        strSQL_ORDER = "ORDER BY " + SS + "Year" + SE + " DESC, "
        strSQL_ORDER = strSQL_ORDER + SS + "MonthDay" + SE + " DESC "

        ' �����o�^�ԍ���肻�̔n�̋����n�}�X�^�A�y�т��̔n�̑��������[�X�̔n�����[�X�����擾
        structUM = ImportUM.SelectDB(strSQL_SELECT_UM + strSQL_WHERE_UM)
        structSE = ImportSE.SelectDB(strSQL_SELECT_SE + strSQL_WHERE + strSQL_ORDER)

        If structSE Is Nothing = False Then
            Dim strRaceId(structSE.Length - 1) As String
            ' �N�����ꃌ�[�X�ԍ���ێ�
            For iLoopCnt = 0 To structSE.Length - 1
                strRaceId(iLoopCnt) = structSE(iLoopCnt).id.Year & structSE(iLoopCnt).id.MonthDay & structSE(iLoopCnt).id.JyoCD & structSE(iLoopCnt).id.RaceNum
            Next iLoopCnt

            ' ���[�X�ڍׂ��擾
            For iLoopCnt = 0 To structSE.Length - 1
                If iLoopCnt = 0 Then
                    strSQL_WHERE = SS + "Year" + SE + "='" + strRaceId(iLoopCnt).Substring(0, 4) + "' AND "
                Else
                    strSQL_WHERE = strSQL_WHERE + SS + "Year" + SE + "='" + strRaceId(iLoopCnt).Substring(0, 4) + "' AND "
                End If
                strSQL_WHERE = strSQL_WHERE + SS + "MonthDay" + SE + "='" + strRaceId(iLoopCnt).Substring(4, 4) + "' AND "
                strSQL_WHERE = strSQL_WHERE + SS + "JyoCD" + SE + "='" + strRaceId(iLoopCnt).Substring(8, 2) + "' AND "
                If iLoopCnt = structSE.Length - 1 Then
                    strSQL_WHERE = strSQL_WHERE + SS + "RaceNum" + SE + "='" + strRaceId(iLoopCnt).Substring(10, 2) + "' "
                Else
                    strSQL_WHERE = strSQL_WHERE + SS + "RaceNum" + SE + "='" + strRaceId(iLoopCnt).Substring(10, 2) + "' OR "
                End If
            Next iLoopCnt
            structRA = ImportRA.SelectDB(strSQL_SELECT + strSQL_WHERE)
        Else
            GoTo ErrorHandler
        End If

        Dim strTmp1 As String = String.Empty
        Dim strTmp2 As String
        Dim iTmp1 As Integer
        Dim iTmp2 As Integer
        Dim iColIdx As Integer
        Dim iIndexRA As Integer

        '' ���x���\���i�f�[�^�쐬���j
        '
        ' [�f�[�^�쐬�N����]
        iTmp1 = structUM(0).head.MakeDate.Month
        iTmp2 = structUM(0).head.MakeDate.Day
        strTmp1 = strTmp1 & structUM(0).head.MakeDate.Year & "/" & iTmp1.ToString.PadLeft(2) & "/" & iTmp2.ToString.PadLeft(2) & " �쐬�f�[�^"
        Me.lblUmaProfile3.Text = strTmp1

        '' ���x���\���i�n�L���A�n���j
        '
        ' �n�L���R�[�h�̕ϊ�
        strTmp1 = " " & objCDCv.GetCodeName(CV_UK_CD, structUM(0).UmaKigoCD, 1)
        ' �n���𕶎���Ɋi�[
        strTmp1 = strTmp1 & structUM(0).Bamei
        ' �\��
        Me.lblUmaProfile1.Text = strTmp1

        '' ���x���\���i����A�ѐF�A�i��A���N�����A�����n�o�^���A�����t�A�Y�n�A�n��A���Y�ҁj
        '
        ' [�n��]/[����]�R�[�h�̕ϊ�
        iTmp1 = structSE(0).Barei
        strTmp2 = " " & SEIB4(structUM(0).SexCD) & iTmp1 & "�� "
        ' [�ѐF]
        strTmp2 = strTmp2 & objCDCv.GetCodeName(CV_FC_CD, structUM(0).KeiroCD, 1) & " "
        ' [�i��]�R�[�h�̕ϊ�
        strTmp2 = strTmp2 & objCDCv.GetCodeName(CV_HS_CD, structUM(0).HinsyuCD, 2) & " "
        ' [���N����]
        iTmp1 = structUM(0).BirthDate.Month
        iTmp2 = structUM(0).BirthDate.Day
        strTmp2 = strTmp2 & structUM(0).BirthDate.Year & "�N" & iTmp1.ToString.PadLeft(2) & "��" & iTmp2.ToString.PadLeft(2) & "����" & vbCrLf
        ' [�����n�o�^�N����]
        iTmp1 = structUM(0).RegDate.Month
        iTmp2 = structUM(0).RegDate.Day
        strTmp2 = strTmp2 & " �����n�o�^: " & structUM(0).RegDate.Year & "�N" & iTmp1.ToString.PadLeft(2) & "��" & iTmp2.ToString.PadLeft(2) & "��  "
        ' [�����t]
        strTmp2 = strTmp2 & " �����t    : " & structUM(0).ChokyosiRyakusyo & vbCrLf
        ' [�Y�n]
        strTmp2 = strTmp2 & " �Y�n      : " & bPadR(TrimSP(structUM(0).SanchiName), 16)
        ' [�n��]
        strTmp2 = strTmp2 & " �n��      : " & structUM(0).BanusiName & vbCrLf
        ' [���Ғn��][���Y��]
        If TrimSP(structUM(0).Syotai).Length <> 0 Then
            strTmp2 = strTmp2 & " ���Ғn��  : " & bPadR(TrimSP(structUM(0).Syotai), 16) & " ���Y��    : " & structUM(0).BreederName & vbCrLf
        Else
            strTmp2 = strTmp2 & Space(29) & " ���Y��    : " & structUM(0).BreederName & vbCrLf
        End If
        ' �\��
        Me.lblUmaProfile2.Text = strTmp2

        '' ���x���\���i�܋��j
        '
        iTmp1 = structUM(0).RuikeiHonsyoHeiti & "00"
        strTmp1 = " ���n�{�܋�  : " & Format(iTmp1, "#,0").PadLeft(11) & "�~ "
        iTmp1 = structUM(0).RuikeiHonsyoSyogai & "00"
        strTmp1 = strTmp1 & " ��Q�{�܋�  : " & Format(iTmp1, "#,0").PadLeft(11) & "�~" & vbCrLf
        iTmp1 = structUM(0).RuikeiFukaHeichi & "00"
        strTmp1 = strTmp1 & " ���n�t���܋�: " & Format(iTmp1, "#,0").PadLeft(11) & "�~ "
        iTmp1 = structUM(0).RuikeiFukaSyogai & "00"
        strTmp1 = strTmp1 & " ��Q�t���܋�: " & Format(iTmp1, "#,0").PadLeft(11) & "�~" & vbCrLf
        iTmp1 = structUM(0).RuikeiSyutokuHeichi & "00"
        strTmp1 = strTmp1 & " ���n�����܋�: " & Format(iTmp1, "#,0").PadLeft(11) & "�~ "
        iTmp1 = structUM(0).RuikeiSyutokuSyogai & "00"
        strTmp1 = strTmp1 & " ��Q�����܋�: " & Format(iTmp1, "#,0").PadLeft(11) & "�~"
        ' �\��
        Me.lblUmaProfile4.Text = strTmp1

        '' ���x���\���i�r���j
        ' 
        iTmp1 = structUM(0).Kyakusitu(0)
        strTmp1 = " ������: " & iTmp1.ToString.PadLeft(3) & "��" & vbCrLf
        iTmp1 = structUM(0).Kyakusitu(1)
        strTmp1 = strTmp1 & " ��s��: " & iTmp1.ToString.PadLeft(3) & "��" & vbCrLf
        iTmp1 = structUM(0).Kyakusitu(2)
        strTmp1 = strTmp1 & " ������: " & iTmp1.ToString.PadLeft(3) & "��" & vbCrLf
        iTmp1 = structUM(0).Kyakusitu(3)
        strTmp1 = strTmp1 & " �Ǎ���: " & iTmp1.ToString.PadLeft(3) & "��"
        Me.lblUmaProfile5.Text = strTmp1


        '' �O���b�h���\��
        '
        ' �s�E�񐔁A�����w��
        Me.grdUmaProfile.Cols = 27
        Me.grdUmaProfile.Rows = 1 + structSE.Length
        Me.grdUmaProfile.set_RowHeight(-1, 220)

        ' �����̕\���ʒu�i1:���� 4:������ 7:�E�񂹁j
        Me.grdUmaProfile.set_ColAlignment(1, 4)
        Me.grdUmaProfile.set_ColAlignment(14, 4)
        Me.grdUmaProfile.set_ColAlignment(15, 4)
        Me.grdUmaProfile.set_ColAlignment(17, 7)
        Me.grdUmaProfile.set_ColAlignment(21, 7)
        Me.grdUmaProfile.set_ColAlignment(25, 1)

        '�^�C�g���s�̕\��
        Dim strTitle() As String = {"�J�Ó�", "����", "R", "���[�X��", "����", "�R�[�X", "�n��", "�K", "�R��", "���S", "B", "����", "�g��", "�n��", "�ُ�", "����", "�^�C��", "�R�[�i�[�ʉߏ�", "�P�I�b�Y", "�P�l�C", "�n�̏d", "������", "�l���{�܋�", "�l���t���܋�", "��3�n����", "1(2)���n", "�^�C����"}
        For iLoopCnt = 0 To strTitle.Length - 1
            Me.grdUmaProfile.set_TextArray(iLoopCnt, strTitle(iLoopCnt))
        Next iLoopCnt

        ' �ߋ����\��
        For iLoopCnt = 0 To structSE.Length - 1
            ' �J�ÔN�����A��A���[�X�ԍ���背�[�X�ڍׂ�T���AiIndexRA�ɕێ�
            For jLoopCnt = 0 To structRA.Length - 1
                strTmp1 = structSE(iLoopCnt).id.Year & structSE(iLoopCnt).id.MonthDay & structSE(iLoopCnt).id.JyoCD & structSE(iLoopCnt).id.RaceNum
                strTmp2 = structRA(jLoopCnt).id.Year & structRA(jLoopCnt).id.MonthDay & structRA(jLoopCnt).id.JyoCD & structRA(jLoopCnt).id.RaceNum
                If strTmp1.Equals(strTmp2) Then
                    iIndexRA = jLoopCnt
                End If
            Next jLoopCnt
            ' �J�����g�s
            Me.grdUmaProfile.Row = iLoopCnt + 1
            ' �J�����g��
            iColIdx = 0
            ' �\��[�J�Ó�]
            Me.grdUmaProfile.Col = iColIdx
            iTmp1 = structSE(iLoopCnt).id.MonthDay.Substring(0, 2)
            iTmp2 = structSE(iLoopCnt).id.MonthDay.Substring(2, 2)
            Me.grdUmaProfile.Text = structSE(iLoopCnt).id.Year & "/" & iTmp1.ToString.PadLeft(2) & "/" & iTmp2.ToString.PadLeft(2)
            iColIdx = iColIdx + 1
            ' �\��[�J��]
            Me.grdUmaProfile.Col = iColIdx
            iTmp1 = structSE(iLoopCnt).id.Kaiji
            If iTmp1 <> 0 Then
                strTmp1 = iTmp1.ToString.PadLeft(2)
            Else
                strTmp1 = "  "
            End If
            iTmp1 = structSE(iLoopCnt).id.Nichiji
            If iTmp1 <> 0 Then
                strTmp2 = iTmp1.ToString.PadLeft(2)
            Else
                strTmp2 = "  "
            End If
            Me.grdUmaProfile.Text = strTmp1 & objCDCv.GetCodeName(CV_JO_CD, structSE(iLoopCnt).id.JyoCD, 3) & strTmp2
            iColIdx = iColIdx + 1
            ' �\��[���[�X�ԍ�]
            Me.grdUmaProfile.Col = iColIdx
            iTmp1 = structSE(iLoopCnt).id.RaceNum
            Me.grdUmaProfile.Text = iTmp1
            iColIdx = iColIdx + 1
            ' �\��[���[�X��]
            Me.grdUmaProfile.Col = iColIdx
            Me.grdUmaProfile.Text = TrimSP(structRA(iIndexRA).RaceInfo.Ryakusyo6) & GRAD2(structRA(iIndexRA).GradeCD)
            iColIdx = iColIdx + 1
            ' �\��[��������]
            Me.grdUmaProfile.Col = iColIdx
            Me.grdUmaProfile.Text = KSSB6(structRA(iIndexRA).JyokenInfo.SyubetuCD) & KSJK4(structRA(iIndexRA).JyokenInfo.JyokenCD(4))
            iColIdx = iColIdx + 1
            ' �\��[�R�[�X]
            Me.grdUmaProfile.Col = iColIdx
            Me.grdUmaProfile.Text = TRCK5(structRA(iIndexRA).TrackCD) & structRA(iIndexRA).Kyori
            iColIdx = iColIdx + 1
            ' �\��[�n��]
            Me.grdUmaProfile.Col = iColIdx
            If structRA(iIndexRA).TenkoBaba.SibaBabaCD.Equals("0") Then
                strTmp1 = ""
            Else
                strTmp1 = "��" & BBJT4(structRA(iIndexRA).TenkoBaba.SibaBabaCD)
            End If
            If structRA(iIndexRA).TenkoBaba.DirtBabaCD.Equals("0") Then
                strTmp1 = strTmp1 & ""
            Else
                If TrimSP(strTmp1).Length <> 0 Then
                    strTmp1 = strTmp1 & ":�_" & BBJT4(structRA(iIndexRA).TenkoBaba.DirtBabaCD)
                Else
                    strTmp1 = strTmp1 & "�_" & BBJT4(structRA(iIndexRA).TenkoBaba.DirtBabaCD)
                End If
            End If
            Me.grdUmaProfile.Text = strTmp1
            iColIdx = iColIdx + 1
            ' �\��[�R�茩�K�敪]
            Me.grdUmaProfile.Col = iColIdx
            Me.grdUmaProfile.Text = objCDCv.GetCodeName(CV_KM_CD, structSE(iLoopCnt).MinaraiCD, 1)
            iColIdx = iColIdx + 1
            ' �\��[�R��]
            Me.grdUmaProfile.Col = iColIdx
            Me.grdUmaProfile.Text = structSE(iLoopCnt).KisyuRyakusyo
            iColIdx = iColIdx + 1
            ' �\��[���S]
            Me.grdUmaProfile.Col = iColIdx
            Me.grdUmaProfile.Text = structSE(iLoopCnt).Futan.Substring(0, 2) & "." & structSE(iLoopCnt).Futan.Substring(2, 1)
            iColIdx = iColIdx + 1
            ' �\��[�u�����J�[�敪]
            Me.grdUmaProfile.Col = iColIdx
            If structSE(iLoopCnt).Blinker.Equals("1") Then
                Me.grdUmaProfile.Text = "B"
            Else
                Me.grdUmaProfile.Text = ""
            End If
            iColIdx = iColIdx + 1
            ' �\��[����]
            Me.grdUmaProfile.Col = iColIdx
            iTmp1 = structRA(iIndexRA).SyussoTosu
            If iTmp1 <> 0 Then
                Me.grdUmaProfile.Text = iTmp1
            End If
            iColIdx = iColIdx + 1
            ' �\��[�g��]
            Me.grdUmaProfile.Col = iColIdx
            iTmp1 = structSE(iLoopCnt).Wakuban
            If iTmp1 <> 0 Then
                Me.grdUmaProfile.Text = iTmp1
            End If
            iColIdx = iColIdx + 1
            ' �\��[�n��]
            Me.grdUmaProfile.Col = iColIdx
            iTmp1 = structSE(iLoopCnt).Umaban
            If iTmp1 <> 0 Then
                Me.grdUmaProfile.Text = iTmp1
            End If
            iColIdx = iColIdx + 1
            ' �\��[�ُ�]
            Me.grdUmaProfile.Col = iColIdx
            Me.grdUmaProfile.Text = objCDCv.GetCodeName(CV_IR_CD, structSE(iLoopCnt).IJyoCD, 2)
            iColIdx = iColIdx + 1
            ' �\��[����]
            Me.grdUmaProfile.Col = iColIdx
            iTmp1 = structSE(iLoopCnt).KakuteiJyuni
            If iTmp1 <> 0 Then
                Me.grdUmaProfile.Text = iTmp1
                ' 1�`3���͐F����
                Me.grdUmaProfile.CellBackColor = Color.FromArgb(CELBK2(structSE(iLoopCnt).KakuteiJyuni))
            End If
            iColIdx = iColIdx + 1
            ' �\��[�^�C��]
            Me.grdUmaProfile.Col = iColIdx
            If structSE(iLoopCnt).Time.Equals("0000") = False Then
                Me.grdUmaProfile.Text = structSE(iLoopCnt).Time.Substring(0, 1) & ":" & structSE(iLoopCnt).Time.Substring(1, 2) & "." & structSE(iLoopCnt).Time.Substring(3, 1)
            End If
            iColIdx = iColIdx + 1
            '' �\��[�R�[�i�[�ʉߏ�]
            ' ��1�R�[�i�[�ʉߏ���
            Me.grdUmaProfile.Col = iColIdx
            iTmp1 = structSE(iLoopCnt).Jyuni1c
            If iTmp1 = 0 Then
                strTmp2 = ""
            Else
                strTmp2 = iTmp1.ToString.PadLeft(2) & "-"
            End If
            strTmp1 = strTmp2
            ' ��2�R�[�i�[�ʉߏ���
            iTmp1 = structSE(iLoopCnt).Jyuni2c
            If iTmp1 = 0 Then
                strTmp2 = ""
            Else
                strTmp2 = iTmp1.ToString.PadLeft(2) & "-"
            End If
            strTmp1 = strTmp1 & strTmp2
            ' ��3�R�[�i�[�ʉߏ���
            iTmp1 = structSE(iLoopCnt).Jyuni3c
            If iTmp1 = 0 Then
                strTmp2 = ""
            Else
                strTmp2 = iTmp1.ToString.PadLeft(2) & "-"
            End If
            strTmp1 = strTmp1 & strTmp2
            ' ��4�R�[�i�[�ʉߏ���
            iTmp1 = structSE(iLoopCnt).Jyuni4c
            If iTmp1 = 0 Then
                strTmp2 = ""
            Else
                strTmp2 = iTmp1.ToString.PadLeft(2)
            End If
            strTmp1 = strTmp1 & strTmp2
            Me.grdUmaProfile.Text = strTmp1
            iColIdx = iColIdx + 1
            ' �\��[�P�I�b�Y]
            Me.grdUmaProfile.Col = iColIdx
            If structSE(iLoopCnt).Odds.Equals("0000") = False Then
                iTmp1 = structSE(iLoopCnt).Odds.Substring(0, 3)
                Me.grdUmaProfile.Text = iTmp1 & "." & structSE(iLoopCnt).Odds.Substring(3, 1)
            End If
            iColIdx = iColIdx + 1
            ' �\��[�P�l�C]
            Me.grdUmaProfile.Col = iColIdx
            iTmp1 = structSE(iLoopCnt).Ninki
            If iTmp1 <> 0 Then
                Me.grdUmaProfile.Text = iTmp1
            End If
            iColIdx = iColIdx + 1
            ' �\��[�n�̏d]
            Me.grdUmaProfile.Col = iColIdx
            If structSE(iLoopCnt).BaTaijyu.Equals("   ") Then
                Me.grdUmaProfile.Text = structSE(iLoopCnt).BaTaijyu
            Else
                iTmp1 = structSE(iLoopCnt).BaTaijyu
                If iTmp1 <> 0 And iTmp1 <> 999 Then
                    Me.grdUmaProfile.Text = iTmp1.ToString & "kg"
                End If
            End If
            iColIdx = iColIdx + 1
            ' �\��[����]
            Me.grdUmaProfile.Col = iColIdx
            If structSE(iLoopCnt).ZogenFugo.Equals(" ") Then
                Select Case structSE(iLoopCnt).ZogenSa
                    Case "000"
                        strTmp1 = "�}0"
                    Case "999"
                        strTmp1 = "----"
                    Case "   "
                        strTmp1 = "    "
                End Select
            Else
                iTmp1 = structSE(iLoopCnt).ZogenSa
                strTmp1 = structSE(iLoopCnt).ZogenFugo & iTmp1
            End If
            Me.grdUmaProfile.Text = strTmp1
            iColIdx = iColIdx + 1
            ' �\��[�l���{�܋�]
            Me.grdUmaProfile.Col = iColIdx
            iTmp1 = structSE(iLoopCnt).Honsyokin & "00"
            Me.grdUmaProfile.Text = Format(iTmp1, "#,0")
            iColIdx = iColIdx + 1
            ' �\��[�l���t���܋�]
            Me.grdUmaProfile.Col = iColIdx
            iTmp1 = structSE(iLoopCnt).Fukasyokin & "00"
            Me.grdUmaProfile.Text = Format(iTmp1, "#,0")
            iColIdx = iColIdx + 1
            ' �\��[��3�n����]
            Me.grdUmaProfile.Col = iColIdx
            If structSE(iLoopCnt).HaronTimeL3 = "999" Then
                Me.grdUmaProfile.Text = "----"
            ElseIf structSE(iLoopCnt).HaronTimeL3 = "000" Then
                Me.grdUmaProfile.Text = ""
            Else
                Me.grdUmaProfile.Text = structSE(iLoopCnt).HaronTimeL3.Substring(0, 2) & "." & structSE(iLoopCnt).HaronTimeL3.Substring(2, 1)
            End If
            iColIdx = iColIdx + 1
            ' �\��[1(2)���n]
            Me.grdUmaProfile.Col = iColIdx
            ' 2���n�̖��O�ɂ͊��ʂ�����
            If structSE(iLoopCnt).KakuteiJyuni.Equals("01") Then
                Me.grdUmaProfile.Text = "(" & TrimSP(structSE(iLoopCnt).ChakuUmaInfo(0).Bamei) & ")"
            Else
                Me.grdUmaProfile.Text = TrimSP(structSE(iLoopCnt).ChakuUmaInfo(0).Bamei)
            End If
            iColIdx = iColIdx + 1
            ' �\��[�^�C����]
            Me.grdUmaProfile.Col = iColIdx
            If structSE(iLoopCnt).TimeDiff.Equals("0000") = False And structSE(iLoopCnt).TimeDiff.Equals("9999") = False Then
                iTmp1 = structSE(iLoopCnt).TimeDiff.Substring(1, 2)
                Me.grdUmaProfile.Text = structSE(iLoopCnt).TimeDiff.Substring(0, 1) & iTmp1 & "." & structSE(iLoopCnt).TimeDiff.Substring(3, 1)
            End If
            iColIdx = iColIdx + 1
        Next iLoopCnt

        '' �Z�����̌���
        ' 
        ' ����ێ�����z��
        Dim strWidth(Me.grdUmaProfile.Cols - 1) As Integer
        ' ��P�ʂŃ��[�v
        For iLoopCnt = 0 To strWidth.Length - 1
            Me.grdUmaProfile.Col = iLoopCnt
            iTmp1 = 0
            iTmp2 = 0
            ' ��s������
            For jLoopCnt = 0 To structSE.Length
                Me.grdUmaProfile.Row = jLoopCnt
                iTmp1 = Str2Byte(Me.grdUmaProfile.get_TextMatrix(jLoopCnt, iLoopCnt)).Length
                If iTmp1 > iTmp2 Then
                    ' ���̗�̍ő啝(byte�P��)��strWidth�Ɋi�[
                    strWidth(iLoopCnt) = iTmp1
                    iTmp2 = iTmp1
                End If
            Next jLoopCnt

        Next iLoopCnt

        ' strWidth�Ɋi�[���ꂽ�������ɃO���b�h�̃Z�������w��
        For iLoopCnt = 0 To strWidth.Length - 1
            Me.grdUmaProfile.set_ColWidth(iLoopCnt, 100 + strWidth(iLoopCnt) * 100)
        Next iLoopCnt

ExitHandler:
        Exit Sub

ErrorHandler:
        Me.Close()
        MsgBox("�Y���f�[�^�͖��擾�ł�", MsgBoxStyle.Information)
        Exit Sub

    End Sub

End Class
