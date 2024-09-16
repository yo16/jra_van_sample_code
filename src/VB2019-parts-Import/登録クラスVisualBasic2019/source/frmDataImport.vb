Option Strict Off
Option Explicit On
Friend Class frmDataImport
	Inherits System.Windows.Forms.Form
#Region "Windows �t�H�[�� �f�U�C�i�ɂ���Đ������ꂽ�R�[�h"
	Public Sub New()
		MyBase.New()
		If m_vb6FormDefInstance Is Nothing Then
			If m_InitializingDefInstance Then
				m_vb6FormDefInstance = Me
			Else
				Try 
					'�X�^�[�g�A�b�v �t�H�[���ɂ��ẮA�ŏ��ɍ쐬���ꂽ�C���X�^���X������C���X�^���X�ɂȂ�܂��B
					If System.Reflection.Assembly.GetExecutingAssembly.EntryPoint.DeclaringType Is Me.GetType Then
						m_vb6FormDefInstance = Me
					End If
				Catch
				End Try
			End If
		End If
		'���̌Ăяo���́AWindows �t�H�[�� �f�U�C�i�ŕK�v�ł��B
		InitializeComponent()
	End Sub
	'Form �́A�R���|�[�l���g�ꗗ�Ɍ㏈�������s���邽�߂� dispose ���I�[�o�[���C�h���܂��B
	Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Windows �t�H�[�� �f�U�C�i�ŕK�v�ł��B
	Private components As System.ComponentModel.IContainer
    Public WithEvents rbtNormal As System.Windows.Forms.RadioButton
    Public WithEvents rbtIsthisweek As System.Windows.Forms.RadioButton
	Public WithEvents rbtSetup As System.Windows.Forms.RadioButton
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
	Public WithEvents txtFromTime As System.Windows.Forms.TextBox
	Public WithEvents txtDataSpec As System.Windows.Forms.TextBox
	Public WithEvents Timer1 As System.Windows.Forms.Timer
	Public WithEvents cmdJVSetting As System.Windows.Forms.Button
	Public WithEvents cmdRead As System.Windows.Forms.Button
    Public WithEvents lblFromTime As System.Windows.Forms.Label
    Public WithEvents lblDataSpec As System.Windows.Forms.Label
	'���� : �ȉ��̃v���V�[�W���� Windows �t�H�[�� �f�U�C�i�ŕK�v�ł��B
	'Windows �t�H�[�� �f�U�C�i���g���ĕύX�ł��܂��B
	'�R�[�h�G�f�B�^���g���ďC�����Ȃ��ł��������B
'    Friend WithEvents JVLink1 As AxJVDTLabLib.AxJVLink
    'Friend WithEvents JVLink1 As AxJVDTLabLib.AxJVLink
    Friend WithEvents JVLink1 As AxJVDTLabLib.AxJVLink

    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmDataImport))
        Me.Frame1 = New System.Windows.Forms.GroupBox
        Me.rbtNormal = New System.Windows.Forms.RadioButton
        Me.rbtIsthisweek = New System.Windows.Forms.RadioButton
        Me.rbtSetup = New System.Windows.Forms.RadioButton
        Me.txtFromTime = New System.Windows.Forms.TextBox
        Me.txtDataSpec = New System.Windows.Forms.TextBox
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.cmdJVSetting = New System.Windows.Forms.Button
        Me.cmdRead = New System.Windows.Forms.Button
        Me.lblFromTime = New System.Windows.Forms.Label
        Me.lblDataSpec = New System.Windows.Forms.Label
        Me.JVLink1 = New AxJVDTLabLib.AxJVLink
        Me.Frame1.SuspendLayout()
        CType(Me.JVLink1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.rbtNormal)
        Me.Frame1.Controls.Add(Me.rbtIsthisweek)
        Me.Frame1.Controls.Add(Me.rbtSetup)
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(8, 72)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(353, 49)
        Me.Frame1.TabIndex = 6
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "�擾�f�[�^"
        '
        'rbtNormal
        '
        Me.rbtNormal.BackColor = System.Drawing.SystemColors.Control
        Me.rbtNormal.Checked = True
        Me.rbtNormal.Cursor = System.Windows.Forms.Cursors.Default
        Me.rbtNormal.ForeColor = System.Drawing.SystemColors.ControlText
        Me.rbtNormal.Location = New System.Drawing.Point(16, 24)
        Me.rbtNormal.Name = "rbtNormal"
        Me.rbtNormal.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.rbtNormal.Size = New System.Drawing.Size(89, 17)
        Me.rbtNormal.TabIndex = 9
        Me.rbtNormal.TabStop = True
        Me.rbtNormal.Text = "�ʏ�f�[�^"
        Me.rbtNormal.UseVisualStyleBackColor = False
        '
        'rbtIsthisweek
        '
        Me.rbtIsthisweek.BackColor = System.Drawing.SystemColors.Control
        Me.rbtIsthisweek.Cursor = System.Windows.Forms.Cursors.Default
        Me.rbtIsthisweek.ForeColor = System.Drawing.SystemColors.ControlText
        Me.rbtIsthisweek.Location = New System.Drawing.Point(112, 24)
        Me.rbtIsthisweek.Name = "rbtIsthisweek"
        Me.rbtIsthisweek.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.rbtIsthisweek.Size = New System.Drawing.Size(113, 17)
        Me.rbtIsthisweek.TabIndex = 8
        Me.rbtIsthisweek.TabStop = True
        Me.rbtIsthisweek.Text = "���T�J�Ãf�[�^"
        Me.rbtIsthisweek.UseVisualStyleBackColor = False
        '
        'rbtSetup
        '
        Me.rbtSetup.BackColor = System.Drawing.SystemColors.Control
        Me.rbtSetup.Cursor = System.Windows.Forms.Cursors.Default
        Me.rbtSetup.ForeColor = System.Drawing.SystemColors.ControlText
        Me.rbtSetup.Location = New System.Drawing.Point(232, 24)
        Me.rbtSetup.Name = "rbtSetup"
        Me.rbtSetup.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.rbtSetup.Size = New System.Drawing.Size(113, 17)
        Me.rbtSetup.TabIndex = 7
        Me.rbtSetup.TabStop = True
        Me.rbtSetup.Text = "�Z�b�g�A�b�v�f�[�^"
        Me.rbtSetup.UseVisualStyleBackColor = False
        '
        'txtFromTime
        '
        Me.txtFromTime.AcceptsReturn = True
        Me.txtFromTime.BackColor = System.Drawing.SystemColors.Window
        Me.txtFromTime.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFromTime.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFromTime.Location = New System.Drawing.Point(264, 40)
        Me.txtFromTime.MaxLength = 0
        Me.txtFromTime.Name = "txtFromTime"
        Me.txtFromTime.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFromTime.Size = New System.Drawing.Size(97, 19)
        Me.txtFromTime.TabIndex = 4
        Me.txtFromTime.Text = "20111201000000"
        '
        'txtDataSpec
        '
        Me.txtDataSpec.AcceptsReturn = True
        Me.txtDataSpec.BackColor = System.Drawing.SystemColors.Window
        Me.txtDataSpec.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDataSpec.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDataSpec.Location = New System.Drawing.Point(18, 42)
        Me.txtDataSpec.MaxLength = 0
        Me.txtDataSpec.Name = "txtDataSpec"
        Me.txtDataSpec.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDataSpec.Size = New System.Drawing.Size(241, 19)
        Me.txtDataSpec.TabIndex = 3
        '
        'Timer1
        '
        Me.Timer1.Interval = 1
        '
        'cmdJVSetting
        '
        Me.cmdJVSetting.BackColor = System.Drawing.SystemColors.Control
        Me.cmdJVSetting.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdJVSetting.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdJVSetting.Location = New System.Drawing.Point(400, 88)
        Me.cmdJVSetting.Name = "cmdJVSetting"
        Me.cmdJVSetting.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdJVSetting.Size = New System.Drawing.Size(101, 37)
        Me.cmdJVSetting.TabIndex = 1
        Me.cmdJVSetting.Text = "JVLink�ݒ�"
        Me.cmdJVSetting.UseVisualStyleBackColor = False
        '
        'cmdRead
        '
        Me.cmdRead.BackColor = System.Drawing.SystemColors.Control
        Me.cmdRead.Cursor = System.Windows.Forms.Cursors.Default
        Me.cmdRead.ForeColor = System.Drawing.SystemColors.ControlText
        Me.cmdRead.Location = New System.Drawing.Point(400, 16)
        Me.cmdRead.Name = "cmdRead"
        Me.cmdRead.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cmdRead.Size = New System.Drawing.Size(101, 51)
        Me.cmdRead.TabIndex = 0
        Me.cmdRead.Text = "�擾�J�n"
        Me.cmdRead.UseVisualStyleBackColor = False
        '
        'lblFromTime
        '
        Me.lblFromTime.BackColor = System.Drawing.SystemColors.Control
        Me.lblFromTime.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblFromTime.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblFromTime.Location = New System.Drawing.Point(264, 16)
        Me.lblFromTime.Name = "lblFromTime"
        Me.lblFromTime.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblFromTime.Size = New System.Drawing.Size(73, 17)
        Me.lblFromTime.TabIndex = 2
        Me.lblFromTime.Text = "FromTime"
        '
        'lblDataSpec
        '
        Me.lblDataSpec.BackColor = System.Drawing.SystemColors.Control
        Me.lblDataSpec.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblDataSpec.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblDataSpec.Location = New System.Drawing.Point(18, 24)
        Me.lblDataSpec.Name = "lblDataSpec"
        Me.lblDataSpec.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblDataSpec.Size = New System.Drawing.Size(73, 17)
        Me.lblDataSpec.TabIndex = 5
        Me.lblDataSpec.Text = "�f�[�^���"
        '
        'JVLink1
        '
        Me.JVLink1.Enabled = True
        Me.JVLink1.Location = New System.Drawing.Point(400, 128)
        Me.JVLink1.Name = "JVLink1"
        Me.JVLink1.OcxState = CType(resources.GetObject("JVLink1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.JVLink1.Size = New System.Drawing.Size(192, 192)
        Me.JVLink1.TabIndex = 7
        '
        'frmDataImport
        '
        Me.AcceptButton = Me.cmdRead
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 12)
        Me.ClientSize = New System.Drawing.Size(512, 173)
        Me.Controls.Add(Me.JVLink1)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.txtFromTime)
        Me.Controls.Add(Me.txtDataSpec)
        Me.Controls.Add(Me.cmdJVSetting)
        Me.Controls.Add(Me.cmdRead)
        Me.Controls.Add(Me.lblFromTime)
        Me.Controls.Add(Me.lblDataSpec)
        Me.Location = New System.Drawing.Point(4, 23)
        Me.Name = "frmDataImport"
        Me.Text = "�����n�f�[�^�ǂݍ���"
        Me.Frame1.ResumeLayout(False)
        CType(Me.JVLink1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
#End Region
#Region "�A�b�v�O���[�h �E�B�U�[�h�̃T�|�[�g �R�[�h"
    Private Shared m_vb6FormDefInstance As frmDataImport
    Private Shared m_InitializingDefInstance As Boolean
    Public Shared Property DefInstance() As frmDataImport
        Get
            If m_vb6FormDefInstance Is Nothing OrElse m_vb6FormDefInstance.IsDisposed Then
                m_InitializingDefInstance = True
                m_vb6FormDefInstance = New frmDataImport()
                m_InitializingDefInstance = False
            End If
            DefInstance = m_vb6FormDefInstance
        End Get
        Set(ByVal Value As frmDataImport)
            m_vb6FormDefInstance = Value
        End Set
    End Property
#End Region
    '========================================================================
    '  JRA-VAN Data Lab. �v���O���~���O�p�[�c�u�T���v���v���O�����v
    '
    '
    '   �쐬: JRA-VAN �\�t�g�E�F�A�H�[  2003�N 6�� 3��
    '	�X�V:                           2007�N11�� 8��
    '
    '========================================================================
    '   (C) Copyright JRA SYSTEM SERVICE CO.,LTD. 2007 All rights reserved
    '========================================================================

    Private objConnect As clsDBImport ''�N���X

    'JVLink�̐ݒ�
    Private Sub cmdJVSetting_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdJVSetting.Click
        If JVLink1.JVSetUIProperties = -1 Then
            MsgBox("�G���[�̂���JV-Link�̐ݒ�Ɏ��s���܂���")
        End If
    End Sub


    '�ǂݍ��ݏ����J�n
    Private Sub cmdRead_Click(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cmdRead.Click
        On Error GoTo ErrH

        Dim lngBuffSize As Integer = 110000
        Const lngFileNameSize As Integer = 256

        Dim lngReturnCode As Integer 'JVLink����̖߂�l
        Dim strDataSpec As String 'JVOpen �f�[�^���
        Dim strFromTime As String
        Dim lngOptionFlag As Integer
        Dim lngReadCount As Integer
        Dim lngDownloadCount As Integer
        Dim strLastTime As String = ""
        Dim strFileName As String
        Dim strBuff As String
        Dim sngTimerStart As Single
        Dim sngTimerEnd As Single
        Dim blnDelFlg As Boolean

        blnDelFlg = False

        objConnect = New clsDBImport()

        If MsgBox("�捞�݂��J�n���܂��B�e�[�u�����N���A���܂����H", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
            Call objConnect.ClearData()
        End If

        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
        cmdRead.Enabled = False
        cmdJVSetting.Enabled = False

        'JVInit
        lngReturnCode = Me.JVLink1.JVInit("UNKNOWN")
        If lngReturnCode <> 0 Then
            MsgBox("JVLink - JVInit�G���[")
            Me.Cursor = System.Windows.Forms.Cursors.Default
            cmdRead.Enabled = True
            cmdJVSetting.Enabled = True
            Exit Sub
        End If

        'JVOpen
        strDataSpec = txtDataSpec.Text '�f�[�^���
        strFromTime = txtFromTime.Text 'FromTime

        If rbtNormal.Checked = True Then
            lngOptionFlag = 1
        ElseIf rbtIsthisweek.Checked = True Then
            lngOptionFlag = 2
        ElseIf rbtSetup.Checked = True Then
            lngOptionFlag = 3
        End If

        lngReturnCode = Me.JVLink1.JVOpen(strDataSpec, strFromTime, lngOptionFlag, lngReadCount, lngDownloadCount, strLastTime)
        'lngReturnCode = Me.JVLink1.JVRTOpen(strDataSpec, strFromTime)
        If lngReturnCode < 0 Then
            MsgBox("JVLink - JVOpen�G���[")
            Me.Cursor = System.Windows.Forms.Cursors.Default
            cmdRead.Enabled = True
            cmdJVSetting.Enabled = True
            Exit Sub
        End If


        'JVRead�̃��[�v����


        '�o�b�t�@�쐬
        strBuff = New String(vbNullChar, lngBuffSize)
        strFileName = New String(vbNullChar, lngFileNameSize)
        Dim recordspec As String

        Do

            Application.DoEvents()

            'JVRead��1�s�ǂݍ���
            lngReturnCode = JVLink1.JVRead(strBuff, lngBuffSize, strFileName)

            '���^�[���R�[�h�ɂ�菈���𕪊�
            Select Case lngReturnCode
                Case 0 ' �S�t�@�C���ǂݍ��ݏI��
                    Exit Do
                Case -1 ' �t�@�C���؂�ւ��
                Case -3 ' �_�E�����[�h��
                Case -201 ' Init����ĂȂ�
                    MsgBox("JVInit���s���Ă��܂���B")
                    Exit Do
                Case -203 ' Open����ĂȂ�
                    MsgBox("JVOpen���s���Ă��܂���B")
                    Exit Do
                Case -503 ' �t�@�C�����Ȃ�
                    Exit Do
                Case Is > 0 ' ����ǂݍ���
                    recordspec = Mid(strBuff, 1, 2)
                    Call objConnect.SetData(strBuff, lngBuffSize)

            End Select
        Loop While (1)


        '���
        objConnect.Close()
        objConnect = Nothing

        'JVClose
        JVLink1.JVClose()

        Me.Cursor = System.Windows.Forms.Cursors.Default
        cmdRead.Enabled = True
        cmdJVSetting.Enabled = True

        MsgBox("�S�f�[�^�̓ǂݍ��ݏ������I�����܂���")

        Exit Sub
ErrH:
        MsgBox(Err.Description)
    End Sub

    Private Sub rbtSetup_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbtSetup.CheckedChanged

    End Sub

    Private Sub frmDataImport_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim iniFilePath As String = CurDir() + "\sample.ini"

        ' �ݒ�t�@�C�����݃`�F�b�N
        If Dir(iniFilePath) = "" Then
            MsgBox("�����ݒ�t�@�C��(sample.ini)��������܂���")
            Exit Sub
        End If

        ' -----�ݒ�t�@�C�����A�e����̓ǂݍ���-----
        ' DB�ڑ�������̎擾
        strConnectString = GetProfileDataStr("Setting", "DBConnectString", iniFilePath)
        If strConnectString = "" Then
            MsgBox("�f�[�^�x�[�X�ڑ�������̎擾�Ɏ��s���܂����B", MessageBoxIcon.Error)
            Exit Sub
        End If

        ' DB���[�h�̎擾
        Dim strDBMode As String
        strDBMode = GetProfileDataStr("Setting", "DBMode", iniFilePath)
        If strDBMode = "" Then
            MsgBox("�f�[�^�x�[�X���[�h�̎擾�Ɏ��s���܂����B", MessageBoxIcon.Error)
            Exit Sub
        End If
        If strDBMode = "0" Then
            SS = "["
            SE = "]"
        Else
            SS = ""
            SE = ""
        End If

    End Sub
End Class