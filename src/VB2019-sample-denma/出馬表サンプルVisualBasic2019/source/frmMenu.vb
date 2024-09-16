Option Explicit On 

Public Class frmMenu
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
        Me.btnGetJVData.Text = "�J�Ï��擾"
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
        Me.btnInitDB.Text = "�c�a������"
        '
        'btnViewDenmaList
        '
        Me.btnViewDenmaList.Location = New System.Drawing.Point(8, 88)
        Me.btnViewDenmaList.Name = "btnViewDenmaList"
        Me.btnViewDenmaList.Size = New System.Drawing.Size(184, 40)
        Me.btnViewDenmaList.TabIndex = 5
        Me.btnViewDenmaList.Text = "�o�n�\�\��"
        '
        'btnSettingJVLink
        '
        Me.btnSettingJVLink.Location = New System.Drawing.Point(104, 144)
        Me.btnSettingJVLink.Name = "btnSettingJVLink"
        Me.btnSettingJVLink.Size = New System.Drawing.Size(88, 40)
        Me.btnSettingJVLink.TabIndex = 6
        Me.btnSettingJVLink.Text = "JV-Link�ݒ�"
        '
        'btnStopJVData
        '
        Me.btnStopJVData.Enabled = False
        Me.btnStopJVData.Location = New System.Drawing.Point(104, 8)
        Me.btnStopJVData.Name = "btnStopJVData"
        Me.btnStopJVData.Size = New System.Drawing.Size(88, 40)
        Me.btnStopJVData.TabIndex = 7
        Me.btnStopJVData.Text = "�L�����Z��"
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
        Me.Text = "�T���v���v���O���� �| ���j���["
        CType(Me.JVLink, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Declare Function GetActiveWindow Lib "USER32" () As Integer

    ' �L�����Z���t���O
    Private bCancelFlag As Boolean

    ' �J�����g�p�X
    Dim strCurPath As String

    ' FromTime
    Dim strFromTime As String

    '�G���[���b�Z�[�W
    Dim strErrMsg As String

    Private Sub btnGetJVData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnGetJVData.Click

        Try
            ' ���^�[���R�[�h
            Dim lReturnCode As Long

            ' JVOpen�p(DataSpec)
            Dim strDataSpec As String
            ' JVOpen�p(Option)
            Dim lOption As Long
            ' JVOpen�p(ReadCount)
            Dim lReadCount As Long
            ' JVOpen�p(DownloadCount)
            Dim lDownloadCount As Long
            ' JVOpen�p(LastFileTimestamp)
            Dim strLastFileTimestamp As String = String.Empty

            ' JVGets�p(�o�b�t�@�|�C���^)
            Dim szBuff(0) As Byte
            ' JVGets�p(�o�b�t�@)
            Dim strBuff As String
            ' JVGets�p(�o�b�t�@�T�C�Y)
            Dim lBuffSize As Long
            ' JVGets�p(�t�@�C����)
            Dim strFileName As String = String.Empty

            ' �f�[�^�敪
            Dim strRecID As String

            ' �L�����Z�����m�t���O�̏�����
            bCancelFlag = False

            Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

            ' �v���O���X�o�[�̏�����
            Me.barFileCount.Value = 0
            Me.barReadSize.Value = 0

            ' JVOpen�̌Ăяo��
            strDataSpec = "RACERCVN"
            lOption = "2"
            'strDataSpec = "RACE"
            'lOption = "4"
            lBuffSize = 110000
            lReturnCode = Me.JVLink.JVOpen(strDataSpec, strFromTime, lOption, lReadCount, lDownloadCount, strLastFileTimestamp)
            If lReturnCode = ST_READ_EOF Then
                Me.Cursor = System.Windows.Forms.Cursors.Default
                MsgBox("�f�[�^�x�[�X�͍ŐV�̏�Ԃł��B")
                Exit Sub
            End If
            If lReturnCode < 0 Then
                Me.Cursor = System.Windows.Forms.Cursors.Default
                MsgBox("JVOpen�G���[ �R�[�h�F" & lReturnCode & "�F", MessageBoxIcon.Error)
                Exit Sub
            End If

            ' �{�^���̗}�~
            Me.btnGetJVData.Enabled = False
            Me.btnViewDenmaList.Enabled = False
            Me.btnInitDB.Enabled = False
            Me.btnSettingJVLink.Enabled = False
            Me.cmbYear.Enabled = False
            Me.btnStopJVData.Enabled = True

            ' ���v�t�@�C�����̃v���O���X�o�[�̏����ݒ�
            Me.barFileCount.Maximum = lReadCount

            ' �v���O���X�o�[�p�ϐ�
            Dim lTotalFileCount As Long
            Dim lTotalReadSize As Long
            lTotalReadSize = 0
            lTotalFileCount = 0

            ' JVSkip����t���O
            Dim bSkipFlg As Boolean

            Do
                ' �o�b�N�O���E���h�ł̏���
                System.Windows.Forms.Application.DoEvents()

                '�L�����Z���������ꂽ�珈��(���[�v)�𔲂���
                If bCancelFlag = True Then Exit Do

                ' JVGets�̌Ăяo��
#Disable Warning BC41999
                lReturnCode = Me.JVLink.JVGets(szBuff, lBuffSize, strFileName)
#Enable Warning BC41999

                ' �G���[����
                Select Case lReturnCode

                    Case Is > ST_READ_SUCCESS
                        ' ����

                        ' �����R�[�h�ϊ�(SJIS��UNICODE)
                        strBuff = System.Text.Encoding.GetEncoding(932).GetString(szBuff)

                        ' �f�[�^�敪�̎擾
                        strRecID = strBuff.Substring(0, 2)

                        bSkipFlg = False

                        ' �����Ώۃf�[�^�̂݃f�[�^�x�[�X�֓o�^
                        If strRecID = ID_RACE Then
                            ' ���[�X��ڍ�
                            ImportRA.Add(strBuff, lBuffSize)
                        ElseIf strRecID = ID_RACE_UMA Then
                            ' �n�����[�X���
                            ImportSE.Add(strBuff, lBuffSize)
                        ElseIf strRecID = ID_UMA Then
                            ' �����n�}�X�^
                            ImportUM.Add(strBuff, lBuffSize)
                        Else
                            '�ΏۊO�t�@�C���̓X�L�b�v(�t���O��ݒ�)
                            bSkipFlg = True
                        End If

                        If bSkipFlg = True Then
                            '�ΏۊO�t�@�C���̓X�L�b�v
                            Me.JVLink.JVSkip()

                            ' �J�����g�t�@�C���̃v���O���X�o�[���X�V
                            Me.barReadSize.Value = Me.barReadSize.Maximum

                            ' ���v�t�@�C�����̃v���O���X�o�[���X�V
                            lTotalFileCount = lTotalFileCount + 1
                            Me.barFileCount.Value = lTotalFileCount

                        Else
                            ' �J�����g�t�@�C���̃v���O���X�o�[���X�V
                            Me.barReadSize.Maximum = Me.JVLink.m_CurrentReadFilesize
                            lTotalReadSize = lTotalReadSize + szBuff.Length - 1
                            Me.barReadSize.Value = lTotalReadSize
                        End If
                        ReDim szBuff(0)

                    Case ST_READ_EOF
                        ' �t�@�C���̋�؂�

                        ' ���v�t�@�C�����̃v���O���X�o�[���X�V
                        lTotalFileCount = lTotalFileCount + 1
                        Me.barFileCount.Value = lTotalFileCount

                        ' �J�����g�t�@�C���̃v���O���X�o�[��������
                        lTotalReadSize = 0

                        ' FromTime��ޔ�
                        strFromTime = Me.JVLink.m_CurrentFileTimeStamp

                    Case ST_READ_EOL
                        ' �S���R�[�h�Ǎ��ݏI��(EOF)
                        Exit Do

                    Case ST_READ_DOWNLOAD_NOW
                        ' �_�E�����[�h���̏ꍇ�A1�b�X���[�v���_�E�����[�h�҂�
                        System.Threading.Thread.Sleep(1000)

                    Case Is <= ST_READ_ERR
                        ' �G���[
                        MsgBox("JVGets�G���[ �R�[�h�F" & lReturnCode & "�F", MessageBoxIcon.Error)
                        Exit Do

                End Select
            Loop

            ' FromTime��ini�t�@�C���ɕۑ�
            WriteProfileDataStr("Setting", "FromTime", strFromTime, strCurPath)

            ' JVClose�̌ďo
            lReturnCode = Me.JVLink.JVClose()
            If lReturnCode <> 0 Then
                MsgBox("JVClose�G���[ �R�[�h�F" & lReturnCode & "�F", MessageBoxIcon.Error)
            End If

            If bCancelFlag = False Then
                MsgBox("�J�Ï��̎擾���I�����܂����B", MsgBoxStyle.ApplicationModal)
            Else
                MsgBox("�J�Ï��̎擾�𒆎~���܂����B", MsgBoxStyle.ApplicationModal)
            End If

            ' �J�ÔN�����I���R���{�{�b�N�X�̕\��
            getRaceYMDList(Me.cmbYear)

            ' �{�^���̗}�~������
            Me.btnGetJVData.Enabled = True
            Me.btnViewDenmaList.Enabled = True
            Me.btnInitDB.Enabled = True
            Me.btnSettingJVLink.Enabled = True
            Me.cmbYear.Enabled = True
            Me.btnStopJVData.Enabled = False
            ' �v���O���X�o�[�����ɖ߂�
            Me.barFileCount.Value = 0
            Me.barReadSize.Value = 0


        Catch
            Debug.WriteLine(Err.Description)
        End Try

        Me.Cursor = System.Windows.Forms.Cursors.Default

    End Sub

    Private Sub btnInitDB_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnInitDB.Click

        ' ���^�[���R�[�h
        Dim lReturnCode As Long

        lReturnCode = MsgBox("�c�a�����������܂����H", MsgBoxStyle.YesNo + MsgBoxStyle.Question, "�m�F")
        If lReturnCode = DialogResult.No Then
            Exit Sub
        End If

        Me.Cursor = System.Windows.Forms.Cursors.WaitCursor

        ' �v���O���X�o�[�̏�����
        Me.barReadSize.Value = 0
        Me.barFileCount.Value = 0

        Me.barFileCount.Maximum = 100

        '�f�[�^�x�[�X�o�^�N���X�̃N���[�Y
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

        ' �e�[�u���̑S���R�[�h���N���A
        gCon.Execute("DELETE FROM RACE")
        Me.barFileCount.Value = 30

        gCon.Execute("DELETE FROM UMA_RACE")
        Me.barFileCount.Value = 60

        gCon.Execute("DELETE FROM UMA")
        Me.barFileCount.Value = 90

        '�f�[�^�x�[�X�o�^�N���X�̐���
        ImportRA = New clsImportRA()
        ImportSE = New clsImportSE()
        ImportUM = New clsImportUM()


        ' FromTime����������ini�t�@�C���ɕۑ�
        strFromTime = "00000000000000"
        WriteProfileDataStr("Setting", "FromTime", "00000000000000", strCurPath)

        ' �J�ÔN�����I���R���{�{�b�N�X�̃N���A
        getRaceYMDList(Me.cmbYear)

        Me.btnViewDenmaList.Enabled = False
        Me.cmbYear.Enabled = False

        Me.barFileCount.Value = 100

        Me.Cursor = System.Windows.Forms.Cursors.Default

        MsgBox("�c�a�̏��������I�����܂����B")

    End Sub

    Private Sub frmMenu_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim lReturnCode As Long
        Dim iHWnd As Integer

        strCurPath = CurDir() + "\sample.ini"

        ' �ݒ�t�@�C�����݃`�F�b�N
        If Dir(strCurPath) = "" Then
            MsgBox("�����ݒ�t�@�C��(sample.ini)��������܂���")
            Exit Sub
        End If

        ' -----�ݒ�t�@�C�����A�e����̓ǂݍ���-----
        ' DB�ڑ�������̎擾
        strConnectString = GetProfileDataStr("Setting", "DBConnectString", strCurPath)
        If strConnectString = "" Then
            MsgBox("�f�[�^�x�[�X�ڑ�������̎擾�Ɏ��s���܂����B", MessageBoxIcon.Error)
            Exit Sub
        End If

        ' DB���[�h�̎擾
        Dim strDBMode As String
        strDBMode = GetProfileDataStr("Setting", "DBMode", strCurPath)
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

        ' FROMTIME
        strFromTime = GetProfileDataStr("Setting", "FromTime", strCurPath)
        If strFromTime = "" Then
            strFromTime = "00000000000000"
        End If

        ' �C���X�^���X����
        objCDCv = New clsCodeConv()

        ' �p�X���w�肵�A�R�[�h�t�@�C����Ǎ���
        Dim strPath As String
        strPath = System.Reflection.Assembly.GetExecutingAssembly.Location()
        strPath = System.IO.Path.GetDirectoryName(strPath)
        objCDCv.FileName = strPath & "\CodeTable.csv"

        ' JVInit�̌Ăяo��
        lReturnCode = Me.JVLink.JVInit("JVLinkSDKSampleAPP1")
        If lReturnCode <> 0 Then
            MsgBox("JVInit�G���[ �R�[�h�F" & lReturnCode & "�F", MessageBoxIcon.Error)
            Me.Cursor = System.Windows.Forms.Cursors.Default
            Exit Sub
        End If

        ' JV-Link�ւ̃E�B���h�E�n���h���o�^
        iHWnd = GetActiveWindow()
        Me.JVLink.ParentHWnd = iHWnd

        ' �f�[�^�x�[�X�Ƃ̐ڑ����s���B
        If ConnectDB() = True Then
            ' �f�[�^�x�[�X�o�^�N���X�̐���
            ImportRA = New clsImportRA()
            ImportSE = New clsImportSE()
            ImportUM = New clsImportUM()

            ' �J�ÔN�����I���R���{�{�b�N�X�̕\��
            getRaceYMDList(Me.cmbYear)

            ' �f�[�^�x�[�X�֘A�@�\�{�^�����g�p�ɐݒ�
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
            ' ADODB�I�u�W�F�N�g�̊J��
            gCon = Nothing

            ' �ڑ����s���Ƀf�[�^�x�[�X�֘A�@�\�{�^�����g�p�s�ɐݒ�
            Me.btnGetJVData.Enabled = False
            Me.btnViewDenmaList.Enabled = False
            Me.btnInitDB.Enabled = False

        End If

    End Sub

    Private Sub frmMenu_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed

        '�f�[�^�x�[�X�o�^�N���X�̃N���[�Y
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

        '�f�[�^�x�[�X�Ƃ̐ؒf���s���B
        If gCon Is Nothing = False Then
            gCon.Close()
            gCon = Nothing
        End If

        System.Diagnostics.Debug.WriteLine("gCon.Close")

    End Sub

    Private Sub btnViewDenmaList_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnViewDenmaList.Click

        Dim frmSubForm As New frmDenmaList()

        ' �p�����[�^�̐ݒ�
        frmSubForm.txtParam.Text = cmbYear.Text()

        '���[�h���X�t�H�[���Ƃ��ĕ\��
        frmSubForm.Show()

    End Sub

    Private Sub btnStopJVData_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStopJVData.Click

        ' �L�����Z���t���O�̐ݒ�
        bCancelFlag = True

    End Sub

    Private Sub btnSettingJVLink_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSettingJVLink.Click

        Try

            ' ���^�[���R�[�h
            Dim lReturnCode As Long

            ' �ݒ��ʕ\��
            lReturnCode = JVLink.JVSetUIProperties()
            If lReturnCode <> 0 Then
                MsgBox("JVSetUIProperties�G���[ �R�[�h�F" & lReturnCode & "�F", MessageBoxIcon.Error)
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
        ' �n���E�C�O���[�X�i�f�[�^�敪"A","B"�j�����O
        strSQL = "SELECT distinct Year, MonthDay FROM RACE WHERE not DataKubun in ('A','B') ORDER BY Year desc, MonthDay desc"

        ' �R���{�{�b�N�X�̃N���A
        cmbYMD.Text = ""
        cmbYMD.Items.Clear()

        ' ���R�[�h�Z�b�g�̃I�[�v��
        dbRS = New ADODB.Recordset()
        dbRS.Open(strSQL, gCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockPessimistic)

        While Not dbRS.EOF
            ' �t�B�[���h�̎擾
            dbFld = dbRS.Fields

            cmbYMD.Items.Add(dbFld("Year").Value() + dbFld("MonthDay").Value())

            dbRS.MoveNext()

        End While

        ' ���ߓ��t�������\��
        If cmbYMD.Items.Count > 0 Then
            cmbYMD.SelectedIndex() = 0
        End If

ExitHandler:
        ' ���R�[�h�Z�b�g�̃N���[�Y
        dbRS.Close()
        dbRS = Nothing

        getRaceYMDList = True

        Exit Function

ErrorHandler:
        'System.Diagnostics.Debug.WriteLine(Err.Description)
        Resume ExitHandler

    End Function

End Class
