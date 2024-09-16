Public Class DataImportForm

    Private Sub btnRead_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRead.Click
        Const lngFileNameSize As Integer = 256
        Dim lngBuffSize As Integer = 110000
        Dim lngReturnCode As Integer 'JVLink����̖߂�l
        Dim strDataSpec As String 'JVOpen �f�[�^���
        Dim strFromTime As String
        Dim lngOptionFlag As Integer
        Dim lngReadCount As Integer
        Dim lngDownloadCount As Integer
        Dim strLastTime As String = String.Empty
        Dim strFileName As String
        Dim strBuff As String
        Dim blnDelFlg As Boolean

        blnDelFlg = False

        Dim objDB As DataImport

        Try
            objDB = New DataImport()

            If MsgBox("�捞�݂��J�n���܂��B�e�[�u�����N���A���܂����H", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                objDB.ClearData()
            End If

            Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
            Me.btnRead.Enabled = False
            Me.btnJVSetting.Enabled = False

            'JVInit
            lngReturnCode = Me.JVLink1.JVInit("UNKNOWN")
            If lngReturnCode <> 0 Then
                MsgBox("JVLink - JVInit�G���[")
                Me.Cursor = System.Windows.Forms.Cursors.Default
                Me.btnRead.Enabled = True
                Me.btnJVSetting.Enabled = True
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
            If lngReturnCode < 0 Then
                MsgBox("JVLink - JVOpen�G���[")
                Me.Cursor = System.Windows.Forms.Cursors.Default
                Me.btnRead.Enabled = True
                Me.btnJVSetting.Enabled = True
                Exit Sub
            End If

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
                        objDB.SetData(strBuff, lngBuffSize)
                End Select
            Loop While (1)

            '���
            objDB.Close()
            objDB = Nothing

            'JVClose
            JVLink1.JVClose()

            Me.Cursor = System.Windows.Forms.Cursors.Default
            Me.btnRead.Enabled = True
            Me.btnJVSetting.Enabled = True

            MsgBox("�S�f�[�^�̓ǂݍ��ݏ������I�����܂���")
        Catch ex As Exception
            Me.btnRead.Enabled = True
            Me.btnJVSetting.Enabled = True
            MsgBox(ex.Message)
        End Try

    End Sub

    Private Sub btnJVSetting_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnJVSetting.Click
        If Me.JVLink1.JVSetUIProperties = -1 Then
            MsgBox("�G���[�̂���JV-Link�̐ݒ�Ɏ��s���܂���")
        End If
    End Sub
End Class
