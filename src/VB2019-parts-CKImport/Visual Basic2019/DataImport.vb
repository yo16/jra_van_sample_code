Imports System.Diagnostics
Public Class DataImport
    Private RecIDCurrent As String = String.Empty ''���R�[�h���(����)
    Private RecIDOld As String = String.Empty ''���R�[�h���(�P�O�̃o�b�t�@)
    Private ImportObject As ImportBase

    Public Sub New()
        MyBase.New()
        Me.Class_Initialize_Renamed()
    End Sub

    Private Sub Class_Initialize_Renamed()
        Dim ConnectionString As String ''�ڑ�������

        Try
            gCon = New ADODB.Connection()

            'DB�̃p�X���w�肵�ăR�l�N�V�����I�[�v��
            ConnectionString = My.Settings.ConnectionString
            If ConnectionString.Equals(String.Empty) Then
                ConnectionString = "Provider=Microsoft.ACE.OLEDB.16.0; Data Source=" & My.Application.Info.DirectoryPath & "\CKData.accdb"
            End If
            gCon.Open(ConnectionString)

            Debug.WriteLine("gCon.Open")
        Catch ex As Exception
            Debug.WriteLine(ex.Message)

            Throw
        End Try
    End Sub

    Public Sub Close()
        '���R�[�h���ID�ɑO��ǂ񂾂��̂��c���Ă����
        If RecIDOld <> "" Then
            '�����N���X��j��
            ImportObject.Close()
            ImportObject = Nothing
        End If

        '�R�l�N�V�����N���[�Y
        gCon.Close()
        Debug.WriteLine("gCon.Close")
    End Sub

    ' @(f)
    '
    ' �@�\      : �e�[�u���N���A
    '
    ' ������    :
    '
    ' �Ԃ�l    :
    '
    ' �@�\����  :
    '
    Public Sub ClearData(Optional ByVal TableName As String = "")
        Dim SQL As String ''SQL��

        Try
            '�J�n����
            gCon.BeginTrans()

            If TableName.Equals(String.Empty) = False Then
                '�w�肵���e�[�u�����폜����
                SQL = "DELETE * FROM " & TableName
                gCon.Execute(SQL)
            Else
                '�e�[�u�������擾
                Dim TableArray() As String = My.Settings.TableName.Split(",")

                '�e�[�u���̓��e��S�č폜����
                For i As Integer = 0 To TableArray.Length - 1
                    SQL = "DELETE * FROM " & TableArray(i)
                    gCon.Execute(SQL)
                Next
            End If

            '�I������
            gCon.CommitTrans()
            Debug.WriteLine("gCon.CommitTrans")
        Catch ex As Exception
            '���~����
            gCon.RollbackTrans()
            Debug.WriteLine(ex.Message)

            Throw
        End Try
    End Sub

    Public Sub SetData(ByRef strBuff As String, ByVal lngBuffSize As Integer)
        Try
            '���R�[�h���ID���擾
            RecIDCurrent = Left(strBuff, 2)
            Debug.WriteLine("SetData " & RecIDCurrent)

            If RecIDOld.Equals(RecIDCurrent) = False Then
                '���R�[�h���ID�ɑO��ǂ񂾂��̂��c���Ă����
                If RecIDOld.Equals(String.Empty) = False Then
                    '�����N���X��j��
                    If RecIDCurrent.Equals(RecIDOld) = False Then
                        ImportObject.Close()
                        ImportObject = Nothing
                    End If
                End If

                '�ΏۃN���X�̃C���X�^���X�𐶐�
                Dim classType As Type = Type.GetType("SampleProject.Import" & RecIDCurrent, True)
                ImportObject = CType(Activator.CreateInstance(classType), ImportBase)
            End If

            'DB�ǉ�����
            ImportObject.Add(strBuff, lngBuffSize)

            '���R�[�h���ID��ێ�
            RecIDOld = RecIDCurrent
        Catch ex As Exception
            Debug.WriteLine(ex.Message)

            Throw
        End Try
    End Sub
End Class
