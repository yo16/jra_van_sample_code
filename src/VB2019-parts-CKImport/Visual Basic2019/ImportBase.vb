Imports System.Collections.Generic
Imports System.Diagnostics

Public MustInherit Class ImportBase
    Protected mBuf As Object
    Protected mRS(0) As ADODB.Recordset

    Protected Sub New()
    End Sub

    Public Overridable Function Add(ByRef strBuf As String, ByVal lngBufSize As Integer) As Boolean
        Dim strMakeDate As String '' �o�^����f�[�^�̍쐬�N����

        Try
            '�\���̂Ƀf�[�^�Z�b�g
            mBuf.SetData(strBuf)

            With mBuf.head.MakeDate
                strMakeDate = .Year & .Month & .Day
            End With

            'INSERT����
            If Not InsertDB() Then
                'UPDATE�����iINSERT�����s�����ꍇ�j
                If Not UpdateDB(strMakeDate) Then System.Diagnostics.Debug.WriteLine("�X�V�Ɏ��s���܂����B" & Left(strBuf, 2))
            End If

            Return True
        Catch ex As Exception
            Debug.WriteLine(ex.Message)
            Throw
        End Try
    End Function

    Protected Sub Class_Initialize_Renamed(ByVal SQL As List(Of String))
        Dim i As Integer

        Try
            For i = 0 To mRS.Length - 1
                mRS(i) = New ADODB.Recordset()
            Next

            For i = 0 To SQL.Count - 1
                mRS(i).Open(SQL(i), gCon, ADODB.CursorTypeEnum.adOpenForwardOnly, ADODB.LockTypeEnum.adLockPessimistic)
            Next
        Catch ex As Exception
            Debug.WriteLine(ex.Message)
            Throw
        End Try
    End Sub

    Public Overridable Sub Close()
        '���R�[�h�Z�b�g�N���[�Y
        For i As Integer = 0 To mRS.Length - 1
            mRS(i).Close()
        Next

        mRS = Nothing
    End Sub

    Protected MustOverride Function InsertDB() As Boolean

    Protected MustOverride Function UpdateDB(ByRef strMakeDate As String) As Boolean

End Class
