Imports System.Diagnostics
Public Class DataImport
    Private RecIDCurrent As String = String.Empty ''レコード種別(現在)
    Private RecIDOld As String = String.Empty ''レコード種別(１つ前のバッファ)
    Private ImportObject As ImportBase

    Public Sub New()
        MyBase.New()
        Me.Class_Initialize_Renamed()
    End Sub

    Private Sub Class_Initialize_Renamed()
        Dim ConnectionString As String ''接続文字列

        Try
            gCon = New ADODB.Connection()

            'DBのパスを指定してコネクションオープン
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
        'レコード種別IDに前回読んだものが残っていれば
        If RecIDOld <> "" Then
            '処理クラスを破棄
            ImportObject.Close()
            ImportObject = Nothing
        End If

        'コネクションクローズ
        gCon.Close()
        Debug.WriteLine("gCon.Close")
    End Sub

    ' @(f)
    '
    ' 機能      : テーブルクリア
    '
    ' 引き数    :
    '
    ' 返り値    :
    '
    ' 機能説明  :
    '
    Public Sub ClearData(Optional ByVal TableName As String = "")
        Dim SQL As String ''SQL文

        Try
            '開始処理
            gCon.BeginTrans()

            If TableName.Equals(String.Empty) = False Then
                '指定したテーブルを削除する
                SQL = "DELETE * FROM " & TableName
                gCon.Execute(SQL)
            Else
                'テーブル名を取得
                Dim TableArray() As String = My.Settings.TableName.Split(",")

                'テーブルの内容を全て削除する
                For i As Integer = 0 To TableArray.Length - 1
                    SQL = "DELETE * FROM " & TableArray(i)
                    gCon.Execute(SQL)
                Next
            End If

            '終了処理
            gCon.CommitTrans()
            Debug.WriteLine("gCon.CommitTrans")
        Catch ex As Exception
            '中止処理
            gCon.RollbackTrans()
            Debug.WriteLine(ex.Message)

            Throw
        End Try
    End Sub

    Public Sub SetData(ByRef strBuff As String, ByVal lngBuffSize As Integer)
        Try
            'レコード種別IDを取得
            RecIDCurrent = Left(strBuff, 2)
            Debug.WriteLine("SetData " & RecIDCurrent)

            If RecIDOld.Equals(RecIDCurrent) = False Then
                'レコード種別IDに前回読んだものが残っていれば
                If RecIDOld.Equals(String.Empty) = False Then
                    '処理クラスを破棄
                    If RecIDCurrent.Equals(RecIDOld) = False Then
                        ImportObject.Close()
                        ImportObject = Nothing
                    End If
                End If

                '対象クラスのインスタンスを生成
                Dim classType As Type = Type.GetType("SampleProject.Import" & RecIDCurrent, True)
                ImportObject = CType(Activator.CreateInstance(classType), ImportBase)
            End If

            'DB追加処理
            ImportObject.Add(strBuff, lngBuffSize)

            'レコード種別IDを保持
            RecIDOld = RecIDCurrent
        Catch ex As Exception
            Debug.WriteLine(ex.Message)

            Throw
        End Try
    End Sub
End Class
