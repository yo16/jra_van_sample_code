Public Class DataImportForm

    Private Sub btnRead_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnRead.Click
        Const lngFileNameSize As Integer = 256
        Dim lngBuffSize As Integer = 110000
        Dim lngReturnCode As Integer 'JVLinkからの戻り値
        Dim strDataSpec As String 'JVOpen データ種別
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

            If MsgBox("取込みを開始します。テーブルをクリアしますか？", MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                objDB.ClearData()
            End If

            Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
            Me.btnRead.Enabled = False
            Me.btnJVSetting.Enabled = False

            'JVInit
            lngReturnCode = Me.JVLink1.JVInit("UNKNOWN")
            If lngReturnCode <> 0 Then
                MsgBox("JVLink - JVInitエラー")
                Me.Cursor = System.Windows.Forms.Cursors.Default
                Me.btnRead.Enabled = True
                Me.btnJVSetting.Enabled = True
                Exit Sub
            End If

            'JVOpen
            strDataSpec = txtDataSpec.Text 'データ種別
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
                MsgBox("JVLink - JVOpenエラー")
                Me.Cursor = System.Windows.Forms.Cursors.Default
                Me.btnRead.Enabled = True
                Me.btnJVSetting.Enabled = True
                Exit Sub
            End If

            'バッファ作成
            strBuff = New String(vbNullChar, lngBuffSize)
            strFileName = New String(vbNullChar, lngFileNameSize)
            Dim recordspec As String

            Do
                Application.DoEvents()

                'JVReadで1行読み込み
                lngReturnCode = JVLink1.JVRead(strBuff, lngBuffSize, strFileName)

                'リターンコードにより処理を分岐
                Select Case lngReturnCode
                    Case 0 ' 全ファイル読み込み終了
                        Exit Do
                    Case -1 ' ファイル切り替わり
                    Case -3 ' ダウンロード中
                    Case -201 ' Initされてない
                        MsgBox("JVInitが行われていません。")
                        Exit Do
                    Case -203 ' Openされてない
                        MsgBox("JVOpenが行われていません。")
                        Exit Do
                    Case -503 ' ファイルがない
                        Exit Do
                    Case Is > 0 ' 正常読み込み
                        recordspec = Mid(strBuff, 1, 2)
                        objDB.SetData(strBuff, lngBuffSize)
                End Select
            Loop While (1)

            '解放
            objDB.Close()
            objDB = Nothing

            'JVClose
            JVLink1.JVClose()

            Me.Cursor = System.Windows.Forms.Cursors.Default
            Me.btnRead.Enabled = True
            Me.btnJVSetting.Enabled = True

            MsgBox("全データの読み込み処理を終了しました")
        Catch ex As Exception
            Me.btnRead.Enabled = True
            Me.btnJVSetting.Enabled = True
            MsgBox(ex.Message)
        End Try

    End Sub

    Private Sub btnJVSetting_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnJVSetting.Click
        If Me.JVLink1.JVSetUIProperties = -1 Then
            MsgBox("エラーのためJV-Linkの設定に失敗しました")
        End If
    End Sub
End Class
