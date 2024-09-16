<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class DataImportForm
    Inherits System.Windows.Forms.Form

    'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing AndAlso components IsNot Nothing Then
            components.Dispose()
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Windows フォーム デザイナで必要です。
    Private components As System.ComponentModel.IContainer

    'メモ: 以下のプロシージャは Windows フォーム デザイナで必要です。
    'Windows フォーム デザイナを使用して変更できます。  
    'コード エディタを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(DataImportForm))
        Me.btnRead = New System.Windows.Forms.Button
        Me.txtFromTime = New System.Windows.Forms.TextBox
        Me.txtDataSpec = New System.Windows.Forms.TextBox
        Me.lblFromTime = New System.Windows.Forms.Label
        Me.lblDataSpec = New System.Windows.Forms.Label
        Me.rbtNormal = New System.Windows.Forms.RadioButton
        Me.Frame1 = New System.Windows.Forms.GroupBox
        Me.rbtIsthisweek = New System.Windows.Forms.RadioButton
        Me.rbtSetup = New System.Windows.Forms.RadioButton
        Me.btnJVSetting = New System.Windows.Forms.Button
        Me.JVLink1 = New AxJVDTLabLib.AxJVLink
        Me.Frame1.SuspendLayout()
        CType(Me.JVLink1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'btnRead
        '
        Me.btnRead.Location = New System.Drawing.Point(403, 15)
        Me.btnRead.Name = "btnRead"
        Me.btnRead.Size = New System.Drawing.Size(101, 51)
        Me.btnRead.TabIndex = 0
        Me.btnRead.Text = "取込開始"
        Me.btnRead.UseVisualStyleBackColor = True
        '
        'txtFromTime
        '
        Me.txtFromTime.AcceptsReturn = True
        Me.txtFromTime.BackColor = System.Drawing.SystemColors.Window
        Me.txtFromTime.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtFromTime.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtFromTime.Location = New System.Drawing.Point(253, 31)
        Me.txtFromTime.MaxLength = 0
        Me.txtFromTime.Name = "txtFromTime"
        Me.txtFromTime.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtFromTime.Size = New System.Drawing.Size(97, 19)
        Me.txtFromTime.TabIndex = 8
        Me.txtFromTime.Text = "20111201000000"
        '
        'txtDataSpec
        '
        Me.txtDataSpec.AcceptsReturn = True
        Me.txtDataSpec.BackColor = System.Drawing.SystemColors.Control
        Me.txtDataSpec.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtDataSpec.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtDataSpec.Location = New System.Drawing.Point(7, 31)
        Me.txtDataSpec.MaxLength = 0
        Me.txtDataSpec.Name = "txtDataSpec"
        Me.txtDataSpec.ReadOnly = True
        Me.txtDataSpec.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtDataSpec.Size = New System.Drawing.Size(241, 19)
        Me.txtDataSpec.TabIndex = 7
        Me.txtDataSpec.Text = "SNPN"
        '
        'lblFromTime
        '
        Me.lblFromTime.BackColor = System.Drawing.SystemColors.Control
        Me.lblFromTime.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblFromTime.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblFromTime.Location = New System.Drawing.Point(251, 13)
        Me.lblFromTime.Name = "lblFromTime"
        Me.lblFromTime.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblFromTime.Size = New System.Drawing.Size(73, 17)
        Me.lblFromTime.TabIndex = 6
        Me.lblFromTime.Text = "FromTime"
        '
        'lblDataSpec
        '
        Me.lblDataSpec.BackColor = System.Drawing.SystemColors.Control
        Me.lblDataSpec.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblDataSpec.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblDataSpec.Location = New System.Drawing.Point(7, 13)
        Me.lblDataSpec.Name = "lblDataSpec"
        Me.lblDataSpec.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblDataSpec.Size = New System.Drawing.Size(73, 17)
        Me.lblDataSpec.TabIndex = 9
        Me.lblDataSpec.Text = "データ種別"
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
        Me.rbtNormal.Text = "通常データ"
        Me.rbtNormal.UseVisualStyleBackColor = False
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.SystemColors.Control
        Me.Frame1.Controls.Add(Me.rbtNormal)
        Me.Frame1.Controls.Add(Me.rbtIsthisweek)
        Me.Frame1.Controls.Add(Me.rbtSetup)
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(7, 73)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(353, 49)
        Me.Frame1.TabIndex = 10
        Me.Frame1.TabStop = False
        Me.Frame1.Text = "取得データ"
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
        Me.rbtIsthisweek.Text = "今週開催データ"
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
        Me.rbtSetup.Text = "セットアップデータ"
        Me.rbtSetup.UseVisualStyleBackColor = False
        '
        'btnJVSetting
        '
        Me.btnJVSetting.BackColor = System.Drawing.SystemColors.Control
        Me.btnJVSetting.Cursor = System.Windows.Forms.Cursors.Default
        Me.btnJVSetting.ForeColor = System.Drawing.SystemColors.ControlText
        Me.btnJVSetting.Location = New System.Drawing.Point(403, 85)
        Me.btnJVSetting.Name = "btnJVSetting"
        Me.btnJVSetting.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.btnJVSetting.Size = New System.Drawing.Size(101, 37)
        Me.btnJVSetting.TabIndex = 11
        Me.btnJVSetting.Text = "JVLink設定"
        Me.btnJVSetting.UseVisualStyleBackColor = False
        '
        'JVLink1
        '
        Me.JVLink1.Enabled = True
        Me.JVLink1.Location = New System.Drawing.Point(368, 125)
        Me.JVLink1.Name = "JVLink1"
        Me.JVLink1.OcxState = CType(resources.GetObject("JVLink1.OcxState"), System.Windows.Forms.AxHost.State)
        Me.JVLink1.Size = New System.Drawing.Size(192, 192)
        Me.JVLink1.TabIndex = 12
        '
        'DataImportForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(516, 160)
        Me.Controls.Add(Me.JVLink1)
        Me.Controls.Add(Me.btnJVSetting)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.txtFromTime)
        Me.Controls.Add(Me.txtDataSpec)
        Me.Controls.Add(Me.lblFromTime)
        Me.Controls.Add(Me.lblDataSpec)
        Me.Controls.Add(Me.btnRead)
        Me.Name = "DataImportForm"
        Me.Text = "出走別着度数データ読み込み"
        Me.Frame1.ResumeLayout(False)
        CType(Me.JVLink1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents btnRead As System.Windows.Forms.Button
    Public WithEvents txtFromTime As System.Windows.Forms.TextBox
    Public WithEvents txtDataSpec As System.Windows.Forms.TextBox
    Public WithEvents lblFromTime As System.Windows.Forms.Label
    Public WithEvents lblDataSpec As System.Windows.Forms.Label
    Public WithEvents rbtNormal As System.Windows.Forms.RadioButton
    Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents rbtIsthisweek As System.Windows.Forms.RadioButton
    Public WithEvents rbtSetup As System.Windows.Forms.RadioButton
    Public WithEvents btnJVSetting As System.Windows.Forms.Button
    Friend WithEvents JVLink1 As AxJVDTLabLib.AxJVLink

End Class
