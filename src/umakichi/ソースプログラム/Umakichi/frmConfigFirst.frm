VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmConfigFirst 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "データセットアップの設定"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7470
   Icon            =   "frmConfigFirst.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows の既定値
   Begin VB.CommandButton cmdCancel 
      Caption         =   "キャンセル(&C)"
      Default         =   -1  'True
      Height          =   375
      Left            =   4050
      TabIndex        =   2
      Top             =   4320
      Width           =   1965
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "データセットアップ開始(&S)"
      Height          =   375
      Left            =   1860
      TabIndex        =   1
      Top             =   4320
      Width           =   2145
   End
   Begin VB.Frame frmJVLMode 
      Caption         =   "JV-Link 取得モード"
      Height          =   4245
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   7395
      Begin VB.PictureBox picXPTheme 
         BorderStyle     =   0  'なし
         Height          =   4005
         Left            =   60
         ScaleHeight     =   4005
         ScaleWidth      =   7275
         TabIndex        =   3
         Top             =   180
         Width           =   7275
         Begin VB.OptionButton optJVMode 
            Caption         =   "今週モード"
            Height          =   240
            Index           =   1
            Left            =   0
            TabIndex        =   7
            Top             =   2160
            Width           =   1305
         End
         Begin MSComCtl2.UpDown updYear 
            Height          =   315
            Left            =   3240
            TabIndex        =   5
            Top             =   480
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   556
            _Version        =   393216
            OrigLeft        =   1260
            OrigTop         =   750
            OrigRight       =   1410
            OrigBottom      =   1065
            Max             =   3000
            Enabled         =   -1  'True
         End
         Begin VB.CheckBox chkBLOD 
            Caption         =   "産駒・繁殖馬を含める"
            Height          =   225
            Left            =   2220
            TabIndex        =   4
            Top             =   240
            Value           =   1  'ﾁｪｯｸ
            Width           =   2175
         End
         Begin VB.CheckBox chkSLOP 
            Caption         =   "坂路調教を含める"
            Height          =   255
            Left            =   390
            TabIndex        =   12
            Top             =   210
            Value           =   1  'ﾁｪｯｸ
            Width           =   1995
         End
         Begin VB.TextBox txtFix 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'なし
            BeginProperty Font 
               Name            =   "MS UI Gothic"
               Size            =   8.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   2
            Left            =   120
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   11
            Text            =   "frmConfigFirst.frx":000C
            Top             =   3450
            Width           =   5565
         End
         Begin VB.TextBox txtFix 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'なし
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   8.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   885
            Index           =   1
            Left            =   360
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   10
            Text            =   "frmConfigFirst.frx":0097
            Top             =   2430
            Width           =   6225
         End
         Begin VB.OptionButton optJVMode 
            Caption         =   "通常モード"
            Height          =   225
            Index           =   0
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Value           =   -1  'True
            Width           =   1425
         End
         Begin VB.TextBox txtFix 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'なし
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   8.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1365
            Index           =   0
            Left            =   360
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   8
            Text            =   "frmConfigFirst.frx":01C6
            Top             =   870
            Width           =   6765
         End
         Begin VB.TextBox txtYear 
            Alignment       =   1  '右揃え
            Height          =   285
            Left            =   2670
            TabIndex        =   6
            Text            =   "1995"
            Top             =   480
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "指定した年度以降のデータのみをセットアップします。"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   8.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   190
            Left            =   3540
            TabIndex        =   14
            Top             =   560
            Width           =   3705
         End
         Begin VB.Label lblFix 
            AutoSize        =   -1  'True
            Caption         =   "セットアップ開始年度："
            Height          =   180
            Index           =   4
            Left            =   960
            TabIndex        =   13
            Top             =   510
            Width           =   1695
         End
      End
   End
End
Attribute VB_Name = "frmConfigFirst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'   (初)データ取得設定 ダイアログ
'
Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mstrButtonType As String   ' 押されたボタンのタイプ

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   プロパティ
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: ボタンタイプを返す
'
'   備考: なし
'
Public Property Get ButtonType() As String
    ButtonType = mstrButtonType
End Property


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: キャンセルボタン選択イベント
'
'   備考: なし
'
Private Sub cmdCancel_Click()
On Error GoTo Errorhandler
    mstrButtonType = "Cancel"
    Unload Me
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   機能: ＯＫボタン選択イベント
'
'   備考: なし
'
Private Sub cmdOK_Click()
On Error GoTo Errorhandler
    gApp.R_JVLGetSLOP = (chkSLOP.value = 1)
    gApp.R_JVLGetBLOD = (chkBLOD.value = 1)
    If optJVMode(0).value = True Then
        gApp.R_JVLMode = ukjUsual
    Else
        gApp.R_JVLMode = ukjThisWeek
    End If
    gApp.R_SetupYear = Format$(val(txtYear.Text), "0000")
    mstrButtonType = "OK"
    
    Unload Me
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   機能: フォームロードイベント
'
'   備考: なし
'
Private Sub Form_Load()
On Error GoTo Errorhandler
    Me.Icon = LoadResPicture(100, vbResIcon)
    optJVMode(0).value = (gApp.R_JVLMode = ukjUsual)
    optJVMode(1).value = (gApp.R_JVLMode = ukjThisWeek)
    chkSLOP.value = IIf(gApp.R_JVLGetSLOP, 1, 0)
    chkBLOD.value = IIf(gApp.R_JVLGetBLOD, 1, 0)
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   機能: JV-Link取得モード選択イベント
'
'   備考: なし
'
Private Sub optJVMode_Click(Index As Integer)
On Error GoTo Errorhandler
    chkSLOP.Enabled = (Index = 0)
    chkBLOD.Enabled = (Index = 0)
    txtYear.Enabled = (Index = 0)
    updYear.Enabled = (Index = 0)
    cmdOK.Caption = IIf(Index = 0, "データセットアップ開始", "取得開始")
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   機能: セットアップ開始年度キー入力イベント
'
'   備考: なし
'
Private Sub txtYear_KeyPress(KeyAscii As Integer)
On Error GoTo Errorhandler
   If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
      KeyAscii = 0      ' 文字を取り消します。
      Beep            ' エラー音を鳴らします。
   End If
   Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   機能: セットアップ開始年度ロストフォーカスイベント
'
'   備考: なし
'
Private Sub txtYear_LostFocus()
On Error GoTo Errorhandler
    If val(txtYear.Text) > val(Year(Now)) Then  '当年より大きいとき当年に置換する
        txtYear.Text = Year(Now)
    ElseIf val(txtYear.Text) <= 1995 Then       '1995年以前のとき1995年に置換する
        txtYear.Text = "1995"
    End If
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub

'
'   機能: セットアップ開始年度バリデイトイベント
'
'   備考: なし
'
Private Sub txtYear_Validate(Cancel As Boolean)
On Error GoTo Errorhandler
    If val(txtYear.Text) > val(Year(Now)) Then
        txtYear.Text = Year(Now)
    ElseIf val(txtYear.Text) <= 1995 Then
        txtYear.Text = "1995"
    End If
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   機能: セットアップ開始年度ダウンクリックイベント
'
'   備考: なし
'
Private Sub updYear_DownClick()
On Error GoTo Errorhandler
    If txtYear.Text > 1995 Then
        txtYear.Text = txtYear.Text - 1
    End If
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   機能: セットアップ開始年度アップクリックイベント
'
'   備考: なし
'
Private Sub updYear_UpClick()
On Error GoTo Errorhandler
    If val(txtYear.Text) < CInt(Year(Now)) Then
        txtYear.Text = txtYear.Text + 1
    End If
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub
