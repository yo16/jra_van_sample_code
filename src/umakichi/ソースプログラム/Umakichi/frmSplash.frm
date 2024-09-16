VERSION 5.00
Begin VB.Form frmSplash 
   Appearance      =   0  'ﾌﾗｯﾄ
   BackColor       =   &H00E0EEEE&
   BorderStyle     =   0  'なし
   Caption         =   "Form1"
   ClientHeight    =   3180
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3180
   ScaleWidth      =   4635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '画面の中央
   Begin VB.Timer tmrAnim 
      Interval        =   1
      Left            =   3810
      Top             =   2190
   End
   Begin VB.Frame fraTop 
      Appearance      =   0  'ﾌﾗｯﾄ
      BackColor       =   &H00C0CCCC&
      BorderStyle     =   0  'なし
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   210
      TabIndex        =   0
      Top             =   240
      Width           =   4215
      Begin VB.Label lblFix 
         Alignment       =   2  '中央揃え
         Appearance      =   0  'ﾌﾗｯﾄ
         AutoSize        =   -1  'True
         BackColor       =   &H00C0CCCC&
         Caption         =   "馬吉オープンソース版"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   915
         TabIndex        =   1
         Top             =   150
         Width           =   2385
      End
   End
   Begin VB.Label lblVersion 
      Appearance      =   0  'ﾌﾗｯﾄ
      AutoSize        =   -1  'True
      BackColor       =   &H00E0EEEE&
      Caption         =   "Ver.0.0"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   2070
      TabIndex        =   2
      Top             =   1650
      Width           =   510
   End
   Begin VB.Shape shpObj 
      BackColor       =   &H00C0CCCC&
      BackStyle       =   1  '不透明
      BorderColor     =   &H00C0CCCC&
      Height          =   225
      Left            =   2250
      Shape           =   3  '円
      Top             =   2370
      Width           =   195
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'   スプラッシュウインドウ
'
'   馬吉の起動時に、著作権表記などを表示する
'   主にデータベースのチェック中の待ち時間に表示
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private w As Long
Private h As Long

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: フォームを削除
'
'   備考: なし
'
Public Sub kill()
    Unload Me
End Sub


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: フォームロードイベント
'
'   備考: なし
'
Private Sub Form_Load()
On Error GoTo ErrorHandler
    Dim color As Long
    
    w = ScaleWidth
    h = ScaleHeight
    With App
    lblVersion.Caption = .Major & "." & .Minor & "." & .Revision
    End With
    ' 起動前で gApp が使えない為、レジストリを直接読み込みます
    color = CLng(val(GetRegData(HKEY_CURRENT_USER, "Software\" & cRegistrySubKey & "\" & "Color", "BackColorDark")))
    If color <> 0 Then
        fraTop.BackColor = color
    End If
    lblFix.BackColor = fraTop.BackColor
    lblFix.ForeColor = Contrast(fraTop.BackColor)
    shpObj.BackColor = fraTop.BackColor
    
    ' 起動前で gApp が使えない為、レジストリを直接読み込みます
    color = CLng(val(GetRegData(HKEY_CURRENT_USER, "Software\" & cRegistrySubKey & "\" & "Color", "BackColorLight")))
    If color <> 0 Then
        Me.BackColor = color
    End If
    lblVersion.BackColor = Me.BackColor
    lblVersion.ForeColor = Contrast(Me.BackColor)
    
    fraTop.Left = (w / 2) - (fraTop.width / 2)
    lblVersion.Left = (w / 2) - (lblVersion.width / 2)
    lblFix.Left = (fraTop.width / 2) - (lblFix.width / 2)
    Exit Sub
ErrorHandler:
    Resume Next
End Sub


'
'   機能: アニメーションタイマーイベント
'
'   備考: なし
'
Private Sub tmrAnim_Timer()
On Error GoTo ErrorHandler
    With shpObj
        .Left = (w / 2) - (.width / 2) + (Sin((Timer) * 3) * w / 5)
        .Top = (h / 2) - (.Height / 2) + (Sin((Timer) * 1) * h / 5)
    End With
    Me.Refresh
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub
