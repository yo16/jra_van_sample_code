VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "馬吉について"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   3030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows の既定値
   Begin VB.PictureBox picLogo 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'なし
      Height          =   3015
      Left            =   0
      ScaleHeight     =   3015
      ScaleWidth      =   3015
      TabIndex        =   1
      Top             =   0
      Width           =   3015
   End
   Begin VB.TextBox txtInfo 
      BorderStyle     =   0  'なし
      Height          =   1815
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  '両方
      TabIndex        =   0
      Text            =   "frmAbout.frx":0000
      Top             =   3000
      Width           =   3045
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'   ヘルプ フォーム
'
Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mJVLink As frmWrappedJVLink


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能： フォームロード処理
'
'   備考： なし
'
Private Sub Form_Load()
On Error GoTo Errorhandler
    Me.Icon = LoadResPicture(100, vbResIcon)

    Call SetTextRelase
    picLogo.Picture = LoadResPicture(1000, vbResBitmap)
    
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   機能： 馬吉について　表示内容の編集
'
'   備考： なし
'
Private Sub SetTextRelase()
On Error GoTo Errorhandler
    Dim strInfo As String
    Dim cn      As New ADODB.Connection
    Dim i       As Long
    
    
    
    strInfo = strInfo & vbCrLf
    strInfo = strInfo & "--- 馬吉 について ---"
    strInfo = strInfo & vbCrLf
    
    strInfo = strInfo & "CompanyName: "
    strInfo = strInfo & App.CompanyName
    strInfo = strInfo & vbCrLf
    strInfo = strInfo & "EXEName: "
    strInfo = strInfo & App.EXEName
    strInfo = strInfo & vbCrLf
    strInfo = strInfo & "LegalCopyright: "
    strInfo = strInfo & App.LegalCopyright
    strInfo = strInfo & vbCrLf
    strInfo = strInfo & "LegalTrademarks: "
    strInfo = strInfo & App.LegalTrademarks
    strInfo = strInfo & vbCrLf
    strInfo = strInfo & "ProductName: "
    strInfo = strInfo & App.ProductName
    strInfo = strInfo & vbCrLf
    strInfo = strInfo & "バージョン: "
    strInfo = strInfo & App.Major & "." & App.Minor & "." & App.Revision
    strInfo = strInfo & vbCrLf
    strInfo = strInfo & "データベースパス: "
    strInfo = strInfo & gApp.R_DBPath
    strInfo = strInfo & vbCrLf
    
    strInfo = strInfo & vbCrLf
    strInfo = strInfo & "--- JV-Link について ---"
    strInfo = strInfo & vbCrLf
    
    Set mJVLink = New frmWrappedJVLink
    Load mJVLink
    If Err.Number = 0 Then
        strInfo = strInfo & "バージョン: "
        strInfo = strInfo & mJVLink.m_JVLinkVersion
        strInfo = strInfo & vbCrLf
        strInfo = strInfo & "サービスキー: "
        strInfo = strInfo & mJVLink.JVlink.m_servicekey
        strInfo = strInfo & vbCrLf
        strInfo = strInfo & "セーブパス: "
        strInfo = strInfo & mJVLink.JVlink.m_savepath
        strInfo = strInfo & vbCrLf
        strInfo = strInfo & "セーブフラグ: "
        strInfo = strInfo & mJVLink.JVlink.m_saveflag
        strInfo = strInfo & vbCrLf
    Else
        strInfo = strInfo & Err.Description & Err.Number
        strInfo = strInfo & vbCrLf
    End If
    Err.Clear
    
    strInfo = strInfo & vbCrLf
    strInfo = strInfo & "--- ADO について ---"
    strInfo = strInfo & vbCrLf
    
    strInfo = strInfo & "バージョン: "
    strInfo = strInfo & cn.Version
    strInfo = strInfo & vbCrLf
    
    txtInfo.Text = strInfo
    Exit Sub
Errorhandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能： フォームリサイズ処理
'
'   備考： なし
'
Private Sub Form_Resize()
On Error GoTo Errorhandler

    txtInfo.Move 0, picLogo.Height, ScaleWidth, ScaleHeight - picLogo.Height
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   機能： フォームアンロード処理
'
'   備考： なし
'
Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Errorhandler
    Unload mJVLink
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub

