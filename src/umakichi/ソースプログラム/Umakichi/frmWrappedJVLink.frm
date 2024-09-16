VERSION 5.00
Object = "{2AB17740-0C41-11D7-916F-0003479BEB3F}#1.18#0"; "JVDTLab.dll"
Begin VB.Form frmWrappedJVLink 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmWrappedJVLink.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'ｵｰﾅｰ ﾌｫｰﾑの中央
   Visible         =   0   'False
   Begin JVDTLabLibCtl.JVLink axJVLink 
      Left            =   990
      OleObjectBlob   =   "frmWrappedJVLink.frx":628A
      Top             =   570
   End
End
Attribute VB_Name = "frmWrappedJVLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'   JVLinkがインストールされているかどうかの判定の為の隠しフォーム
'
'   Visible=Falseでユーザーからは隠匿されます。
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   プロパティ
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: JVLinkオブジェクトを返す
'
'   備考: なし
'
Public Property Get JVlink() As JVlink
    Set JVlink = axJVLink
End Property


'
'   機能: JVLInkオブジェクトのバージョンを返す
'
'   備考: なし
'
Public Property Get m_JVLinkVersion() As String
    m_JVLinkVersion = axJVLink.m_JVLinkVersion
End Property
