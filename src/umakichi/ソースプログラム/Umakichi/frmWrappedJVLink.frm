VERSION 5.00
Object = "{2AB17740-0C41-11D7-916F-0003479BEB3F}#1.18#0"; "JVDTLab.dll"
Begin VB.Form frmWrappedJVLink 
   BorderStyle     =   3  '�Œ��޲�۸�
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
   StartUpPosition =   1  '��Ű ̫�т̒���
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
'   JVLink���C���X�g�[������Ă��邩�ǂ����̔���ׂ̈̉B���t�H�[��
'
'   Visible=False�Ń��[�U�[����͉B������܂��B
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �v���p�e�B
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   �@�\: JVLink�I�u�W�F�N�g��Ԃ�
'
'   ���l: �Ȃ�
'
Public Property Get JVlink() As JVlink
    Set JVlink = axJVLink
End Property


'
'   �@�\: JVLInk�I�u�W�F�N�g�̃o�[�W������Ԃ�
'
'   ���l: �Ȃ�
'
Public Property Get m_JVLinkVersion() As String
    m_JVLinkVersion = axJVLink.m_JVLinkVersion
End Property
