VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.UserControl ctlVOD 
   BackColor       =   &H00E0EEEE&
   ClientHeight    =   7785
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10560
   ScaleHeight     =   7785
   ScaleWidth      =   10560
   Begin VB.Timer tmrTBS 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   8400
      Top             =   1560
   End
   Begin VB.Timer tmrTrigger 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   9000
      Top             =   1560
   End
   Begin VB.Frame fraHeader 
      BackColor       =   &H00E0EEEE&
      BorderStyle     =   0  '�Ȃ�
      Caption         =   "fraHeader"
      Height          =   585
      Left            =   240
      TabIndex        =   42
      Top             =   840
      Width           =   8655
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0EEEE&
         Caption         =   "."
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   8
         Left            =   2400
         TabIndex        =   79
         Top             =   360
         Width           =   105
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0EEEE&
         Caption         =   "."
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   2
         Left            =   60
         TabIndex        =   48
         Top             =   30
         Width           =   90
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0EEEE&
         Caption         =   "."
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   3
         Left            =   1500
         TabIndex        =   47
         Top             =   30
         Width           =   90
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0EEEE&
         Caption         =   "."
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   4
         Left            =   1500
         TabIndex        =   46
         Top             =   210
         Width           =   90
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0EEEE&
         Caption         =   "."
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   5
         Left            =   1500
         TabIndex        =   45
         Top             =   360
         Width           =   90
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0EEEE&
         Caption         =   "."
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   6
         Left            =   3570
         TabIndex        =   44
         Top             =   210
         Width           =   90
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0EEEE&
         Caption         =   "."
         BeginProperty Font 
            Name            =   "�l�r �S�V�b�N"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   7
         Left            =   2040
         TabIndex        =   43
         Top             =   360
         Width           =   90
      End
   End
   Begin MSComctlLib.ImageList ilsTbrSmall 
      Left            =   480
      Top             =   5430
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   12
      ImageHeight     =   12
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Frame fraTop 
      BackColor       =   &H00C0CCCC&
      BorderStyle     =   0  '�Ȃ�
      Caption         =   "Frame1"
      Height          =   465
      Left            =   120
      TabIndex        =   1
      Top             =   180
      Width           =   8955
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "�f�[�^�F 1�� 1��16��10��"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   1
         Left            =   6870
         TabIndex        =   4
         Top             =   150
         Visible         =   0   'False
         Width           =   1965
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0CCCC&
         Caption         =   "���\�[�X�s���ł��B�s�v�ȉ�ʂ���Ă�������"
         BeginProperty Font 
            Name            =   "�l�r �o�S�V�b�N"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   210
         TabIndex        =   2
         Top             =   120
         Width           =   5010
      End
   End
   Begin TabDlg.SSTab mstTab 
      Height          =   3105
      Left            =   150
      TabIndex        =   0
      Top             =   1800
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   5477
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      BackColor       =   14741230
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "�P���������g�A"
      TabPicture(0)   =   "ctlVOD.ctx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "paneTab(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "�@�@�n�A�@�@"
      TabPicture(1)   =   "ctlVOD.ctx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "paneTab(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "�@�@���C�h�@�@"
      TabPicture(2)   =   "ctlVOD.ctx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "paneTab(2)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "�@�@�n�P�@�@"
      TabPicture(3)   =   "ctlVOD.ctx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "paneTab(3)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "�@�@�R�A���@�@"
      TabPicture(4)   =   "ctlVOD.ctx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "paneTab(4)"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "�@�@�R�A�P�@�@"
      TabPicture(5)   =   "ctlVOD.ctx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "paneTab(5)"
      Tab(5).ControlCount=   1
      Begin Umakichi.ctlPane paneTab 
         Height          =   2415
         Index           =   1
         Left            =   -74940
         TabIndex        =   9
         Top             =   360
         Width           =   7515
         _ExtentX        =   13256
         _ExtentY        =   4260
         Begin Umakichi.ctlWrappedGrid flexTab 
            Height          =   1455
            Index           =   1
            Left            =   0
            TabIndex        =   51
            Top             =   210
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   2566
         End
         Begin VB.Label lblHyo 
            AutoSize        =   -1  'True
            Caption         =   "lblHyo3"
            Height          =   180
            Index           =   3
            Left            =   1410
            TabIndex        =   13
            Top             =   0
            Visible         =   0   'False
            Width           =   570
         End
      End
      Begin Umakichi.ctlPane paneTab 
         Height          =   2295
         Index           =   0
         Left            =   60
         TabIndex        =   5
         Top             =   360
         Width           =   6525
         _ExtentX        =   11509
         _ExtentY        =   4048
         Begin Umakichi.ctlWrappedGrid flexTab 
            Height          =   1455
            Index           =   0
            Left            =   0
            TabIndex        =   50
            Top             =   210
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   2566
         End
         Begin VB.Label lblHyo 
            AutoSize        =   -1  'True
            Caption         =   "lblHyo0"
            Height          =   180
            Index           =   0
            Left            =   570
            TabIndex        =   8
            Top             =   0
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.Label lblHyo 
            AutoSize        =   -1  'True
            Caption         =   "lblHyo1"
            Height          =   180
            Index           =   1
            Left            =   1290
            TabIndex        =   7
            Top             =   0
            Visible         =   0   'False
            Width           =   570
         End
         Begin VB.Label lblHyo 
            AutoSize        =   -1  'True
            Caption         =   "lblHyo2"
            Height          =   180
            Index           =   2
            Left            =   2010
            TabIndex        =   6
            Top             =   0
            Visible         =   0   'False
            Width           =   570
         End
      End
      Begin Umakichi.ctlPane paneTab 
         Height          =   2385
         Index           =   2
         Left            =   -74940
         TabIndex        =   10
         Top             =   360
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   4207
         Begin Umakichi.ctlWrappedGrid flexTab 
            Height          =   1455
            Index           =   2
            Left            =   0
            TabIndex        =   52
            Top             =   210
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   2566
         End
         Begin VB.Label lblHyo 
            AutoSize        =   -1  'True
            Caption         =   "lblHyo4"
            Height          =   180
            Index           =   4
            Left            =   1770
            TabIndex        =   14
            Top             =   0
            Visible         =   0   'False
            Width           =   570
         End
      End
      Begin Umakichi.ctlPane paneTab 
         Height          =   2415
         Index           =   3
         Left            =   -74940
         TabIndex        =   11
         Top             =   360
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   4260
         Begin Umakichi.ctlWrappedGrid flexTab 
            Height          =   1455
            Index           =   3
            Left            =   0
            TabIndex        =   53
            Top             =   210
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   2566
         End
         Begin VB.Label lblHyo 
            AutoSize        =   -1  'True
            Caption         =   "lblHyo5"
            Height          =   180
            Index           =   5
            Left            =   1890
            TabIndex        =   15
            Top             =   0
            Visible         =   0   'False
            Width           =   570
         End
      End
      Begin Umakichi.ctlPane paneTab 
         Height          =   2355
         Index           =   4
         Left            =   -74940
         TabIndex        =   12
         Top             =   360
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   4154
         Begin Umakichi.ctlWrappedGrid flexTab 
            Height          =   1455
            Index           =   4
            Left            =   0
            TabIndex        =   49
            Top             =   300
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   2566
         End
         Begin VB.Frame fraOptAxis 
            BorderStyle     =   0  '�Ȃ�
            Caption         =   "Frame2"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   16
            Top             =   150
            Width           =   4905
            Begin VB.OptionButton optAxisNinki 
               Caption         =   "601�`"
               Height          =   255
               Index           =   3
               Left            =   3600
               Style           =   1  '���̨���
               TabIndex        =   20
               Top             =   0
               Width           =   1200
            End
            Begin VB.OptionButton optAxisNinki 
               Caption         =   "401�`"
               Height          =   255
               Index           =   2
               Left            =   2400
               Style           =   1  '���̨���
               TabIndex        =   19
               Top             =   0
               Width           =   1200
            End
            Begin VB.OptionButton optAxisNinki 
               Caption         =   "201�`"
               Height          =   255
               Index           =   1
               Left            =   1200
               Style           =   1  '���̨���
               TabIndex        =   18
               Top             =   0
               Width           =   1200
            End
            Begin VB.OptionButton optAxisNinki 
               Caption         =   "1�`"
               Height          =   255
               Index           =   0
               Left            =   0
               Style           =   1  '���̨���
               TabIndex        =   17
               Top             =   0
               Value           =   -1  'True
               Width           =   1200
            End
         End
         Begin VB.Frame fraOptAxis 
            BorderStyle     =   0  '�Ȃ�
            Caption         =   "Frame1"
            Height          =   255
            Index           =   0
            Left            =   0
            TabIndex        =   21
            Top             =   60
            Width           =   6495
            Begin VB.OptionButton optAxisNum 
               Caption         =   "18"
               Height          =   255
               Index           =   17
               Left            =   6090
               Style           =   1  '���̨���
               TabIndex        =   39
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNum 
               Caption         =   "17"
               Height          =   255
               Index           =   16
               Left            =   5760
               Style           =   1  '���̨���
               TabIndex        =   38
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNum 
               Caption         =   "16"
               Height          =   255
               Index           =   15
               Left            =   5430
               Style           =   1  '���̨���
               TabIndex        =   37
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNum 
               Caption         =   "15"
               Height          =   255
               Index           =   14
               Left            =   5100
               Style           =   1  '���̨���
               TabIndex        =   36
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNum 
               Caption         =   "14"
               Height          =   255
               Index           =   13
               Left            =   4770
               Style           =   1  '���̨���
               TabIndex        =   35
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNum 
               Caption         =   "13"
               Height          =   255
               Index           =   12
               Left            =   4440
               Style           =   1  '���̨���
               TabIndex        =   34
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNum 
               Caption         =   "12"
               Height          =   255
               Index           =   11
               Left            =   4110
               Style           =   1  '���̨���
               TabIndex        =   33
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNum 
               Caption         =   "11"
               Height          =   255
               Index           =   10
               Left            =   3780
               Style           =   1  '���̨���
               TabIndex        =   32
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNum 
               Caption         =   "10"
               Height          =   255
               Index           =   9
               Left            =   3450
               Style           =   1  '���̨���
               TabIndex        =   31
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNum 
               Caption         =   "9"
               Height          =   255
               Index           =   8
               Left            =   3120
               Style           =   1  '���̨���
               TabIndex        =   30
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNum 
               Caption         =   "8"
               Height          =   255
               Index           =   7
               Left            =   2790
               Style           =   1  '���̨���
               TabIndex        =   29
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNum 
               Caption         =   "7"
               Height          =   255
               Index           =   6
               Left            =   2460
               Style           =   1  '���̨���
               TabIndex        =   28
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNum 
               Caption         =   "6"
               Height          =   255
               Index           =   5
               Left            =   2130
               Style           =   1  '���̨���
               TabIndex        =   27
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNum 
               Caption         =   "5"
               Height          =   255
               Index           =   4
               Left            =   1800
               Style           =   1  '���̨���
               TabIndex        =   26
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNum 
               Caption         =   "4"
               Height          =   255
               Index           =   3
               Left            =   1470
               Style           =   1  '���̨���
               TabIndex        =   25
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNum 
               Caption         =   "3"
               Height          =   255
               Index           =   2
               Left            =   1140
               Style           =   1  '���̨���
               TabIndex        =   24
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNum 
               Caption         =   "2"
               Height          =   255
               Index           =   1
               Left            =   810
               Style           =   1  '���̨���
               TabIndex        =   23
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNum 
               Caption         =   "1"
               Height          =   255
               Index           =   0
               Left            =   480
               Style           =   1  '���̨���
               TabIndex        =   22
               Top             =   0
               Value           =   -1  'True
               Width           =   315
            End
            Begin VB.Label lblFixJiku 
               AutoSize        =   -1  'True
               Caption         =   "���n"
               BeginProperty Font 
                  Name            =   "�l�r �o�S�V�b�N"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   0
               TabIndex        =   40
               Top             =   15
               Width           =   450
            End
         End
         Begin VB.Label lblHyo 
            AutoSize        =   -1  'True
            Caption         =   "lblHyo6"
            Height          =   180
            Index           =   6
            Left            =   6690
            TabIndex        =   41
            Top             =   0
            Visible         =   0   'False
            Width           =   570
         End
      End
      Begin Umakichi.ctlPane paneTab 
         Height          =   2355
         Index           =   5
         Left            =   -74940
         TabIndex        =   54
         Top             =   360
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   4154
         Begin Umakichi.ctlWrappedGrid flexTab 
            Height          =   1455
            Index           =   5
            Left            =   0
            TabIndex        =   77
            Top             =   300
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   2566
         End
         Begin VB.Frame fraOptAxis 
            BorderStyle     =   0  '�Ȃ�
            Caption         =   "Frame2"
            Height          =   300
            Index           =   3
            Left            =   120
            TabIndex        =   75
            Top             =   120
            Width           =   1785
            Begin VB.ComboBox cmbAxisNinki 
               BeginProperty Font 
                  Name            =   "�l�r �S�V�b�N"
                  Size            =   9
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               ItemData        =   "ctlVOD.ctx":00A8
               Left            =   60
               List            =   "ctlVOD.ctx":00C4
               Style           =   2  '��ۯ���޳� ؽ�
               TabIndex        =   76
               Top             =   0
               Width           =   1215
            End
         End
         Begin VB.Frame fraOptAxis 
            BorderStyle     =   0  '�Ȃ�
            Caption         =   "Frame1"
            Height          =   255
            Index           =   2
            Left            =   0
            TabIndex        =   55
            Top             =   0
            Width           =   6495
            Begin VB.OptionButton optAxisNumST 
               Caption         =   "18"
               Height          =   255
               Index           =   17
               Left            =   6090
               Style           =   1  '���̨���
               TabIndex        =   73
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNumST 
               Caption         =   "17"
               Height          =   255
               Index           =   16
               Left            =   5760
               Style           =   1  '���̨���
               TabIndex        =   72
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNumST 
               Caption         =   "16"
               Height          =   255
               Index           =   15
               Left            =   5430
               Style           =   1  '���̨���
               TabIndex        =   71
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNumST 
               Caption         =   "15"
               Height          =   255
               Index           =   14
               Left            =   5100
               Style           =   1  '���̨���
               TabIndex        =   70
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNumST 
               Caption         =   "14"
               Height          =   255
               Index           =   13
               Left            =   4770
               Style           =   1  '���̨���
               TabIndex        =   69
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNumST 
               Caption         =   "13"
               Height          =   255
               Index           =   12
               Left            =   4440
               Style           =   1  '���̨���
               TabIndex        =   68
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNumST 
               Caption         =   "12"
               Height          =   255
               Index           =   11
               Left            =   4110
               Style           =   1  '���̨���
               TabIndex        =   67
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNumST 
               Caption         =   "11"
               Height          =   255
               Index           =   10
               Left            =   3780
               Style           =   1  '���̨���
               TabIndex        =   66
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNumST 
               Caption         =   "10"
               Height          =   255
               Index           =   9
               Left            =   3450
               Style           =   1  '���̨���
               TabIndex        =   65
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNumST 
               Caption         =   "9"
               Height          =   255
               Index           =   8
               Left            =   3120
               Style           =   1  '���̨���
               TabIndex        =   64
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNumST 
               Caption         =   "8"
               Height          =   255
               Index           =   7
               Left            =   2790
               Style           =   1  '���̨���
               TabIndex        =   63
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNumST 
               Caption         =   "7"
               Height          =   255
               Index           =   6
               Left            =   2460
               Style           =   1  '���̨���
               TabIndex        =   62
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNumST 
               Caption         =   "6"
               Height          =   255
               Index           =   5
               Left            =   2130
               Style           =   1  '���̨���
               TabIndex        =   61
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNumST 
               Caption         =   "5"
               Height          =   255
               Index           =   4
               Left            =   1800
               Style           =   1  '���̨���
               TabIndex        =   60
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNumST 
               Caption         =   "4"
               Height          =   255
               Index           =   3
               Left            =   1470
               Style           =   1  '���̨���
               TabIndex        =   59
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNumST 
               Caption         =   "3"
               Height          =   255
               Index           =   2
               Left            =   1140
               Style           =   1  '���̨���
               TabIndex        =   58
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNumST 
               Caption         =   "2"
               Height          =   255
               Index           =   1
               Left            =   810
               Style           =   1  '���̨���
               TabIndex        =   57
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNumST 
               Caption         =   "1"
               Height          =   255
               Index           =   0
               Left            =   480
               Style           =   1  '���̨���
               TabIndex        =   56
               Top             =   0
               Value           =   -1  'True
               Width           =   315
            End
            Begin VB.Label lblFixJikuST 
               AutoSize        =   -1  'True
               Caption         =   "���n"
               BeginProperty Font 
                  Name            =   "�l�r �o�S�V�b�N"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   0
               TabIndex        =   74
               Top             =   15
               Width           =   450
            End
         End
         Begin VB.Label lblHyo 
            AutoSize        =   -1  'True
            Caption         =   "lblHyo7"
            Height          =   180
            Index           =   7
            Left            =   6690
            TabIndex        =   78
            Top             =   0
            Visible         =   0   'False
            Width           =   570
         End
      End
   End
   Begin VB.Label lblMakeDate 
      BackColor       =   &H00E0EEEE&
      Caption         =   "."
      Height          =   135
      Left            =   6780
      TabIndex        =   3
      Top             =   60
      Width           =   1935
   End
End
Attribute VB_Name = "ctlVOD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   �I�b�Y��[�� �\���R���g���[��
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �O���֐�(�C�x���g)
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Public Event ChangeTo(strViewerName As String, strKey As String)        '' Vierer�ύX�C�x���g
Public Event WindowTitle(strKey As String)                              '' �E�C���h�E�^�C�g���ύX�C�x���g
Public Event LinkContextMenu(strViewerName As String, strKey As String) '' �E�N���b�N���j���[�\���C�x���g
Public Event Reload()                                                   '' �ēǂݍ���

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �����ϐ�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mVB As clsViewerBase            '' Viewer Base
Private mViewerState As clsVSOdds       '' Viewer State
Private WithEvents mToolBar As ctlToolBars  '' Viewer ToolBar
Attribute mToolBar.VB_VarHelpID = -1

Private WithEvents mData As clsDataOD   '' �f�[�^�擾�I�u�W�F�N�g
Attribute mData.VB_VarHelpID = -1
Private mstrTitle As String             '' �E�C���h�E�^�C�g��
Private mKey      As clsKeyRA           '' �L�[
Private mblnNoData As Boolean           '' �f�[�^�����t���O

Private mblnOddsHyosuIsOdds As Boolean  '' ���W�I�{�^���E�I�b�Y�[�����A�I�b�Y�ł���ΐ^
Private mblnNumNinkiIsNum As Boolean    '' ���W�I�{�^���E�ԍ��l�C���A�ԍ��ł���ΐ^
Private mblnAction As Boolean           '' ���W�I�{�^���E�ԍ��l�C���A���쒆�ł���ΐ^
Private mblnFetchedAll As Boolean       '' �擾�I�����Ă���ΐ^

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �����萔
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

' ��ʍŏ����l
Const MINIMUMWIDTH  As Long = 10000
Const MINIMUMHEIGHT As Long = 5000


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �v���p�e�B
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8


'
'   �@�\: �L�[�ݒ�v���p�e�B
'
'   ���l: Viewer�K�{�v���p�e�B
'
Public Property Let Key(strKey As String)
    gApp.Log "OD: " & strKey
    mKey.str = strKey
    mViewerState.OddsHyosuIsOdds = (Right$(strKey, 1) = "0")
        
    Call Update
End Property


'
'   �@�\: �^�C�g���擾�v���p�e�B
'
'   ���l: Viewer�K�{�v���p�e�B�A�@Browser ���Q��
'
Public Property Get Title() As String
    Title = mstrTitle
End Property


'
'   �@�\: �^�C�g���ݒ�v���p�e�B
'
'   ���l: �u���E�U�ɕύX�ʒm�̃C�x���g����
'
Public Property Let Title(strTitle As String)
    mstrTitle = strTitle
    RaiseEvent WindowTitle(mstrTitle)
End Property


'
'   �@�\: �c�[���o�[��ݒ肷��
'
'   ���l: �u���E�U����c�[���o�[���󂯎��A�c�[���o�[���Z�b�g�A�b�v����
'
Public Property Set ToolBar(RHS As ctlToolBars)
On Error GoTo ErrorHandler
    Dim p   As Long
    
    Set mToolBar = RHS

    With mToolBar.ToolBar(1)
        .Buttons.Clear
        .ImageList = ilsTbrSmall
        
        p = 1
        .Buttons.Add p, "", "", tbrSeparator, 1
        p = p + 1
        .Buttons.Add p, "NUM", "�n�ԏ�", tbrButtonGroup, 2
        .Buttons.item(p).value = IIf(mblnNumNinkiIsNum, tbrPressed, tbrUnpressed)
        p = p + 1
        .Buttons.Add p, "NINKI", "�l�C��", tbrButtonGroup, 2
        .Buttons.item(p).value = IIf(Not mblnNumNinkiIsNum, tbrPressed, tbrUnpressed)
        p = p + 1
        .Buttons.Add p, "", "", tbrSeparator, 1
        p = p + 1
        .Buttons.Add p, "RACE", "�o�n�\", tbrDefault, 1
    End With
    With mToolBar.ToolBar(2)
        .Buttons(1).Caption = "" & _
        IIf(mViewerState.OddsHyosuIsOdds, "�I�b�Y", "�[��") & _
        "�擾"
    End With
    
    Exit Property
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Property


'
'   �@�\: Viewer��Ԓ�
'
'   ���l: �Ȃ�
'
Public Property Get ViewerState() As clsVSOdds
    Set ViewerState = mViewerState
End Property


'
'   �@�\: Viewer��Ԏ󂯎��
'
'   ���l: �Ȃ�
'
Public Property Set ViewerState(RHS As clsVSOdds)
    Set mViewerState = RHS
End Property


'
'   �@�\: �f�[�^�������u���E�U�ɓ`����
'
'   ���l:�@Viewer�K�{�v���p�e�B
'
Public Property Get NoData() As Boolean
    NoData = mblnNoData
End Property

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �����֐�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   �@�\: �N���b�N�C�x���g
'
'   ���l: �����N���ʂ֐؂�ւ���
'
Private Sub flexTab_Click(Index As Integer)
On Error GoTo ErrorHandler
    Dim msrow As Long             '' �}�E�X���E
    Dim mscol As Long             '' �}�E�X�J����
    Dim item As clsGridItem     '' �O���b�h�A�C�e��
    
    ' �}�E�X�ʒu�̃O���b�h���W���擾
    With flexTab(Index).Grid
        msrow = .MouseRow
        mscol = .MouseCol
    End With
    
    '�Z���͈͊O����
    If mscol < 0 Or msrow < 0 Then
        Exit Sub
    End If
    
    ' �O���b�h�A�C�e�����Z��������o��
    Call SetItem(item, flexTab(Index), msrow, mscol)
    
    
    ' �A�C�e���������N�������Ă���ꍇ
    If item.HasAKey Then
        ' ��ʐ؂�ւ��C�x���g���M
        RaiseEvent ChangeTo(item.Link, item.Key)
    End If
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �E�N���b�N�ŁA�R���e�L�X�g���j���[���o��
'
'   ���l: �Ȃ�
'
Private Sub flexTab_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrorHandler
    Dim msrow As Long
    Dim mscol As Long
    Dim item As clsGridItem
    
    ' �}�E�X�̎����O���b�h���W���擾
    msrow = flexTab(Index).Grid.MouseRow
    mscol = flexTab(Index).Grid.MouseCol
    
    '�Z���͈͊O����
    If mscol < 0 Or msrow < 0 Then
        Exit Sub
    End If
    
    ' �O���b�h�A�C�e�����Z��������o��
    Call SetItem(item, flexTab(Index), msrow, mscol)
    
    ' �f�[�^�������N�L�[�������Ă���
    If item.HasAKey Then
        ' ���A�E�N���b�N�ł���
        If Button = vbRightButton Then
            RaiseEvent LinkContextMenu(item.Link, item.Key)
        End If
    End If
    Exit Sub
    
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   �@�\: �}�E�X�̉��������N�\�ȃO���b�h�Ȃ�Δ�������ׂ̃C�x���g
'
'   ���l: �W���I�Ȕ����́AclsGridData.MouseMoveDriven�v���V�[�W���ɔC����
'
Private Sub flexTab_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrorHandler
    
    If Index = 0 Then
        Call flexTab(0).MouseMoveDriven
    End If
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �^�u�N���b�N�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub mstTab_Click(PreviousTab As Integer)
On Error GoTo ErrorHandler
    Dim i As Long
        
    '��ԕ��A�̈וϐ��Ɋi�[
    With mViewerState
        .LastTabNumber = mstTab.Tab
        .NumNinkiIsNum = mblnNumNinkiIsNum
        .OddsHyosuIsOdds = mblnOddsHyosuIsOdds
    End With
    
    '�^�u���ݒ�
    If mstTab.Tab = 4 Then Call Ninki  '3�A��
    If mstTab.Tab = 5 Then Call NinkiST '3�A�P

    For i = 0 To mstTab.Tabs - 1
        paneTab(i).Visible = (mstTab.Tab = i)
    Next i
    
    '���\���Ԃ��X�V
    Call Happyo
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   �@�\: Viewer�c�[���o�[�N���b�N�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub mToolBar_ButtonClick(Index As Integer, ByVal Button As MSComctlLib.Button)
On Error GoTo ErrorHandler
    tmrTBS.tag = Button.Key
    tmrTBS.Enabled = True
        
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   �@�\: 3�A���@�ԍ����@���{�^���N���b�N�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub optAxisNum_Click(Index As Integer)
On Error GoTo ErrorHandler
    Dim i       As Long
    Dim gridNum As Long
    
    '��ԕۑ�
    mViewerState.AxisNum = Index
    
    ' �R�A���ԍ����O���b�h�̂ݍĎ擾
    If mData.GridisExist(8) Then  '3�A���@�I�b�Y
        Call mData.FetchSanrenOddsNum(Index + 1)
    End If
    If mData.GridisExist(9) Then  '3�A���@�[��
        Call mData.FetchSanrenHyoNum(Index + 1)
    End If
    
    '���݂̃O���b�h��K�؂ȃ^�u�ɃC���T�[�g
    gridNum = 12 * IIf(mblnNumNinkiIsNum, 0, 1) + 4 * 2 + IIf(mblnOddsHyosuIsOdds, 0, 1) '�l�Cor�n�ԁ{3�A���^�u�{�I�b�Yor�[��
    If mData.GridisExist(gridNum) Then
        Call flexTab(Int(gridNum / 2) Mod 6).InsertGrid(mData.GridDatas(gridNum))
    End If
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: 3�A�P�@�ԍ��� ���{�^���N���b�N�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub optAxisNumST_Click(Index As Integer)
On Error GoTo ErrorHandler
    Dim i       As Long
    Dim gridNum As Long
    
    '��ԕۑ�
    mViewerState.AxisNumST = Index
    
    ' �R�A�P�ԍ����O���b�h�̂ݍĎ擾
    If mData.GridisExist(10) Then  '3�A�P�@�I�b�Y
        Call mData.FetchSanrentanOddsNum(Index + 1)
    End If
    If mData.GridisExist(11) Then  '3�A�P�@�[��
        Call mData.FetchSanrentanHyoNum(Index + 1)
    End If
    
    '���݂̃O���b�h��K�؂ȃ^�u�ɃC���T�[�g
    gridNum = 12 * IIf(mblnNumNinkiIsNum, 0, 1) + 5 * 2 + IIf(mblnOddsHyosuIsOdds, 0, 1) '�l�Cor�n�ԁ{3�A�P�^�u�{�I�b�Yor�[��
    If mData.GridisExist(gridNum) Then
        Call flexTab(Int(gridNum / 2) Mod 6).InsertGrid(mData.GridDatas(gridNum))
    End If
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: 3�A���@�l�C���@���{�^���N���b�N�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub optAxisNinki_Click(Index As Integer)
On Error GoTo ErrorHandler
    Dim lngCP As Long
       
    '��ԕۑ�
    mViewerState.AxisNinki = Index
    
    If mData.GridisExist(18) Or mData.GridisExist(19) Then
        With flexTab(4).Grid
            lngCP = Index * 30
            Do While (lngCP <> 0) And ((200 * Index + 1 < val(.TextMatrix(0, lngCP))) Or (.TextMatrix(0, lngCP) = ""))
                lngCP = lngCP - 3
            Loop
            .LeftCol = lngCP
        End With
    End If
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   �@�\: 3�A�P�@�l�C�� �R���{�{�b�N�X�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub cmbAxisNinki_click()
On Error GoTo ErrorHandler
    Dim i As Long
    Dim Index As Long
    Dim lngCP As Long

    With cmbAxisNinki
        '�R���{�{�b�N�X�̌��ݒl��Index�Ɏ擾
        If .ListIndex < 0 Then
            Index = 0
        Else
            Index = .ListIndex
        End If
    End With
    
    '�R���{�̏�Ԃ�ۑ�
    mViewerState.AxisNinkiST = Index
    
    If mData.GridisExist(22) Or mData.GridisExist(23) Then
        cmbAxisNinki.Enabled = True
        With flexTab(5).Grid
            lngCP = Index * 30
            Do While (lngCP <> 0) And ((200 * Index + 1 < val(.TextMatrix(0, lngCP))) Or (.TextMatrix(0, lngCP) = ""))  '�w��I�v�V�����{�^���ƃZ���ʒu�Ƃ��K�؂łȂ��Ȃ�
                lngCP = lngCP - 3
            Loop
            .LeftCol = lngCP
        End With
    End If
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   �@�\: ������
'
'   ���l: �Ȃ�
'
Private Sub tmrTBS_Timer()
On Error GoTo ErrorHandler
    tmrTBS.Enabled = False
    
    Select Case tmrTBS.tag
    Case "RACE"
        RaiseEvent ChangeTo("RA", mKey.str)
    Case "ODDS", "HYO", "NUM", "NINKI"
        
        mblnNumNinkiIsNum = (mToolBar.ToolBar(1).Buttons("NUM").value = tbrPressed)
        
        ViewerState.OddsHyosuIsOdds = mblnOddsHyosuIsOdds
        ViewerState.NumNinkiIsNum = mblnNumNinkiIsNum
        
        ' �K�؂ȃO���b�h�f�[�^���O���b�h�ɑ}��
        Call InsertGrid
        ' �R�A�����I���{�^���̐؂�ւ�
        Call SwitchOptAxis
        ' �R�A�P���I���{�^���̐؂�ւ�
        Call SwitchOptAxisST
        ' ���\�^�C���\���ؑ�
        Call Happyo
        
    Case "RTOPEN"
        Call gApp.DBPrompt(ukpOD, mKey.Year & mKey.MonthDay & mKey.JyoCD & mKey.RaceNum)
    End Select
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   �@�\: �擾�J�n�^�C�}�[
'
'   ���l: �Ȃ�
'
Private Sub tmrTrigger_Timer()
On Error GoTo ErrorHandler
    tmrTrigger.Enabled = False
    If Not mData Is Nothing Then
        Call mData.AsyncFetch
    End If
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: ���[�U�R���g���[��������
'
'   ���l: �Ȃ�
'
Private Sub UserControl_Initialize()
On Error GoTo ErrorHandler
    gApp.InitLog Me
    
    Dim i As Long
    Set mVB = New clsViewerBase
    Set mKey = New clsKeyRA
    Set mData = New clsDataOD
    Set mViewerState = New clsVSOdds
    
    ' �A�C�R���C���[�W���[�h
    With ilsTbrSmall
        .ImageHeight = 16
        .ImageWidth = 16
        .ListImages.Add 1, , LoadResPicture(200, vbResIcon)
        .ListImages.Add 2, , LoadResPicture(106, vbResIcon)
    End With
    
    ' ���ʂf�t�h�ݒ�
    Call mVB.InitGUI(UserControl.width, UserControl.Height, fraTop, lblMakeDate, fraHeader, mstTab)
    
    ' Font Asign
    Call mVB.FraTopFontType1(lblInfo(0).Font)
    
    ' FlexGrid���ʐݒ�
    For i = flexTab.LBound To flexTab.UBound
        Call mVB.FlexGridCommonSetting(flexTab(i).Grid)
        flexTab(i).Grid.GridLines = flexGridNone
    Next i
    
    ' Color Asign
    UserControl.BackColor = gApp.ColBG
    lblInfo(0).BackColor = gApp.ColDarkBG
    lblInfo(0).ForeColor = Contrast(gApp.ColDarkBG)
    ' skip lblInfo(1)
    For i = 2 To lblInfo.UBound
        lblInfo(i).BackColor = gApp.ColBG
        lblInfo(i).ForeColor = Contrast(gApp.ColBG)
    Next i
        
    '3�A�I�v�V�����{�^���ʒu�ݒ�
    fraOptAxis(0).Top = 0
    fraOptAxis(1).Top = 0
    fraOptAxis(2).Top = 0
    fraOptAxis(3).Top = 0
    
    mblnNumNinkiIsNum = True
    mblnOddsHyosuIsOdds = True
    
    Call mstTab_Click(0)
    
    ' ���ׂẴy�C�����A�f�[�^�擾���ɐݒ肷��B
    For i = paneTab.LBound To paneTab.UBound
        paneTab(i).Mode = 0
    Next i
    
    ' ���ׂă^�u�𖳌���Ԃɂ���
    For i = 0 To mstTab.Tabs - 1
        mstTab.TabEnabled(i) = False
    Next i
        
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: ���[�U�R���g���[���̃��T�C�Y�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub UserControl_Resize()
On Error GoTo ErrorHandler
    Dim i As Long
    
    Call mVB.ResizeGUI(UserControl.width, UserControl.Height, fraTop, lblMakeDate, fraHeader, mstTab)
    
    For i = 0 To 5
        With paneTab(i)
            .Left = 60
            .width = Bigger(1, mstTab.width - .Left * 2)
            .Height = Bigger(1, (mstTab.Height - .Top) - (.Top - mstTab.TabHeight))
        End With ' paneTab(mstTab.Index)
        With flexTab(i)
            .width = paneTab(i).width - .Left
            .Height = paneTab(i).Height - .Top
        End With
    Next i
    
    '�[���\���ʒu
    Const HyoPosition = 1800
    For i = 0 To 7
        lblHyo(i).Top = 30
    Next i
    For i = 0 To 2
        lblHyo(i).Left = paneTab(0).width - HyoPosition * (3 - i)
    Next i
    For i = 3 To 7
        lblHyo(i).Left = paneTab(i - 2).width - HyoPosition
    Next i
    
    '���\�^�C���\���ʒu
    Call Happyo
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   �@�\: ���[�U�R���g���[���I���C�x���g
'
'   ���l: �Ȃ�
'
Private Sub UserControl_Terminate()
On Error GoTo ErrorHandler
    gApp.TermLog Me
    
    Set mKey = Nothing
    Set mData = Nothing
    Set mVB = Nothing
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �f�[�^���Ȃ�
'
'   ���l: �Ȃ�
'
Private Sub mData_NoData()
On Error GoTo ErrorHandler
    gApp.Log "d:�Y�����R�[�h�͂���܂���BOD�̑��݂���key���w�肵�Ă��������B" & vbCrLf _
            & "�Ăяo�������`�F�b�N���܂��傤�����J����"
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �����n���[�X�f�[�^���Ȃ�
'
'   ���l: �Ȃ�
'
Private Sub mData_NoUMARACE()
On Error GoTo ErrorHandler
    Dim i As Long
    
    'UmaRace(���邢�͔n��)���Ȃ��̂ŃI�b�Y�֘AmToolBar���g�p�s�ł�
    For i = 1 To 5
        mToolBar.ToolBar(1).Buttons(i).Enabled = False
    Next i
    'UmaRace(���邢�͔n��)���Ȃ��̂őS�Ă̔ԍ����^�u���g�p�s�ł�
    For i = 0 To 5
        paneTab(i).Mode = 1
        mstTab.TabEnabled(i) = False
    Next i
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �eGridDatas(Index)�擾�����ʒm�C�x���g
'
'   ���l: �Ȃ�
'
Private Sub mData_Fetched(Index As Long)
On Error GoTo ErrorHandler
    Dim i As Long
    Dim tabNum As Long
    
    If mblnOddsHyosuIsOdds = (0 = Index Mod 2) And mblnNumNinkiIsNum = (Index < 12) Then
        tabNum = Int((Index Mod 12) / 2)
        Call flexTab(tabNum).InsertGrid(mData.GridDatas(Index))
        If Index = 8 Or Index = 9 Then  '�I�b�Y3�A��
            Call optAxisNum_Click(mViewerState.AxisNum)
        ElseIf Index = 10 Or Index = 11 Then  '�I�b�Y3�A�P
            Call optAxisNumST_Click(mViewerState.AxisNumST)
        End If
    
        Call ResizeGrid(tabNum)
        
        '�}�[�W�A��
        With flexTab(tabNum).Grid
            .FixedCols = 0
            .FixedRows = 0
            .MergeRow(0) = True
            .MergeCells = flexMergeFree
            If tabNum <> 0 Then
                .MergeRow(0) = False
            End If
        End With
        
        '�^�u����Ԃ̕ύX
        paneTab(tabNum).Mode = 2
        mstTab.TabEnabled(tabNum) = True
        If mblnOddsHyosuIsOdds And (Index = 10 Or Index = 22) Then mstTab.TabVisible(5) = True '�I�b�Y3�A�P�f�[�^�������3�A�P�^�u��\��
        If Not mblnOddsHyosuIsOdds And (Index = 11 Or Index = 23) Then mstTab.TabVisible(5) = True '�[��3�A�P�f�[�^�������3�A�P�^�u��\��
        For i = 0 To 7
            lblHyo(i).Visible = (Left$(lblHyo(i), 3) <> "lbl" And Trim$(lblHyo(i)) <> "")
        Next i
    End If
    
    '3�A�P�l�C����
    If Index = 22 Or Index = 23 Then
        cmbAxisNinki.ListIndex = mViewerState.AxisNinkiST
    End If

    '���{�^���̕\����؂�ւ���
    Select Case Index
    Case 8, 9, 20, 21: '3�A��
        Call SwitchOptAxis
    Case 10, 11, 22, 23: '3�A�P
        Call SwitchOptAxisST
    End Select
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �f�[�^�擾���S�ďI�������Ƃ��̃C�x���g
'
'   ���l: �Ȃ�
'
Private Sub mData_FetchedAll()
On Error GoTo ErrorHandler
    Dim i As Long
    
    '�@�[����ݒ�
    For i = 0 To 7
        lblHyo(i) = mData.Hyo(i)
        lblHyo(i).Visible = (Left$(lblHyo(i), 3) <> "lbl" And Trim$(lblHyo(i)) <> "")
    Next i
    
    '�I�v�V�����{�^���̉s�ݒ�
    For i = 0 To 17
        optAxisNum(i).Enabled = (i < mData.TOSU)
        optAxisNumST(i).Enabled = (i < mData.TOSU)
    Next i
    
    mblnFetchedAll = True
    Call InsertGrid
    
    ' ���\�^�C����ݒ�
    Call Happyo
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   �@�\: �f�[�^�擾�A���f
'
'   ���l: �Ȃ�
'
Private Sub Update()
On Error GoTo ErrorHandler
    Dim i As Long
    
    ' �f�[�^���擾
    mData.OddsHyosuIsOdds = mViewerState.OddsHyosuIsOdds
    If Not mData.Fetch(mKey) Then
        mblnNoData = True
        Exit Sub
    End If
    
    ' �E�C���h�E�^�C�g���̕ύX
    Me.Title = IIf(mViewerState.OddsHyosuIsOdds, "�I�b�Y ", "�[�� ") & mData.Title
    
    ' ���x�����擾
    For i = 0 To 7
        lblInfo(i) = ReplaceAmpersand(mData.Labels(i))
    Next i
    lblMakeDate = mData.Labels(8)
    lblInfo(8).Caption = mData.Labels(9) ' ���R�[�h
    
    ' ���x���𐮗񂳂���
    lblInfo(2).Left = 0
    lblInfo(2).Top = 30
    lblInfo(3).Left = lblInfo(2).Left + lblInfo(2).width
    lblInfo(3).Top = lblInfo(2).Top
    lblInfo(4).Left = lblInfo(3).Left
    lblInfo(4).Top = lblInfo(3).Top + lblInfo(3).Height
    lblInfo(5).Left = lblInfo(4).Left
    lblInfo(5).Top = lblInfo(4).Top + lblInfo(4).Height
    lblInfo(6).Left = Bigger(lblInfo(4).Left + lblInfo(4).width, lblInfo(5).Left + lblInfo(5).width)
    lblInfo(6).Top = lblInfo(4).Top
    lblInfo(7).Left = lblInfo(6).Left
    lblInfo(7).Top = lblInfo(6).Top + lblInfo(6).Height
    lblInfo(8).Left = lblInfo(7).Left + lblInfo(7).width
    lblInfo(8).Top = lblInfo(7).Top
    lblInfo(8).ForeColor = vbRed
    
    '�I�v�V�����{�^���̉s�ݒ�
    For i = 0 To 17
        optAxisNum(i).Enabled = False
        optAxisNumST(i).Enabled = False '3�A�P
    Next i
    For i = 0 To 3
        optAxisNinki(i).Enabled = False
    Next i
    cmbAxisNinki.Enabled = False '3�A�P
    
    '3�A�P�^�u�̕\����\������
    If mblnOddsHyosuIsOdds Then  '�I�b�Y
        mstTab.TabVisible(5) = mData.O6Exist
    Else '�[��
        mstTab.TabVisible(5) = mData.H6Exist
    End If
    
    '�O���Ԃ̕���
    If Not mViewerState.IsNoTouch Then
        With mViewerState
            mblnNumNinkiIsNum = IIf(.NumNinkiIsNum, tbrPressed, tbrUnpressed)
            mblnOddsHyosuIsOdds = IIf(.OddsHyosuIsOdds, tbrPressed, tbrUnpressed)
            Call InsertGrid
            optAxisNum(.AxisNum).value = True
            optAxisNinki(.AxisNinki).value = True
            optAxisNumST(.AxisNumST).value = True  '3�A�P
            cmbAxisNinki.ListIndex = .AxisNinkiST '3�A�P
            mstTab.Tab = .LastTabNumber
        End With
    Else
        Call InsertGrid
        mstTab.Tab = 0
    End If
    
    tmrTrigger.Enabled = True
        
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   �@�\: �O���b�h�ɁA�K�؂ȃf�[�^���Z�b�g
'
'   ���l: �t���O�ɉ����āAflexTab��I�����āA6�̃O���b�h�ɑ}������
'
Private Sub InsertGrid()
On Error GoTo ErrorHandler
    Dim i       As Long
    Dim j       As Long
    Dim num     As Long
    Dim Index   As Long
    
    gApp.Log "InsertGrid"
    
    For i = 0 To 5
        flexTab(i).Grid.Clear
        num = (12 * IIf(mblnNumNinkiIsNum, 0, 1)) _
                + (i * 2 + IIf(mblnOddsHyosuIsOdds, 0, 1))
        If mData.GridisExist(num) = True Then
            Call flexTab(i).InsertGrid(mData.GridDatas(num))
            
            Call ResizeGrid(i)
            

            '�}�[�W�A��
            With flexTab(i).Grid
                .FixedCols = 0
                .FixedRows = 0
                .MergeRow(0) = True
                .MergeCells = flexMergeFree
                If i <> 0 Then
                    .MergeRow(0) = False
                End If
            End With
            
            paneTab(i).Mode = 2
            mstTab.TabEnabled(i) = True
        Else
            paneTab(i).Mode = IIf(mblnFetchedAll, 1, 0)
            mstTab.TabEnabled(i) = False
        End If
    Next i
    
    For j = 0 To 7
        lblHyo(j).Visible = (Left$(lblHyo(j), 3) <> "lbl" And Trim$(lblHyo(j)) <> "")
    Next j
    
    If 4 = mstTab.Tab Then Call Ninki
    If 5 = mstTab.Tab Then Call NinkiST

    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   �@�\: �I�b�Y�[���̃��[�h�ɂ��킹�āA�R�A���I�v�V�����{�^���̕\����؂�ւ���
'
'   ���l: �Ȃ�
'
Private Sub SwitchOptAxis()
    If mblnOddsHyosuIsOdds Then  '�I�b�Y�{�^����������Ă����
        '���I�v�V�����{�^���\���ݒ�
        fraOptAxis(0).Visible = mblnNumNinkiIsNum And mData.GridisExist(8)
        '�l�C���I�v�V�����{�^���\���ݒ�
        fraOptAxis(1).Visible = Not mblnNumNinkiIsNum And mData.GridisExist(20)
    Else                         '�i�[���{�^����������Ă���΁j
        '���I�v�V�����{�^���\���ݒ�
        fraOptAxis(0).Visible = mblnNumNinkiIsNum And mData.GridisExist(9)
        '�l�C���I�v�V�����{�^���\���ݒ�
        fraOptAxis(1).Visible = Not mblnNumNinkiIsNum And mData.GridisExist(21)
    End If
End Sub


'
'   �@�\: �I�b�Y�[���̃��[�h�ɂ��킹�āA�R�A�P�I�v�V�����{�^���̕\����؂�ւ���
'
'   ���l: �Ȃ�
'
Private Sub SwitchOptAxisST()
    If mblnOddsHyosuIsOdds Then  '�I�b�Y�{�^����������Ă����
        '���I�v�V�����{�^���\���ݒ�
        fraOptAxis(2).Visible = mblnNumNinkiIsNum And mData.GridisExist(10)
        '�l�C���I�v�V�����{�^���\���ݒ�
        fraOptAxis(3).Visible = Not mblnNumNinkiIsNum And mData.GridisExist(22)
    Else                         '�i�[���{�^����������Ă���΁j
        '���I�v�V�����{�^���\���ݒ�
        fraOptAxis(2).Visible = mblnNumNinkiIsNum And mData.GridisExist(11)
        '�l�C���I�v�V�����{�^���\���ݒ�
        fraOptAxis(3).Visible = Not mblnNumNinkiIsNum And mData.GridisExist(23)
    End If
End Sub


'
'   �@�\: ���\�^�C���������ݒ�
'
'   ���l: �Ȃ�
'
Private Sub Happyo()
    With lblInfo(1)
        If mblnFetchedAll = True Then
            If mblnOddsHyosuIsOdds Then
                .Caption = mData.Happyo(mstTab.Tab)
                .Left = fraTop.width - .width - 100
                .Visible = (Trim$(lblInfo(1)) <> "")
            Else
                .Caption = IIf(5 = mstTab.Tab, mData.Happyo(7), mData.Happyo(6)) '�R�[�h��
                .Left = fraTop.width - .width - 100
                .Visible = (Trim$(lblInfo(1)) <> "")
            End If
        Else
            .Visible = False
        End If
    End With
End Sub


'
'   �@�\: 3�A���^�u���\���ݒ�
'
'   ���l: �Ȃ�
'
Private Sub Ninki()
    Dim i As Long
    Dim Index As Long
    Dim lngCP As Long

    'optAxisNinki�ŉ����ꂽ�{�^���̎w��ʒu�ɉ�ʈړ�
    If (False = mblnNumNinkiIsNum) And (mData.GridisExist(20) Or mData.GridisExist(21)) Then  '�l�C���@���@�l�C���R�A���f�[�^�擾�ς݂Ȃ�
        For i = 0 To 3
            '�e�l�C���I�v�V�����{�^���̉s���A�l�C���������画��
            optAxisNinki(i).Enabled = (i * 30 < flexTab(4).Grid.Cols)
            '�I�v�V�����{�^���̌��ݒl��Index�Ɏ擾
            If (optAxisNinki(i).value = True) Then Index = i
        Next i
        With flexTab(4).Grid
            lngCP = Index * 30
            Do While (lngCP <> 0) And (200 * Index + 1 < val(.TextMatrix(0, lngCP))) Or (.TextMatrix(0, lngCP) = "")  '�w��I�v�V�����{�^���ƃZ���ʒu�Ƃ��K�؂łȂ��Ȃ�
                lngCP = lngCP - 3
            Loop
        End With
    End If
    
    '�I�v�V�����{�^���\���ؑ�
    Call SwitchOptAxis
    
End Sub


'
'   �@�\: 3�A�P�^�u���\���ݒ�
'
'   ���l: �Ȃ�
'
Private Sub NinkiST()
    Dim i As Long
    Dim Index As Long
    Dim lngCP As Long

    '�I�v�V�����{�^���\���ؑ�
    Call SwitchOptAxisST
    
    'optAxisNinkiST�ŉ����ꂽ�{�^���̎w��ʒu�ɉ�ʈړ�
    If (False = mblnNumNinkiIsNum) And (mData.GridisExist(22) Or mData.GridisExist(23)) Then  '�l�C���@���@�l�C���R�A�P�f�[�^�擾�ς݂Ȃ�
        With cmbAxisNinki
            '�R���{�{�b�N�X�̌��ݒl��Index�Ɏ擾
            If .ListIndex < 0 Then
                Index = 0
                .ListIndex = 0
            Else
                Index = .ListIndex
            End If
            .Clear
            .AddItem "   1" & "�`"
            For i = 1 To 24
                '�e�l�C���R���{�{�b�N�X�A�C�e�����A�l�C�������i�O���b�h�J�����j���画��
                If 30 * i < flexTab(5).Grid.Cols Then
                    .AddItem Format(200 * i + 1, "@@@@") & "�`"
                End If
            Next i
            .ListIndex = Index
        End With
        cmbAxisNinki.Enabled = True
        With flexTab(5).Grid
            lngCP = Index * 30
            Do While (lngCP <> 0) And ((200 * Index + 1 < val(.TextMatrix(0, lngCP))) Or (.TextMatrix(0, lngCP) = ""))  '�w��I�v�V�����{�^���ƃZ���ʒu�Ƃ��K�؂łȂ��Ȃ�
                lngCP = lngCP - 3
            Loop
        End With
    End If
    
    '�I�v�V�����{�^���\���ؑ�
    Call SwitchOptAxisST
    
End Sub


'
'   �@�\: �O���b�h���̒���
'
'   ���l: �Ȃ�
'
Private Sub ResizeGrid(ByVal i As Integer)
On Error GoTo ErrHandler
    Dim intLoop As Integer
    Dim size1 As Integer, size3 As Integer
    If Not mblnNumNinkiIsNum And mblnOddsHyosuIsOdds Then
        Select Case i
            Case 0:
                flexTab(0).Grid.ColWidth(0) = 360
                flexTab(0).Grid.ColWidth(1) = 360
                flexTab(0).Grid.ColWidth(2) = 630
                flexTab(0).Grid.ColWidth(3) = 210
                flexTab(0).Grid.ColWidth(4) = 360
                flexTab(0).Grid.ColWidth(5) = 360
                flexTab(0).Grid.ColWidth(6) = 1260
                flexTab(0).Grid.ColWidth(7) = 210
                flexTab(0).Grid.ColWidth(8) = 360
                flexTab(0).Grid.ColWidth(9) = 630
                flexTab(0).Grid.ColWidth(10) = 540
                flexTab(0).Grid.ColWidth(11) = 360
                flexTab(0).Grid.ColWidth(12) = 630
                flexTab(0).Grid.ColWidth(13) = 630
                flexTab(0).Grid.ColWidth(14) = 360
                flexTab(0).Grid.ColWidth(15) = 630
                flexTab(0).Grid.ColWidth(16) = 720
            Case 1:
                flexTab(1).Grid.ColWidth(0) = 360
                flexTab(1).Grid.ColWidth(1) = 630
                flexTab(1).Grid.ColWidth(2) = 540
                For intLoop = 3 To 9 Step 3
                    flexTab(i).Grid.ColWidth(intLoop) = 360
                    flexTab(i).Grid.ColWidth(intLoop + 1) = 630
                    flexTab(i).Grid.ColWidth(intLoop + 2) = 630
                Next
                If flexTab(i).Grid.Cols > 11 Then
                    For intLoop = 12 To flexTab(i).Grid.Cols - 1 Step 3
                        flexTab(i).Grid.ColWidth(intLoop) = 450
                        flexTab(i).Grid.ColWidth(intLoop + 1) = 630
                        flexTab(i).Grid.ColWidth(intLoop + 2) = 720
                    Next
                End If
            Case 2:
                For intLoop = 0 To 9 Step 3
                    flexTab(i).Grid.ColWidth(intLoop) = 360
                    flexTab(i).Grid.ColWidth(intLoop + 1) = 630
                    flexTab(i).Grid.ColWidth(intLoop + 2) = 1440
                Next
                If flexTab(i).Grid.Cols > 11 Then
                    For intLoop = 12 To flexTab(i).Grid.Cols - 1 Step 3
                        flexTab(i).Grid.ColWidth(intLoop) = 450
                        flexTab(i).Grid.ColWidth(intLoop + 1) = 630
                        flexTab(i).Grid.ColWidth(intLoop + 2) = 1440
                    Next
                End If
            Case 3:
                flexTab(1).Grid.ColWidth(0) = 360
                flexTab(1).Grid.ColWidth(1) = 630
                flexTab(1).Grid.ColWidth(2) = 540
                For intLoop = 3 To 9 Step 3
                    flexTab(i).Grid.ColWidth(intLoop) = 360
                    flexTab(i).Grid.ColWidth(intLoop + 1) = 630
                    flexTab(i).Grid.ColWidth(intLoop + 2) = 630
                Next
                For intLoop = 12 To flexTab(i).Grid.Cols - 1 Step 3
                    flexTab(i).Grid.ColWidth(intLoop) = 450
                    flexTab(i).Grid.ColWidth(intLoop + 1) = 630
                    flexTab(i).Grid.ColWidth(intLoop + 2) = 720
                Next
                flexTab(i).Grid.ColWidth(flexTab(i).Grid.Cols - 1) = 810
            Case 4:
                flexTab(1).Grid.ColWidth(0) = 360
                flexTab(1).Grid.ColWidth(1) = 900
                flexTab(1).Grid.ColWidth(2) = 540
                For intLoop = 3 To 9 Step 3
                    flexTab(i).Grid.ColWidth(intLoop) = 360
                    flexTab(i).Grid.ColWidth(intLoop + 1) = 900
                    flexTab(i).Grid.ColWidth(intLoop + 2) = 630
                Next
                For intLoop = 12 To 24 Step 3
                    flexTab(i).Grid.ColWidth(intLoop) = 450
                    flexTab(i).Grid.ColWidth(intLoop + 1) = 900
                    flexTab(i).Grid.ColWidth(intLoop + 2) = 630
                Next
                For intLoop = 27 To flexTab(i).Grid.Cols - 1 Step 3
                    flexTab(i).Grid.ColWidth(intLoop) = 450
                    flexTab(i).Grid.ColWidth(intLoop + 1) = 900
                    flexTab(i).Grid.ColWidth(intLoop + 2) = 720
                Next
            Case 5:
                flexTab(1).Grid.ColWidth(0) = 360
                flexTab(1).Grid.ColWidth(1) = 900
                flexTab(1).Grid.ColWidth(2) = 540
                size1 = 360
                size3 = 630
                For intLoop = 3 To flexTab(i).Grid.Cols - 1 Step 3
                    If intLoop = 12 Then
                        size1 = 450
                    ElseIf intLoop = 27 Then
                        size3 = 720
                    ElseIf intLoop = 54 Then
                        size3 = 810
                    ElseIf intLoop = 108 Then
                        size3 = 900
                    End If
                    flexTab(i).Grid.ColWidth(intLoop) = size1
                    flexTab(i).Grid.ColWidth(intLoop + 1) = 900
                    flexTab(i).Grid.ColWidth(intLoop + 2) = size3
                Next
        End Select
    ElseIf mblnNumNinkiIsNum And mblnOddsHyosuIsOdds Then
        Select Case i
            Case 0:
                flexTab(0).Grid.ColWidth(0) = 360
                flexTab(0).Grid.ColWidth(1) = 360
                flexTab(0).Grid.ColWidth(2) = 1800
                flexTab(0).Grid.ColWidth(3) = 450
                flexTab(0).Grid.ColWidth(4) = 900
                flexTab(0).Grid.ColWidth(5) = 1080
                flexTab(0).Grid.ColWidth(6) = 1260
                flexTab(0).Grid.ColWidth(7) = 210
                For intLoop = 8 To flexTab(i).Grid.Cols - 2 Step 2
                    flexTab(0).Grid.ColWidth(intLoop) = 450
                    flexTab(0).Grid.ColWidth(intLoop + 1) = 720
                Next
            Case 1, 3:
                For intLoop = 0 To flexTab(i).Grid.Cols - 2 Step 2
                    flexTab(i).Grid.ColWidth(intLoop) = 360
                    flexTab(i).Grid.ColWidth(intLoop + 1) = 810
                Next
            Case 2:
                For intLoop = 0 To flexTab(i).Grid.Cols - 2 Step 2
                    flexTab(i).Grid.ColWidth(intLoop) = 360
                    flexTab(i).Grid.ColWidth(intLoop + 1) = 1440
                Next
            Case 4, 5:
                For intLoop = 0 To flexTab(i).Grid.Cols - 2 Step 2
                    flexTab(i).Grid.ColWidth(intLoop) = 360
                    flexTab(i).Grid.ColWidth(intLoop + 1) = 810
                Next
        End Select
    ElseIf mblnNumNinkiIsNum And Not mblnOddsHyosuIsOdds Then
        Select Case i
            Case 0:
                flexTab(0).Grid.ColWidth(0) = 360
                flexTab(0).Grid.ColWidth(1) = 360
                flexTab(0).Grid.ColWidth(2) = 1800
                flexTab(0).Grid.ColWidth(3) = 450
                flexTab(0).Grid.ColWidth(4) = 900
                flexTab(0).Grid.ColWidth(5) = 900
                flexTab(0).Grid.ColWidth(6) = 900
                flexTab(0).Grid.ColWidth(7) = 210
                For intLoop = 8 To flexTab(i).Grid.Cols - 2 Step 2
                    flexTab(0).Grid.ColWidth(intLoop) = 450
                    flexTab(0).Grid.ColWidth(intLoop + 1) = 630
                Next
            Case 1, 2, 3, 4, 5:
                For intLoop = 0 To flexTab(i).Grid.Cols - 2 Step 2
                    flexTab(i).Grid.ColWidth(intLoop) = 360
                    flexTab(i).Grid.ColWidth(intLoop + 1) = 630
                Next
        End Select
    Else
        Select Case i
            Case 0:
                flexTab(0).Grid.ColWidth(0) = 360
                flexTab(0).Grid.ColWidth(1) = 360
                flexTab(0).Grid.ColWidth(2) = 630
                flexTab(0).Grid.ColWidth(3) = 210
                flexTab(0).Grid.ColWidth(4) = 360
                flexTab(0).Grid.ColWidth(5) = 360
                flexTab(0).Grid.ColWidth(6) = 630
                flexTab(0).Grid.ColWidth(7) = 210
                flexTab(0).Grid.ColWidth(8) = 360
                flexTab(0).Grid.ColWidth(9) = 630
                flexTab(0).Grid.ColWidth(10) = 360
                flexTab(0).Grid.ColWidth(11) = 360
                flexTab(0).Grid.ColWidth(12) = 630
                flexTab(0).Grid.ColWidth(13) = 360
                flexTab(0).Grid.ColWidth(14) = 360
                flexTab(0).Grid.ColWidth(15) = 210
                flexTab(0).Grid.ColWidth(16) = 630
            Case 1, 2, 3:
                size1 = 360
                size3 = 630
                For intLoop = 0 To flexTab(i).Grid.Cols - 1 Step 3
                    If intLoop = 12 Then
                        size1 = 450
                    ElseIf intLoop = 27 Then
                        size3 = 720
                    ElseIf intLoop = 54 Then
                        size3 = 810
                    ElseIf intLoop = 108 Then
                        size3 = 900
                    End If
                    flexTab(i).Grid.ColWidth(intLoop) = size1
                    flexTab(i).Grid.ColWidth(intLoop + 1) = 630
                    flexTab(i).Grid.ColWidth(intLoop + 2) = size3
                Next
            Case 4, 5:
                size1 = 360
                size3 = 630
                For intLoop = 0 To flexTab(i).Grid.Cols - 1 Step 3
                    If intLoop = 12 Then
                        size1 = 450
                    ElseIf intLoop = 27 Then
                        size3 = 720
                    ElseIf intLoop = 54 Then
                        size3 = 810
                    ElseIf intLoop = 108 Then
                        size3 = 900
                    End If
                    flexTab(i).Grid.ColWidth(intLoop) = size1
                    flexTab(i).Grid.ColWidth(intLoop + 1) = 810
                    flexTab(i).Grid.ColWidth(intLoop + 2) = size3
                Next
        End Select
    End If
    
    Exit Sub
ErrHandler:
    gApp.ErrLog
    Resume Next
End Sub


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   �O���֐�
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   �@�\: �I������
'
'   ���l: �Ȃ�
'
Public Sub Free()
    On Error GoTo ErrHandler
    Dim i As Long
    
    If Not mData Is Nothing Then
        mData.CancelFetching
        For i = 0 To 23
            If Not mData.GridDatas(i) Is Nothing Then
                mData.GridDatas(i).Free
            Else
                Exit For
            End If
        Next i
        Set mData = Nothing
    End If
    
    Exit Sub
ErrHandler:
    gApp.Log Err.Description
End Sub


