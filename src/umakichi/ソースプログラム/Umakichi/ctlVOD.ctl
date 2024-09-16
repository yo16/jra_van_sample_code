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
      BorderStyle     =   0  'なし
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
            Name            =   "ＭＳ ゴシック"
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
            Name            =   "ＭＳ ゴシック"
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
            Name            =   "ＭＳ ゴシック"
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
            Name            =   "ＭＳ ゴシック"
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
            Name            =   "ＭＳ ゴシック"
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
            Name            =   "ＭＳ ゴシック"
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
            Name            =   "ＭＳ ゴシック"
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
      BorderStyle     =   0  'なし
      Caption         =   "Frame1"
      Height          =   465
      Left            =   120
      TabIndex        =   1
      Top             =   180
      Width           =   8955
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "データ： 1月 1日16時10分"
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
         Caption         =   "リソース不足です。不要な画面を閉じてください"
         BeginProperty Font 
            Name            =   "ＭＳ Ｐゴシック"
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
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "単勝･複勝･枠連"
      TabPicture(0)   =   "ctlVOD.ctx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "paneTab(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "　　馬連　　"
      TabPicture(1)   =   "ctlVOD.ctx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "paneTab(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "　　ワイド　　"
      TabPicture(2)   =   "ctlVOD.ctx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "paneTab(2)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "　　馬単　　"
      TabPicture(3)   =   "ctlVOD.ctx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "paneTab(3)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "　　３連複　　"
      TabPicture(4)   =   "ctlVOD.ctx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "paneTab(4)"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "　　３連単　　"
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
            BorderStyle     =   0  'なし
            Caption         =   "Frame2"
            Height          =   255
            Index           =   1
            Left            =   0
            TabIndex        =   16
            Top             =   150
            Width           =   4905
            Begin VB.OptionButton optAxisNinki 
               Caption         =   "601～"
               Height          =   255
               Index           =   3
               Left            =   3600
               Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
               TabIndex        =   20
               Top             =   0
               Width           =   1200
            End
            Begin VB.OptionButton optAxisNinki 
               Caption         =   "401～"
               Height          =   255
               Index           =   2
               Left            =   2400
               Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
               TabIndex        =   19
               Top             =   0
               Width           =   1200
            End
            Begin VB.OptionButton optAxisNinki 
               Caption         =   "201～"
               Height          =   255
               Index           =   1
               Left            =   1200
               Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
               TabIndex        =   18
               Top             =   0
               Width           =   1200
            End
            Begin VB.OptionButton optAxisNinki 
               Caption         =   "1～"
               Height          =   255
               Index           =   0
               Left            =   0
               Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
               TabIndex        =   17
               Top             =   0
               Value           =   -1  'True
               Width           =   1200
            End
         End
         Begin VB.Frame fraOptAxis 
            BorderStyle     =   0  'なし
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
               Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
               TabIndex        =   39
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNum 
               Caption         =   "17"
               Height          =   255
               Index           =   16
               Left            =   5760
               Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
               TabIndex        =   38
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNum 
               Caption         =   "16"
               Height          =   255
               Index           =   15
               Left            =   5430
               Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
               TabIndex        =   37
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNum 
               Caption         =   "15"
               Height          =   255
               Index           =   14
               Left            =   5100
               Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
               TabIndex        =   36
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNum 
               Caption         =   "14"
               Height          =   255
               Index           =   13
               Left            =   4770
               Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
               TabIndex        =   35
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNum 
               Caption         =   "13"
               Height          =   255
               Index           =   12
               Left            =   4440
               Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
               TabIndex        =   34
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNum 
               Caption         =   "12"
               Height          =   255
               Index           =   11
               Left            =   4110
               Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
               TabIndex        =   33
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNum 
               Caption         =   "11"
               Height          =   255
               Index           =   10
               Left            =   3780
               Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
               TabIndex        =   32
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNum 
               Caption         =   "10"
               Height          =   255
               Index           =   9
               Left            =   3450
               Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
               TabIndex        =   31
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNum 
               Caption         =   "9"
               Height          =   255
               Index           =   8
               Left            =   3120
               Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
               TabIndex        =   30
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNum 
               Caption         =   "8"
               Height          =   255
               Index           =   7
               Left            =   2790
               Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
               TabIndex        =   29
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNum 
               Caption         =   "7"
               Height          =   255
               Index           =   6
               Left            =   2460
               Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
               TabIndex        =   28
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNum 
               Caption         =   "6"
               Height          =   255
               Index           =   5
               Left            =   2130
               Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
               TabIndex        =   27
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNum 
               Caption         =   "5"
               Height          =   255
               Index           =   4
               Left            =   1800
               Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
               TabIndex        =   26
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNum 
               Caption         =   "4"
               Height          =   255
               Index           =   3
               Left            =   1470
               Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
               TabIndex        =   25
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNum 
               Caption         =   "3"
               Height          =   255
               Index           =   2
               Left            =   1140
               Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
               TabIndex        =   24
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNum 
               Caption         =   "2"
               Height          =   255
               Index           =   1
               Left            =   810
               Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
               TabIndex        =   23
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNum 
               Caption         =   "1"
               Height          =   255
               Index           =   0
               Left            =   480
               Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
               TabIndex        =   22
               Top             =   0
               Value           =   -1  'True
               Width           =   315
            End
            Begin VB.Label lblFixJiku 
               AutoSize        =   -1  'True
               Caption         =   "軸馬"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
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
            BorderStyle     =   0  'なし
            Caption         =   "Frame2"
            Height          =   300
            Index           =   3
            Left            =   120
            TabIndex        =   75
            Top             =   120
            Width           =   1785
            Begin VB.ComboBox cmbAxisNinki 
               BeginProperty Font 
                  Name            =   "ＭＳ ゴシック"
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
               Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
               TabIndex        =   76
               Top             =   0
               Width           =   1215
            End
         End
         Begin VB.Frame fraOptAxis 
            BorderStyle     =   0  'なし
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
               Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
               TabIndex        =   73
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNumST 
               Caption         =   "17"
               Height          =   255
               Index           =   16
               Left            =   5760
               Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
               TabIndex        =   72
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNumST 
               Caption         =   "16"
               Height          =   255
               Index           =   15
               Left            =   5430
               Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
               TabIndex        =   71
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNumST 
               Caption         =   "15"
               Height          =   255
               Index           =   14
               Left            =   5100
               Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
               TabIndex        =   70
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNumST 
               Caption         =   "14"
               Height          =   255
               Index           =   13
               Left            =   4770
               Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
               TabIndex        =   69
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNumST 
               Caption         =   "13"
               Height          =   255
               Index           =   12
               Left            =   4440
               Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
               TabIndex        =   68
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNumST 
               Caption         =   "12"
               Height          =   255
               Index           =   11
               Left            =   4110
               Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
               TabIndex        =   67
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNumST 
               Caption         =   "11"
               Height          =   255
               Index           =   10
               Left            =   3780
               Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
               TabIndex        =   66
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNumST 
               Caption         =   "10"
               Height          =   255
               Index           =   9
               Left            =   3450
               Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
               TabIndex        =   65
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNumST 
               Caption         =   "9"
               Height          =   255
               Index           =   8
               Left            =   3120
               Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
               TabIndex        =   64
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNumST 
               Caption         =   "8"
               Height          =   255
               Index           =   7
               Left            =   2790
               Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
               TabIndex        =   63
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNumST 
               Caption         =   "7"
               Height          =   255
               Index           =   6
               Left            =   2460
               Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
               TabIndex        =   62
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNumST 
               Caption         =   "6"
               Height          =   255
               Index           =   5
               Left            =   2130
               Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
               TabIndex        =   61
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNumST 
               Caption         =   "5"
               Height          =   255
               Index           =   4
               Left            =   1800
               Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
               TabIndex        =   60
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNumST 
               Caption         =   "4"
               Height          =   255
               Index           =   3
               Left            =   1470
               Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
               TabIndex        =   59
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNumST 
               Caption         =   "3"
               Height          =   255
               Index           =   2
               Left            =   1140
               Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
               TabIndex        =   58
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNumST 
               Caption         =   "2"
               Height          =   255
               Index           =   1
               Left            =   810
               Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
               TabIndex        =   57
               Top             =   0
               Width           =   315
            End
            Begin VB.OptionButton optAxisNumST 
               Caption         =   "1"
               Height          =   255
               Index           =   0
               Left            =   480
               Style           =   1  'ｸﾞﾗﾌｨｯｸｽ
               TabIndex        =   56
               Top             =   0
               Value           =   -1  'True
               Width           =   315
            End
            Begin VB.Label lblFixJikuST 
               AutoSize        =   -1  'True
               Caption         =   "軸馬"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
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
'   オッズ･票数 表示コントロール
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数(イベント)
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Public Event ChangeTo(strViewerName As String, strKey As String)        '' Vierer変更イベント
Public Event WindowTitle(strKey As String)                              '' ウインドウタイトル変更イベント
Public Event LinkContextMenu(strViewerName As String, strKey As String) '' 右クリックメニュー表示イベント
Public Event Reload()                                                   '' 再読み込み

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mVB As clsViewerBase            '' Viewer Base
Private mViewerState As clsVSOdds       '' Viewer State
Private WithEvents mToolBar As ctlToolBars  '' Viewer ToolBar
Attribute mToolBar.VB_VarHelpID = -1

Private WithEvents mData As clsDataOD   '' データ取得オブジェクト
Attribute mData.VB_VarHelpID = -1
Private mstrTitle As String             '' ウインドウタイトル
Private mKey      As clsKeyRA           '' キー
Private mblnNoData As Boolean           '' データ無しフラグ

Private mblnOddsHyosuIsOdds As Boolean  '' ラジオボタン・オッズ票数が、オッズであれば真
Private mblnNumNinkiIsNum As Boolean    '' ラジオボタン・番号人気が、番号であれば真
Private mblnAction As Boolean           '' ラジオボタン・番号人気が、動作中であれば真
Private mblnFetchedAll As Boolean       '' 取得終了していれば真

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部定数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

' 画面最小幅値
Const MINIMUMWIDTH  As Long = 10000
Const MINIMUMHEIGHT As Long = 5000


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   プロパティ
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8


'
'   機能: キー設定プロパティ
'
'   備考: Viewer必須プロパティ
'
Public Property Let Key(strKey As String)
    gApp.Log "OD: " & strKey
    mKey.str = strKey
    mViewerState.OddsHyosuIsOdds = (Right$(strKey, 1) = "0")
        
    Call Update
End Property


'
'   機能: タイトル取得プロパティ
'
'   備考: Viewer必須プロパティ、　Browser が参照
'
Public Property Get Title() As String
    Title = mstrTitle
End Property


'
'   機能: タイトル設定プロパティ
'
'   備考: ブラウザに変更通知のイベント発生
'
Public Property Let Title(strTitle As String)
    mstrTitle = strTitle
    RaiseEvent WindowTitle(mstrTitle)
End Property


'
'   機能: ツールバーを設定する
'
'   備考: ブラウザからツールバーを受け取り、ツールバーをセットアップする
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
        .Buttons.Add p, "NUM", "馬番順", tbrButtonGroup, 2
        .Buttons.item(p).value = IIf(mblnNumNinkiIsNum, tbrPressed, tbrUnpressed)
        p = p + 1
        .Buttons.Add p, "NINKI", "人気順", tbrButtonGroup, 2
        .Buttons.item(p).value = IIf(Not mblnNumNinkiIsNum, tbrPressed, tbrUnpressed)
        p = p + 1
        .Buttons.Add p, "", "", tbrSeparator, 1
        p = p + 1
        .Buttons.Add p, "RACE", "出馬表", tbrDefault, 1
    End With
    With mToolBar.ToolBar(2)
        .Buttons(1).Caption = "" & _
        IIf(mViewerState.OddsHyosuIsOdds, "オッズ", "票数") & _
        "取得"
    End With
    
    Exit Property
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Property


'
'   機能: Viewer状態提供
'
'   備考: なし
'
Public Property Get ViewerState() As clsVSOdds
    Set ViewerState = mViewerState
End Property


'
'   機能: Viewer状態受け取り
'
'   備考: なし
'
Public Property Set ViewerState(RHS As clsVSOdds)
    Set mViewerState = RHS
End Property


'
'   機能: データ無しをブラウザに伝える
'
'   備考:　Viewer必須プロパティ
'
Public Property Get NoData() As Boolean
    NoData = mblnNoData
End Property

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: クリックイベント
'
'   備考: リンク先画面へ切り替える
'
Private Sub flexTab_Click(Index As Integer)
On Error GoTo ErrorHandler
    Dim msrow As Long             '' マウスロウ
    Dim mscol As Long             '' マウスカラム
    Dim item As clsGridItem     '' グリッドアイテム
    
    ' マウス位置のグリッド座標を取得
    With flexTab(Index).Grid
        msrow = .MouseRow
        mscol = .MouseCol
    End With
    
    'セル範囲外検査
    If mscol < 0 Or msrow < 0 Then
        Exit Sub
    End If
    
    ' グリッドアイテムをセルから取り出す
    Call SetItem(item, flexTab(Index), msrow, mscol)
    
    
    ' アイテムがリンクを持っている場合
    If item.HasAKey Then
        ' 画面切り替えイベント送信
        RaiseEvent ChangeTo(item.Link, item.Key)
    End If
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 右クリックで、コンテキストメニューを出す
'
'   備考: なし
'
Private Sub flexTab_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrorHandler
    Dim msrow As Long
    Dim mscol As Long
    Dim item As clsGridItem
    
    ' マウスの示すグリッド座標を取得
    msrow = flexTab(Index).Grid.MouseRow
    mscol = flexTab(Index).Grid.MouseCol
    
    'セル範囲外検査
    If mscol < 0 Or msrow < 0 Then
        Exit Sub
    End If
    
    ' グリッドアイテムをセルから取り出す
    Call SetItem(item, flexTab(Index), msrow, mscol)
    
    ' データがリンクキーを持っている
    If item.HasAKey Then
        ' かつ、右クリックである
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
'   機能: マウスの下がリンク可能なグリッドならば反応する為のイベント
'
'   備考: 標準的な反応は、clsGridData.MouseMoveDrivenプロシージャに任せる
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
'   機能: タブクリックイベント
'
'   備考: なし
'
Private Sub mstTab_Click(PreviousTab As Integer)
On Error GoTo ErrorHandler
    Dim i As Long
        
    '状態復帰の為変数に格納
    With mViewerState
        .LastTabNumber = mstTab.Tab
        .NumNinkiIsNum = mblnNumNinkiIsNum
        .OddsHyosuIsOdds = mblnOddsHyosuIsOdds
    End With
    
    'タブ内設定
    If mstTab.Tab = 4 Then Call Ninki  '3連複
    If mstTab.Tab = 5 Then Call NinkiST '3連単

    For i = 0 To mstTab.Tabs - 1
        paneTab(i).Visible = (mstTab.Tab = i)
    Next i
    
    '発表時間を更新
    Call Happyo
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: Viewerツールバークリックイベント
'
'   備考: なし
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
'   機能: 3連複　番号順　軸ボタンクリックイベント
'
'   備考: なし
'
Private Sub optAxisNum_Click(Index As Integer)
On Error GoTo ErrorHandler
    Dim i       As Long
    Dim gridNum As Long
    
    '状態保存
    mViewerState.AxisNum = Index
    
    ' ３連複番号順グリッドのみ再取得
    If mData.GridisExist(8) Then  '3連複　オッズ
        Call mData.FetchSanrenOddsNum(Index + 1)
    End If
    If mData.GridisExist(9) Then  '3連複　票数
        Call mData.FetchSanrenHyoNum(Index + 1)
    End If
    
    '現在のグリッドを適切なタブにインサート
    gridNum = 12 * IIf(mblnNumNinkiIsNum, 0, 1) + 4 * 2 + IIf(mblnOddsHyosuIsOdds, 0, 1) '人気or馬番＋3連複タブ＋オッズor票数
    If mData.GridisExist(gridNum) Then
        Call flexTab(Int(gridNum / 2) Mod 6).InsertGrid(mData.GridDatas(gridNum))
    End If
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 3連単　番号順 軸ボタンクリックイベント
'
'   備考: なし
'
Private Sub optAxisNumST_Click(Index As Integer)
On Error GoTo ErrorHandler
    Dim i       As Long
    Dim gridNum As Long
    
    '状態保存
    mViewerState.AxisNumST = Index
    
    ' ３連単番号順グリッドのみ再取得
    If mData.GridisExist(10) Then  '3連単　オッズ
        Call mData.FetchSanrentanOddsNum(Index + 1)
    End If
    If mData.GridisExist(11) Then  '3連単　票数
        Call mData.FetchSanrentanHyoNum(Index + 1)
    End If
    
    '現在のグリッドを適切なタブにインサート
    gridNum = 12 * IIf(mblnNumNinkiIsNum, 0, 1) + 5 * 2 + IIf(mblnOddsHyosuIsOdds, 0, 1) '人気or馬番＋3連単タブ＋オッズor票数
    If mData.GridisExist(gridNum) Then
        Call flexTab(Int(gridNum / 2) Mod 6).InsertGrid(mData.GridDatas(gridNum))
    End If
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 3連複　人気順　軸ボタンクリックイベント
'
'   備考: なし
'
Private Sub optAxisNinki_Click(Index As Integer)
On Error GoTo ErrorHandler
    Dim lngCP As Long
       
    '状態保存
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
'   機能: 3連単　人気順 コンボボックスイベント
'
'   備考: なし
'
Private Sub cmbAxisNinki_click()
On Error GoTo ErrorHandler
    Dim i As Long
    Dim Index As Long
    Dim lngCP As Long

    With cmbAxisNinki
        'コンボボックスの現在値をIndexに取得
        If .ListIndex < 0 Then
            Index = 0
        Else
            Index = .ListIndex
        End If
    End With
    
    'コンボの状態を保存
    mViewerState.AxisNinkiST = Index
    
    If mData.GridisExist(22) Or mData.GridisExist(23) Then
        cmbAxisNinki.Enabled = True
        With flexTab(5).Grid
            lngCP = Index * 30
            Do While (lngCP <> 0) And ((200 * Index + 1 < val(.TextMatrix(0, lngCP))) Or (.TextMatrix(0, lngCP) = ""))  '指定オプションボタンとセル位置とが適切でないなら
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
'   機能: 初期化
'
'   備考: なし
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
        
        ' 適切なグリッドデータをグリッドに挿入
        Call InsertGrid
        ' ３連複軸選択ボタンの切り替え
        Call SwitchOptAxis
        ' ３連単軸選択ボタンの切り替え
        Call SwitchOptAxisST
        ' 発表タイム表示切替
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
'   機能: 取得開始タイマー
'
'   備考: なし
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
'   機能: ユーザコントロール初期化
'
'   備考: なし
'
Private Sub UserControl_Initialize()
On Error GoTo ErrorHandler
    gApp.InitLog Me
    
    Dim i As Long
    Set mVB = New clsViewerBase
    Set mKey = New clsKeyRA
    Set mData = New clsDataOD
    Set mViewerState = New clsVSOdds
    
    ' アイコンイメージロード
    With ilsTbrSmall
        .ImageHeight = 16
        .ImageWidth = 16
        .ListImages.Add 1, , LoadResPicture(200, vbResIcon)
        .ListImages.Add 2, , LoadResPicture(106, vbResIcon)
    End With
    
    ' 共通ＧＵＩ設定
    Call mVB.InitGUI(UserControl.width, UserControl.Height, fraTop, lblMakeDate, fraHeader, mstTab)
    
    ' Font Asign
    Call mVB.FraTopFontType1(lblInfo(0).Font)
    
    ' FlexGrid共通設定
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
        
    '3連オプションボタン位置設定
    fraOptAxis(0).Top = 0
    fraOptAxis(1).Top = 0
    fraOptAxis(2).Top = 0
    fraOptAxis(3).Top = 0
    
    mblnNumNinkiIsNum = True
    mblnOddsHyosuIsOdds = True
    
    Call mstTab_Click(0)
    
    ' すべてのペインを、データ取得中に設定する。
    For i = paneTab.LBound To paneTab.UBound
        paneTab(i).Mode = 0
    Next i
    
    ' すべてタブを無効状態にする
    For i = 0 To mstTab.Tabs - 1
        mstTab.TabEnabled(i) = False
    Next i
        
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: ユーザコントロールのリサイズイベント
'
'   備考: なし
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
    
    '票数表示位置
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
    
    '発表タイム表示位置
    Call Happyo
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: ユーザコントロール終了イベント
'
'   備考: なし
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
'   機能: データがない
'
'   備考: なし
'
Private Sub mData_NoData()
On Error GoTo ErrorHandler
    gApp.Log "d:該当レコードはありません。ODの存在するkeyを指定してください。" & vbCrLf _
            & "呼び出し元をチェックしましょう＞＞開発者"
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 競走馬レースデータがない
'
'   備考: なし
'
Private Sub mData_NoUMARACE()
On Error GoTo ErrorHandler
    Dim i As Long
    
    'UmaRace(あるいは馬番)がないのでオッズ関連mToolBarが使用不可です
    For i = 1 To 5
        mToolBar.ToolBar(1).Buttons(i).Enabled = False
    Next i
    'UmaRace(あるいは馬番)がないので全ての番号順タブが使用不可です
    For i = 0 To 5
        paneTab(i).Mode = 1
        mstTab.TabEnabled(i) = False
    Next i
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 各GridDatas(Index)取得完了通知イベント
'
'   備考: なし
'
Private Sub mData_Fetched(Index As Long)
On Error GoTo ErrorHandler
    Dim i As Long
    Dim tabNum As Long
    
    If mblnOddsHyosuIsOdds = (0 = Index Mod 2) And mblnNumNinkiIsNum = (Index < 12) Then
        tabNum = Int((Index Mod 12) / 2)
        Call flexTab(tabNum).InsertGrid(mData.GridDatas(Index))
        If Index = 8 Or Index = 9 Then  'オッズ3連複
            Call optAxisNum_Click(mViewerState.AxisNum)
        ElseIf Index = 10 Or Index = 11 Then  'オッズ3連単
            Call optAxisNumST_Click(mViewerState.AxisNumST)
        End If
    
        Call ResizeGrid(tabNum)
        
        'マージ、寄せ
        With flexTab(tabNum).Grid
            .FixedCols = 0
            .FixedRows = 0
            .MergeRow(0) = True
            .MergeCells = flexMergeFree
            If tabNum <> 0 Then
                .MergeRow(0) = False
            End If
        End With
        
        'タブ内状態の変更
        paneTab(tabNum).Mode = 2
        mstTab.TabEnabled(tabNum) = True
        If mblnOddsHyosuIsOdds And (Index = 10 Or Index = 22) Then mstTab.TabVisible(5) = True 'オッズ3連単データがあれば3連単タブを表示
        If Not mblnOddsHyosuIsOdds And (Index = 11 Or Index = 23) Then mstTab.TabVisible(5) = True '票数3連単データがあれば3連単タブを表示
        For i = 0 To 7
            lblHyo(i).Visible = (Left$(lblHyo(i), 3) <> "lbl" And Trim$(lblHyo(i)) <> "")
        Next i
    End If
    
    '3連単人気順の
    If Index = 22 Or Index = 23 Then
        cmbAxisNinki.ListIndex = mViewerState.AxisNinkiST
    End If

    '軸ボタンの表示を切り替える
    Select Case Index
    Case 8, 9, 20, 21: '3連複
        Call SwitchOptAxis
    Case 10, 11, 22, 23: '3連単
        Call SwitchOptAxisST
    End Select
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: データ取得が全て終了したときのイベント
'
'   備考: なし
'
Private Sub mData_FetchedAll()
On Error GoTo ErrorHandler
    Dim i As Long
    
    '　票数を設定
    For i = 0 To 7
        lblHyo(i) = mData.Hyo(i)
        lblHyo(i).Visible = (Left$(lblHyo(i), 3) <> "lbl" And Trim$(lblHyo(i)) <> "")
    Next i
    
    'オプションボタンの可不可設定
    For i = 0 To 17
        optAxisNum(i).Enabled = (i < mData.TOSU)
        optAxisNumST(i).Enabled = (i < mData.TOSU)
    Next i
    
    mblnFetchedAll = True
    Call InsertGrid
    
    ' 発表タイムを設定
    Call Happyo
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: データ取得、反映
'
'   備考: なし
'
Private Sub Update()
On Error GoTo ErrorHandler
    Dim i As Long
    
    ' データを取得
    mData.OddsHyosuIsOdds = mViewerState.OddsHyosuIsOdds
    If Not mData.Fetch(mKey) Then
        mblnNoData = True
        Exit Sub
    End If
    
    ' ウインドウタイトルの変更
    Me.Title = IIf(mViewerState.OddsHyosuIsOdds, "オッズ ", "票数 ") & mData.Title
    
    ' ラベルを取得
    For i = 0 To 7
        lblInfo(i) = ReplaceAmpersand(mData.Labels(i))
    Next i
    lblMakeDate = mData.Labels(8)
    lblInfo(8).Caption = mData.Labels(9) ' レコード
    
    ' ラベルを整列させる
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
    
    'オプションボタンの可不可設定
    For i = 0 To 17
        optAxisNum(i).Enabled = False
        optAxisNumST(i).Enabled = False '3連単
    Next i
    For i = 0 To 3
        optAxisNinki(i).Enabled = False
    Next i
    cmbAxisNinki.Enabled = False '3連単
    
    '3連単タブの表示非表示判定
    If mblnOddsHyosuIsOdds Then  'オッズ
        mstTab.TabVisible(5) = mData.O6Exist
    Else '票数
        mstTab.TabVisible(5) = mData.H6Exist
    End If
    
    '前回状態の復元
    If Not mViewerState.IsNoTouch Then
        With mViewerState
            mblnNumNinkiIsNum = IIf(.NumNinkiIsNum, tbrPressed, tbrUnpressed)
            mblnOddsHyosuIsOdds = IIf(.OddsHyosuIsOdds, tbrPressed, tbrUnpressed)
            Call InsertGrid
            optAxisNum(.AxisNum).value = True
            optAxisNinki(.AxisNinki).value = True
            optAxisNumST(.AxisNumST).value = True  '3連単
            cmbAxisNinki.ListIndex = .AxisNinkiST '3連単
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
'   機能: グリッドに、適切なデータをセット
'
'   備考: フラグに応じて、flexTabを選択して、6つのグリッドに挿入する
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
            

            'マージ、寄せ
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
'   機能: オッズ票数のモードにあわせて、３連複オプションボタンの表示を切り替える
'
'   備考: なし
'
Private Sub SwitchOptAxis()
    If mblnOddsHyosuIsOdds Then  'オッズボタンが押されていれば
        '軸オプションボタン表示設定
        fraOptAxis(0).Visible = mblnNumNinkiIsNum And mData.GridisExist(8)
        '人気順オプションボタン表示設定
        fraOptAxis(1).Visible = Not mblnNumNinkiIsNum And mData.GridisExist(20)
    Else                         '（票数ボタンが押されていれば）
        '軸オプションボタン表示設定
        fraOptAxis(0).Visible = mblnNumNinkiIsNum And mData.GridisExist(9)
        '人気順オプションボタン表示設定
        fraOptAxis(1).Visible = Not mblnNumNinkiIsNum And mData.GridisExist(21)
    End If
End Sub


'
'   機能: オッズ票数のモードにあわせて、３連単オプションボタンの表示を切り替える
'
'   備考: なし
'
Private Sub SwitchOptAxisST()
    If mblnOddsHyosuIsOdds Then  'オッズボタンが押されていれば
        '軸オプションボタン表示設定
        fraOptAxis(2).Visible = mblnNumNinkiIsNum And mData.GridisExist(10)
        '人気順オプションボタン表示設定
        fraOptAxis(3).Visible = Not mblnNumNinkiIsNum And mData.GridisExist(22)
    Else                         '（票数ボタンが押されていれば）
        '軸オプションボタン表示設定
        fraOptAxis(2).Visible = mblnNumNinkiIsNum And mData.GridisExist(11)
        '人気順オプションボタン表示設定
        fraOptAxis(3).Visible = Not mblnNumNinkiIsNum And mData.GridisExist(23)
    End If
End Sub


'
'   機能: 発表タイム文字列を設定
'
'   備考: なし
'
Private Sub Happyo()
    With lblInfo(1)
        If mblnFetchedAll = True Then
            If mblnOddsHyosuIsOdds Then
                .Caption = mData.Happyo(mstTab.Tab)
                .Left = fraTop.width - .width - 100
                .Visible = (Trim$(lblInfo(1)) <> "")
            Else
                .Caption = IIf(5 = mstTab.Tab, mData.Happyo(7), mData.Happyo(6)) 'コード中
                .Left = fraTop.width - .width - 100
                .Visible = (Trim$(lblInfo(1)) <> "")
            End If
        Else
            .Visible = False
        End If
    End With
End Sub


'
'   機能: 3連複タブ内表示設定
'
'   備考: なし
'
Private Sub Ninki()
    Dim i As Long
    Dim Index As Long
    Dim lngCP As Long

    'optAxisNinkiで押されたボタンの指定位置に画面移動
    If (False = mblnNumNinkiIsNum) And (mData.GridisExist(20) Or mData.GridisExist(21)) Then  '人気順　且　人気順３連複データ取得済みなら
        For i = 0 To 3
            '各人気順オプションボタンの可不可を、人気順総数から判定
            optAxisNinki(i).Enabled = (i * 30 < flexTab(4).Grid.Cols)
            'オプションボタンの現在値をIndexに取得
            If (optAxisNinki(i).value = True) Then Index = i
        Next i
        With flexTab(4).Grid
            lngCP = Index * 30
            Do While (lngCP <> 0) And (200 * Index + 1 < val(.TextMatrix(0, lngCP))) Or (.TextMatrix(0, lngCP) = "")  '指定オプションボタンとセル位置とが適切でないなら
                lngCP = lngCP - 3
            Loop
        End With
    End If
    
    'オプションボタン表示切替
    Call SwitchOptAxis
    
End Sub


'
'   機能: 3連単タブ内表示設定
'
'   備考: なし
'
Private Sub NinkiST()
    Dim i As Long
    Dim Index As Long
    Dim lngCP As Long

    'オプションボタン表示切替
    Call SwitchOptAxisST
    
    'optAxisNinkiSTで押されたボタンの指定位置に画面移動
    If (False = mblnNumNinkiIsNum) And (mData.GridisExist(22) Or mData.GridisExist(23)) Then  '人気順　且　人気順３連単データ取得済みなら
        With cmbAxisNinki
            'コンボボックスの現在値をIndexに取得
            If .ListIndex < 0 Then
                Index = 0
                .ListIndex = 0
            Else
                Index = .ListIndex
            End If
            .Clear
            .AddItem "   1" & "～"
            For i = 1 To 24
                '各人気順コンボボックスアイテムを、人気順総数（グリッドカラム）から判定
                If 30 * i < flexTab(5).Grid.Cols Then
                    .AddItem Format(200 * i + 1, "@@@@") & "～"
                End If
            Next i
            .ListIndex = Index
        End With
        cmbAxisNinki.Enabled = True
        With flexTab(5).Grid
            lngCP = Index * 30
            Do While (lngCP <> 0) And ((200 * Index + 1 < val(.TextMatrix(0, lngCP))) Or (.TextMatrix(0, lngCP) = ""))  '指定オプションボタンとセル位置とが適切でないなら
                lngCP = lngCP - 3
            Loop
        End With
    End If
    
    'オプションボタン表示切替
    Call SwitchOptAxisST
    
End Sub


'
'   機能: グリッド幅の調整
'
'   備考: なし
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
'   外部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: 終了処理
'
'   備考: なし
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


