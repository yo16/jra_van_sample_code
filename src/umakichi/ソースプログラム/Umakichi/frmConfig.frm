VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmConfig 
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "馬吉設定"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7800
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '画面の中央
   Begin TabDlg.SSTab mstTab 
      Height          =   7695
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   7785
      _ExtentX        =   13732
      _ExtentY        =   13573
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "データベース"
      TabPicture(0)   =   "frmConfig.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdCancel(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdOK(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "fraDB"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraJVL"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "色の設定"
      TabPicture(1)   =   "frmConfig.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "dlgColorChoice"
      Tab(1).Control(1)=   "fraTheme"
      Tab(1).Control(2)=   "fraPreview"
      Tab(1).Control(3)=   "cmdOK(1)"
      Tab(1).Control(4)=   "cmdCancel(1)"
      Tab(1).Control(5)=   "cmdReverse"
      Tab(1).ControlCount=   6
      Begin VB.Frame fraJVL 
         Caption         =   "JV-Link設定"
         Height          =   705
         Left            =   90
         TabIndex        =   11
         Top             =   330
         Width           =   7575
         Begin VB.PictureBox picXPTheme 
            BorderStyle     =   0  'なし
            Height          =   375
            Index           =   1
            Left            =   120
            ScaleHeight     =   375
            ScaleWidth      =   6975
            TabIndex        =   25
            Top             =   240
            Width           =   6975
            Begin VB.CommandButton cmdJVSetUIProperties 
               Caption         =   "JVLink設定ダイアログを呼び出す"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   9
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   0
               TabIndex        =   26
               Top             =   0
               Width           =   3225
            End
            Begin VB.Label lblFix 
               Caption         =   "サービスキーの設定や、動作オプションの変更をする場合は、このボタンを押してください。"
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   8.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Index           =   3
               Left            =   3420
               TabIndex        =   27
               Top             =   0
               Width           =   3300
            End
         End
      End
      Begin VB.Frame fraDB 
         Caption         =   "データベース設定"
         Height          =   6075
         Left            =   90
         TabIndex        =   10
         Top             =   1110
         Width           =   7545
         Begin VB.Frame frmJVLMode 
            Caption         =   "JV-Link 取得モード"
            Height          =   4125
            Left            =   120
            TabIndex        =   31
            Top             =   1200
            Width           =   7365
            Begin VB.PictureBox picXPThema 
               BorderStyle     =   0  'なし
               Height          =   3795
               Index           =   0
               Left            =   60
               ScaleHeight     =   3795
               ScaleWidth      =   7245
               TabIndex        =   32
               Top             =   240
               Width           =   7245
               Begin VB.OptionButton optJVMode 
                  Caption         =   "通常モード"
                  BeginProperty Font 
                     Name            =   "ＭＳ Ｐゴシック"
                     Size            =   9
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Index           =   0
                  Left            =   30
                  TabIndex        =   39
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   1425
               End
               Begin VB.CheckBox chkBLOD 
                  Caption         =   "産駒・繁殖馬を含める"
                  BeginProperty Font 
                     Name            =   "ＭＳ Ｐゴシック"
                     Size            =   9
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   2100
                  TabIndex        =   36
                  Top             =   210
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   2355
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
                  Height          =   585
                  Index           =   2
                  Left            =   60
                  Locked          =   -1  'True
                  MultiLine       =   -1  'True
                  TabIndex        =   34
                  Text            =   "frmConfig.frx":0038
                  Top             =   3360
                  Width           =   5565
               End
               Begin VB.OptionButton optJVMode 
                  Caption         =   "今週モード"
                  BeginProperty Font 
                     Name            =   "ＭＳ Ｐゴシック"
                     Size            =   9
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Index           =   1
                  Left            =   0
                  TabIndex        =   33
                  Top             =   2280
                  Width           =   1305
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
                  Height          =   1095
                  Index           =   1
                  Left            =   390
                  Locked          =   -1  'True
                  MultiLine       =   -1  'True
                  TabIndex        =   35
                  Text            =   "frmConfig.frx":00C3
                  Top             =   2550
                  Width           =   6225
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
                  Height          =   2025
                  Index           =   0
                  Left            =   450
                  Locked          =   -1  'True
                  MultiLine       =   -1  'True
                  TabIndex        =   38
                  Text            =   "frmConfig.frx":01F2
                  Top             =   510
                  Width           =   6765
               End
               Begin VB.CheckBox chkSLOP 
                  Caption         =   "坂路調教を含める"
                  BeginProperty Font 
                     Name            =   "ＭＳ Ｐゴシック"
                     Size            =   9
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   270
                  TabIndex        =   37
                  Top             =   210
                  Value           =   1  'ﾁｪｯｸ
                  Width           =   2085
               End
               Begin VB.Label lblFix 
                  AutoSize        =   -1  'True
                  BackColor       =   &H000000FF&
                  Caption         =   "セットアップ完了まで変更できません"
                  BeginProperty Font 
                     Name            =   "ＭＳ Ｐゴシック"
                     Size            =   8.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H8000000E&
                  Height          =   165
                  Index           =   1
                  Left            =   4530
                  TabIndex        =   40
                  Top             =   240
                  Width           =   2520
               End
            End
         End
         Begin VB.PictureBox picXPTheme 
            BorderStyle     =   0  'なし
            Height          =   345
            Index           =   0
            Left            =   150
            ScaleHeight     =   345
            ScaleWidth      =   7095
            TabIndex        =   21
            Top             =   240
            Width           =   7095
            Begin VB.TextBox txtPath 
               Height          =   300
               Left            =   1920
               TabIndex        =   23
               Text            =   "C:\db\"
               Top             =   0
               Width           =   4485
            End
            Begin VB.CommandButton cmdDirRef 
               Caption         =   "参照"
               Height          =   300
               Left            =   6390
               TabIndex        =   22
               Top             =   0
               Width           =   645
            End
            Begin VB.Label lblFix 
               AutoSize        =   -1  'True
               Caption         =   "データベースフォルダ："
               BeginProperty Font 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   9
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   24
               Top             =   60
               Width           =   1905
            End
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            Caption         =   "セットアップ開始年度： xxxx年"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   2070
            TabIndex        =   42
            Top             =   840
            Width           =   2610
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            Caption         =   "最終データ取得日時"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   0
            Left            =   2070
            TabIndex        =   41
            Top             =   600
            Width           =   1620
         End
      End
      Begin VB.CommandButton cmdReverse 
         Caption         =   "色を逆にする"
         Height          =   375
         Left            =   -70950
         TabIndex        =   9
         Top             =   4800
         Width           =   2175
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "キャンセル"
         Height          =   375
         Index           =   1
         Left            =   -70950
         TabIndex        =   8
         Top             =   5430
         Width           =   1905
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   375
         Index           =   1
         Left            =   -72930
         TabIndex        =   7
         Top             =   5430
         Width           =   1905
      End
      Begin VB.Frame fraPreview 
         Caption         =   "プレビュー"
         Height          =   3255
         Left            =   -70950
         TabIndex        =   4
         Top             =   1440
         Width           =   2175
         Begin VB.Label lblFGDisp 
            BackColor       =   &H80000009&
            Caption         =   "　　　　　　　　　前景"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   690
            Left            =   360
            TabIndex        =   6
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Label lblBGDisp 
            BackColor       =   &H80000007&
            Caption         =   "　　　　　　　　　　　　背景"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   2655
            Left            =   120
            TabIndex        =   5
            Top             =   360
            Width           =   1935
         End
      End
      Begin VB.Frame fraTheme 
         Caption         =   "色のテーマ "
         Height          =   3735
         Left            =   -74430
         TabIndex        =   3
         Top             =   1440
         Width           =   3255
         Begin VB.PictureBox picXPThema 
            BorderStyle     =   0  'なし
            Height          =   3285
            Index           =   1
            Left            =   240
            ScaleHeight     =   3285
            ScaleWidth      =   2775
            TabIndex        =   12
            Top             =   360
            Width           =   2775
            Begin VB.OptionButton optTheme 
               Caption         =   "ハロウィン"
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   20
               Top             =   0
               Width           =   1575
            End
            Begin VB.OptionButton optTheme 
               Caption         =   "クールミント"
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   19
               Top             =   360
               Width           =   1575
            End
            Begin VB.OptionButton optTheme 
               Caption         =   "秋"
               Height          =   255
               Index           =   2
               Left            =   0
               TabIndex        =   18
               Top             =   720
               Width           =   1575
            End
            Begin VB.OptionButton optTheme 
               Caption         =   "血のように赤い"
               Height          =   255
               Index           =   3
               Left            =   0
               TabIndex        =   17
               Top             =   1080
               Width           =   1575
            End
            Begin VB.OptionButton optTheme 
               Caption         =   "プリティーピンク"
               Height          =   255
               Index           =   4
               Left            =   0
               TabIndex        =   16
               Top             =   1440
               Width           =   1725
            End
            Begin VB.OptionButton optTheme 
               Caption         =   "ディフォルト"
               Height          =   255
               Index           =   5
               Left            =   0
               TabIndex        =   15
               Top             =   1800
               Width           =   1575
            End
            Begin VB.OptionButton optTheme 
               Caption         =   "カスタマイズ"
               Height          =   255
               Index           =   6
               Left            =   0
               TabIndex        =   14
               Top             =   2160
               Width           =   2055
            End
            Begin VB.Frame fraAdvanced 
               Caption         =   "カラーカスタマイズ"
               Height          =   660
               Left            =   120
               TabIndex        =   13
               Top             =   2520
               Width           =   2475
               Begin VB.PictureBox picXPTheme 
                  BorderStyle     =   0  'なし
                  Height          =   345
                  Index           =   2
                  Left            =   150
                  ScaleHeight     =   345
                  ScaleWidth      =   2265
                  TabIndex        =   28
                  Top             =   240
                  Width           =   2265
                  Begin VB.CommandButton cmdColor 
                     Caption         =   "前景色"
                     Height          =   300
                     Index           =   1
                     Left            =   1200
                     TabIndex        =   30
                     Top             =   0
                     Width           =   975
                  End
                  Begin VB.CommandButton cmdColor 
                     Caption         =   "背景色"
                     Height          =   300
                     Index           =   0
                     Left            =   0
                     TabIndex        =   29
                     Top             =   0
                     Width           =   975
                  End
               End
            End
         End
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   375
         Index           =   0
         Left            =   1950
         TabIndex        =   2
         Top             =   7230
         Width           =   1905
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "キャンセル"
         Height          =   375
         Index           =   0
         Left            =   3930
         TabIndex        =   1
         Top             =   7230
         Width           =   1905
      End
      Begin MSComDlg.CommonDialog dlgColorChoice 
         Left            =   -74880
         Top             =   420
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'   環境設定画面
'
Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mblnNoEdit As Boolean       '' 未編集フラグ

Private mlngBGColor As Long         '' 背景色
Private mlngFGColor As Long         '' 前面色
Private mlngPrevBGColor As Long     '' 変更前の背景色
Private mlngPrevFGColor As Long     '' 変更前の前面色

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8


'
'   機能: データベースの不足数を返す
'
'   備考: なし
'
Public Function MissingDBNum() As Long
    Dim SpecIDs()   As String
    Dim dbpath      As String
    Dim fso         As FileSystemObject
    Dim count       As Long
    Dim i           As Long

    gApp.Log "FSO New"
    Set fso = New FileSystemObject
    gApp.Log "FSO New finish"
    dbpath = gApp.R_DBPath

    SpecIDs = Split("BANUSI BATAIJYU CHOKYO CHOKYO_SEISEKI HANRO HANSYOKU HARAI KISHU KISHU_CHANGE KISHU_SEISEKI MINING ODDS_SANREN0 ODDS_SANREN1 ODDS_SANREN2 ODDS_SANREN3 ODDS_SANREN4 ODDS_SANREN5 ODDS_SANREN6 ODDS_SANREN7 ODDS_SANREN8 ODDS_SANREN9 ODDS_TANPUKUWAKU ODDS_UMAREN ODDS_UMATAN0 ODDS_UMATAN1 ODDS_UMATAN2 ODDS_UMATAN3 ODDS_UMATAN4 ODDS_UMATAN5 ODDS_UMATAN6 ODDS_UMATAN7 ODDS_UMATAN8 ODDS_UMATAN9 ODDS_WIDE RACE RECORD SANKU SCHEDULE SEISAN TENKO_BABA TOKU TOKU_RACE TORIKESI_JYOGAI UMA UMA_RACE_A UMA_RACE_B")

    For i = 0 To UBound(SpecIDs)
        If Not fso.FileExists(dbpath & "\sub" & SpecIDs(i) & ".mdb") Then
            count = count + 1
        End If
    Next i

    If count = UBound(SpecIDs) + 1 Then
        count = -1
    End If

    MissingDBNum = count
End Function


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: BLODデータ同期処理をする
'
'   備考: なし
'
Private Sub SetAndSyncBLOD()
On Error GoTo Errorhandler
    Dim DBUpdate As frmDBUpdate
    Dim MsgResult As VbMsgBoxResult

    gApp.Log "Config BLOD option: " & chkBLOD.value
    If gApp.R_JVLGetBLOD = False And chkBLOD.value = 1 And gApp.R_JVDLastTime <> String$(14, "0") Then
        If gApp.R_JVDLastTimeBLOD = String$(14, "0") Then
            MsgResult = MsgBox("産駒繁殖馬データはまだセットアップされていません。" & vbCrLf & _
                        "産駒繁殖馬データのセットアップを開始しますか？", vbYesNo + vbQuestion, "馬吉：産駒繁殖馬データセットアップの確認")
        Else
            MsgResult = MsgBox("不足している分の産駒繁殖馬データを取得します。", vbOKCancel + vbInformation, "馬吉：産駒繁殖馬データ未取得分の取得開始の確認")
        End If
        
        If MsgResult = vbCancel Then
            chkBLOD.value = 0
        Else
            Set DBUpdate = New frmDBUpdate
            DBUpdate.GettingMode = 2 '0:Other 1:SLOP 2:BLOD 3:O6H6 else:err
            DBUpdate.Show vbModal, Me
            If DBUpdate.Finish Then
                MsgResult = MsgBox("次回の更新からは、産駒繁殖馬データも含まれるようになりました。", vbInformation, "馬吉：設定変更の確認")
                gApp.R_JVLGetBLOD = True
            Else
                chkBLOD.value = 0
            End If
        End If
    End If
    gApp.R_JVLGetBLOD = (chkBLOD.value = 1)
    Exit Sub
Errorhandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: データ同期処理をする
'
'   備考: なし
'
Private Sub SetAndSyncSLOP()
On Error GoTo Errorhandler
    Dim DBUpdate As frmDBUpdate
    Dim MsgResult As VbMsgBoxResult

    gApp.Log "Config SLOP option: " & chkSLOP.value
    If gApp.R_JVLGetSLOP = False And chkSLOP.value = 1 And gApp.R_JVDLastTime <> String$(14, "0") Then
        If gApp.R_JVDLastTimeSLOP = String$(14, "0") Then
            MsgResult = MsgBox("坂路調教データはまだセットアップされていません。" & vbCrLf & _
                        "坂路調教データのセットアップを開始しますか？", vbYesNo + vbQuestion, "馬吉：坂路調教データセットアップの確認")
        Else
            MsgResult = MsgBox("不足している分の坂路調教データを取得します。", vbOKCancel + vbInformation, "馬吉：坂路調教データセットアップの確認")
        End If
        
        If MsgResult = vbCancel Then
            chkSLOP.value = 0
        Else
            Set DBUpdate = New frmDBUpdate
            DBUpdate.GettingMode = 1 '0:Other 1:SLOP 2:BLOD 3:O6H6 else:err
            DBUpdate.Show vbModal, Me
            If DBUpdate.Finish Then
                MsgResult = MsgBox("次回の更新からは、坂路調教データも含まれるようになりました。", vbInformation, "馬吉：坂路調教データ")
                gApp.R_JVLGetBLOD = True
            Else
                chkSLOP.value = 0
            End If
        End If
    End If
    gApp.R_JVLGetSLOP = (chkSLOP.value = 1)
    Exit Sub
Errorhandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: キャンセルボタンクリックイベント
'
'   備考: 編集がなにもなされていなければ即終了、なされていれば確認
'
Private Sub cmdCancel_Click(Index As Integer)
On Error GoTo Errorhandler
    If Index = 1 Then
        Call SetDisplayColors(mlngPrevBGColor, mlngPrevFGColor)
    End If
    
    gApp.R_DBPath = txtPath.tag
    
    Unload Me
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   機能: カラーカスタマイズボタンイベント
'
'   備考: なし
'
Private Sub cmdColor_Click(Index As Integer)
On Error GoTo Errorhandler
    dlgColorChoice.CancelError = True
    dlgColorChoice.Flags = cdlCCRGBInit 'フラグの設定
    dlgColorChoice.color = IIf(Index = 0, lblBGDisp.BackColor, lblFGDisp.BackColor) 'ダイアログの初期色を設定
    dlgColorChoice.ShowColor
    If Err.Number = 32755 Then Exit Sub
    Select Case Index
    Case 0:
        mlngBGColor = dlgColorChoice.color
    Case 1:
        mlngFGColor = dlgColorChoice.color
    End Select
    Call SetDisplayColors(mlngBGColor, mlngFGColor)

    Exit Sub
Errorhandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: 参照ボタンイベント
'
'   備考: なし
'
Private Sub cmdDirRef_Click()
On Error GoTo Errorhandler
    On Error GoTo Errorhandler
    Dim FolderSelectionDialog As frmDirRef
    
    Set FolderSelectionDialog = New frmDirRef
    With FolderSelectionDialog
        .BeginingPath = txtPath.Text
        .Show vbModal, Me
        txtPath.Text = .ReturnPath
    End With
    Call DBChange
    
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   機能: JV-Link設定ダイアログボックスの表示
'
'   備考: なし
'
Private Sub cmdJVSetUIProperties_Click()
On Error GoTo Errorhandler
    Dim ReturnCode As Long              ''JVLink返値
    Dim JVlink As frmWrappedJVLink
    
    'JVLink設定画面表示
    Set JVlink = New frmWrappedJVLink
    
    On Error Resume Next
    Load JVlink
    If Err.Number <> "0" Then
        gApp.ErrLog
        MsgBox "JV-Linkがインストールされていません。", vbOKOnly + vbExclamation, "馬吉：JV-Linkエラー"
        Exit Sub
    End If
    On Error GoTo Errorhandler
    
    ReturnCode = JVlink.axJVLink.JVSetUIProperties()

    'エラー判定
    If ReturnCode <> 0 Then         ''エラー
        Call MsgBox("JV-Linkの設定ができませんでした。(" & ReturnCode & ")", vbOKOnly + vbCritical, "馬吉：JV-Linkエラー")
    End If

    Exit Sub
Errorhandler:
    gApp.ErrLog
    Resume Next
    Exit Sub
End Sub


'
'   機能: OKボタンクリックイベント
'
'   備考: レジストリに書き込んで終了
'
Private Sub cmdOK_Click(Index As Integer)
On Error GoTo Errorhandler
    Dim i As Long
    Dim fso As New FileSystemObject
    
    Dim f1 As Folder
    Dim f2 As Folder
    Dim cmdb As New clsCreateMDB
    
    Dim result As Long      '' MsgBoxの戻り値
    Dim DirRef As frmDirRef '' フォルダ参照ダイアログ
    
    Set DirRef = New frmDirRef
    DirRef.Message = "データベースフォルダを指定してください"

    gApp.R_DBPath = fso.GetFolder(txtPath.Text).Path
    
    Do While MissingDBNum <> 0

        ' データベースがあるか調べる
        If MissingDBNum = -1 Then
            result = MsgBox(gApp.R_DBPath & "に、データベースを新規作成しますか？" _
                            , vbYesNoCancel + vbQuestion, "馬吉：新規データベース作成の確認")
            If result = vbYes Then
                If cmdb.createMDB = False Then
                    ' 作成失敗ならフォルダ選択
                    DirRef.BeginingPath = gApp.R_DBPath
                    DirRef.Show vbModal
                    
                    gApp.R_DBPath = DirRef.ReturnPath
                End If
            ElseIf result = vbCancel Then
                ' 元に戻す
                gApp.R_DBPath = txtPath.tag
                Exit Sub
            ElseIf result = vbNo Then
                ' なければ、フォルダ選択
                DirRef.BeginingPath = gApp.R_DBPath
                DirRef.Show vbModal
                
                gApp.R_DBPath = DirRef.ReturnPath
            End If
        Else
            MsgBox "データベースが壊れています。他のフォルダを指定してください。", vbExclamation, "馬吉:エラー"
            ' なければ、フォルダ選択
            DirRef.BeginingPath = gApp.R_DBPath
            DirRef.Show vbModal
            
            gApp.R_DBPath = DirRef.ReturnPath
        End If
    Loop
    
    gApp.R_JVLMode = IIf(optJVMode(0).value, ukJVLMode.ukjUsual, ukJVLMode.ukjThisWeek)

    gApp.R_BackColorDark = mlngFGColor
    gApp.R_BackColorLight = mlngBGColor

    Call SetAndSyncSLOP
    Call SetAndSyncBLOD

    ' 未編集状態にしてアンロード（QueryUnload時に未編集で無い場合破棄確認をするため）
    mblnNoEdit = True
    
    Unload Me
    Exit Sub
Errorhandler:
    gApp.ErrLog
    Resume Next
End Sub


'
'   機能: 色を逆にするボタンイベント
'
'   備考: なし
'
Private Sub cmdReverse_Click()
On Error GoTo Errorhandler
    Call SetDisplayColors(mlngFGColor, mlngBGColor)
    mlngBGColor = lblBGDisp.BackColor
    mlngFGColor = lblFGDisp.BackColor
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   機能: フォームの初期化
'
'   備考: レジストリ情報をコントロールに設定する
'
Private Sub Form_Load()
On Error GoTo Errorhandler
    Me.Icon = LoadResPicture(100, vbResIcon)
    mblnNoEdit = True ' 未編集状態
    
    ' DBパス
    txtPath.Text = gApp.R_DBPath
    ' tagプロパティに初期値を保存する
    txtPath.tag = gApp.R_DBPath
    
    Call setDBStatus
    
    mlngPrevBGColor = gApp.R_BackColorLight
    mlngPrevFGColor = gApp.R_BackColorDark
    mlngBGColor = gApp.R_BackColorLight
    mlngFGColor = gApp.R_BackColorDark

    Call SetDisplayColors(mlngBGColor, mlngFGColor)
    
    fraAdvanced.Enabled = False
    cmdColor(0).Enabled = False
    cmdColor(1).Enabled = False
    
    Call ColorThemeCheck
    
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   機能: ＤＢの状態を設定する
'
'   備考: なし
'
Private Sub setDBStatus()
On Error GoTo Errorhandler
    ' JV-Link取得モード
    optJVMode(0).value = (gApp.R_JVLMode = ukjUsual)
    optJVMode(1).value = (gApp.R_JVLMode = ukjThisWeek)
    
    ' 坂路、血統
    chkSLOP.value = IIf(gApp.R_JVLGetSLOP, 1, 0)
    chkBLOD.value = IIf(gApp.R_JVLGetBLOD, 1, 0)
    lblFix(1).Visible = (gApp.R_SetupCancelLastTime <> "")
    If gApp.R_SetupCancelLastTime <> "" Then
        chkSLOP.Enabled = False
        chkBLOD.Enabled = False
    Else
        chkSLOP.Enabled = True
        chkBLOD.Enabled = True
    End If
    
    If gApp.R_JVDLastTime <> "00000000000000" Then
        lblInfo(0).Caption = "最終データ取得日時： " _
            & Format$(gApp.R_JVDLastTime, "@@@@/@@/@@ @@:@@:@@")
    ElseIf gApp.R_SetupCancelLastTime <> "" Then
        lblInfo(0).Caption = "セットアップ中断中"
    Else
        lblInfo(0).Caption = "未取得"
    End If
    
    If gApp.R_JVDLastTime = "00000000000000" And gApp.R_SetupCancelLastTime = "" Then
        lblInfo(1).Caption = ""
    ElseIf gApp.R_SetupYear > 0 Then
        lblInfo(1).Caption = "セットアップ開始年度： " & gApp.R_SetupYear & "年"
    Else
        lblInfo(1).Caption = "セットアップ開始年度： 全部"
    End If
    
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   機能: フォームアンロードキャンセル確認
'
'   備考: 編集を破棄して終了するかどうか。OKボタンで終了すれば、破棄確認は出ない。
'
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo Errorhandler
If Not mblnNoEdit Then
        If MsgBox("設定の変更を破棄し、変更前にもどして、設定を終了しますか", vbYesNo + vbQuestion, "馬吉：変更破棄の確認") = vbNo Then
            Cancel = 1
        End If
    End If
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   機能: 表示色を設定する
'
'   備考: なし
'
Private Sub SetDisplayColors(BGColor As Long, fgcolor As Long)
    lblBGDisp.BackColor = BGColor
    lblBGDisp.ForeColor = Contrast(BGColor)
    lblFGDisp.BackColor = fgcolor
    lblFGDisp.ForeColor = Contrast(fgcolor)
End Sub


'
'   機能: JV-Link取得モード設定イベント
'
'   備考: なし
'
Private Sub optJVMode_Click(Index As Integer)
On Error GoTo Errorhandler
    chkSLOP.Enabled = (optJVMode(0).value And gApp.R_SetupCancelLastTime = "")
    chkBLOD.Enabled = (optJVMode(0).value And gApp.R_SetupCancelLastTime = "")
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   機能: 色のテーマ設定イベント
'
'   備考: なし
'
Private Sub optTheme_Click(Index As Integer)
On Error GoTo Errorhandler
    Dim lngBack As Long
    Dim lngFore As Long
    
    Select Case Index
        Case 0: ' ハロウィン
            lngBack = RGB(238, 102, 51)
            lngFore = RGB(1, 1, 1)
        Case 1: ' クールミント
            lngBack = RGB(151, 255, 255)
            lngFore = RGB(0, 205, 205)
        Case 2: ' 秋
            lngBack = RGB(255, 204, 0)
            lngFore = RGB(255, 69, 0)
        Case 3: ' 血のように赤い
            lngBack = RGB(238, 51, 51)
            lngFore = RGB(153, 1, 1)
        Case 4: ' プリーティーピンク
            lngBack = RGB(255, 182, 193)
            lngFore = RGB(240, 128, 128)
        Case 5: ' ディフォルト (UK Default)
            lngBack = RGB(238, 238, 224) ' &HE0EEEE
            lngFore = RGB(204, 204, 192) ' &HC0CCCC
        Case 6: ' カスタマイズ
            fraAdvanced.Enabled = True
            cmdColor(0).Enabled = True
            cmdColor(1).Enabled = True
    End Select
    If optTheme(Index).value = True And Index <> 6 Then
        fraAdvanced.Enabled = False
        cmdColor(0).Enabled = False
        cmdColor(1).Enabled = False
        mlngBGColor = lngBack
        mlngFGColor = lngFore
    End If
    Call SetDisplayColors(mlngBGColor, mlngFGColor)
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   機能: 色のテーマをチェック
'
'   備考: なし
'
Private Sub ColorThemeCheck()
On Error GoTo Errorhandler
    Dim Index As Long
    Dim BG As Long
    Dim fg As Long
    
    BG = gApp.R_BackColorLight
    fg = gApp.R_BackColorDark
    
    If BG = RGB(238, 102, 51) And _
       fg = RGB(1, 1, 1) Then
        Index = 0 ' ハロウィン
    ElseIf BG = RGB(151, 255, 255) And _
       fg = RGB(0, 205, 205) Then
        Index = 1  ' クールミント
    ElseIf BG = RGB(255, 204, 0) And _
       fg = RGB(255, 69, 0) Then
        Index = 2  ' 秋
    ElseIf BG = RGB(238, 51, 51) And _
       fg = RGB(153, 1, 1) Then
        Index = 3  ' 血のように赤い
    ElseIf BG = RGB(255, 182, 193) And _
       fg = RGB(240, 128, 128) Then
        Index = 4  ' プリーティーピンク
    ElseIf BG = &HE0EEEE And _
       fg = &HC0CCCC Then
        Index = 5  ' ディフォルト (UK Default)
    Else
        Index = 6  ' カスタマイズ
        fraAdvanced.Enabled = True
        cmdColor(0).Enabled = True
        cmdColor(1).Enabled = True
    End If
    optTheme(Index).value = True
    
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   機能: データベースフォルダの参照をチェック
'
'   備考: なし
'
Private Sub txtPath_Validate(Cancel As Boolean)
On Error GoTo Errorhandler
    Cancel = Not DBChange
    Exit Sub
Errorhandler:
    gApp.ErrLog
End Sub


'
'   機能: データベースフォルダを選択する
'
'   備考: なし
'
Private Function DBChange() As Boolean
On Error GoTo Errorhandler
    Dim fso As New FileSystemObject
    
    Dim cmdb As New clsCreateMDB
    
    Dim result As VbMsgBoxResult    '' MsgBoxの戻り値
    Dim DirRef As frmDirRef         '' フォルダ参照ダイアログ
    
    DBChange = False
    
    Set DirRef = New frmDirRef
    DirRef.Message = "データベースフォルダを指定してください"
    
    If Not fso.FolderExists(txtPath.Text) Then
        result = MsgBox("フォルダがありません。作成しますか？", vbYesNo + vbQuestion, "馬吉：フォルダ作成の確認")
        If result = vbYes Then
            On Error Resume Next
                fso.CreateFolder (txtPath.Text)
                If Err.Number <> 0 Then
                    MsgBox "フォルダが作成できませんでした。", vbExclamation, "馬吉：エラー"
                    txtPath.Text = txtPath.tag
                    DBChange = False
                    Exit Function
                End If
            On Error GoTo Errorhandler
        ElseIf result = vbNo Then
            txtPath.Text = txtPath.tag
            DBChange = False
            Exit Function
        End If
    End If
    
    gApp.R_DBPath = fso.GetFolder(txtPath.Text).Path
    ' レジストリに保存
    Do While MissingDBNum <> 0

        ' データベースがあるか調べる
        If MissingDBNum = -1 Then
            result = MsgBox(gApp.R_DBPath & "に、データベースを新規作成しますか？" _
                            , vbYesNoCancel + vbQuestion, "馬吉：新規データベース作成の確認")
            If result = vbYes Then
                Call cmdb.createMDB
            ElseIf result = vbCancel Then
                ' 元に戻す
                gApp.R_DBPath = txtPath.tag
                DBChange = False
                Exit Function
            ElseIf result = vbNo Then
                ' なければ、フォルダ選択
                DirRef.BeginingPath = gApp.R_DBPath
                DirRef.Show vbModal
                
                gApp.R_DBPath = DirRef.ReturnPath
            End If
        Else
            MsgBox "データベースが壊れています。他のフォルダを指定してください。", vbExclamation, "馬吉:エラー"
            ' なければ、フォルダ選択
            DirRef.BeginingPath = gApp.R_DBPath
            DirRef.Show vbModal
            
            gApp.R_DBPath = DirRef.ReturnPath
        End If
    Loop
    
    Call setDBStatus
    
    DBChange = True
    
    Exit Function
Errorhandler:
    gApp.ErrLog
End Function
