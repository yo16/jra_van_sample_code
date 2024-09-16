VERSION 5.00
Object = "{1C0489F8-9EFD-423D-887A-315387F18C8F}#1.0#0"; "vsflex8l.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.UserControl ctlVHC 
   BackColor       =   &H00E0EEEE&
   ClientHeight    =   5745
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8655
   ScaleHeight     =   5745
   ScaleWidth      =   8655
   Begin VB.Frame fraTop 
      BackColor       =   &H00C0CCCC&
      BorderStyle     =   0  'なし
      Caption         =   "Frame1"
      Height          =   465
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8955
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0CCCC&
         Caption         =   "3003年 4月 18日"
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
         TabIndex        =   1
         Top             =   120
         Width           =   1905
      End
   End
   Begin TabDlg.SSTab mstTab 
      Height          =   2985
      Left            =   240
      TabIndex        =   2
      Top             =   840
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   5265
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   14741230
      TabCaption(0)   =   "坂路調教"
      TabPicture(0)   =   "ctlVHC.ctx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraTab(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame fraTab 
         BorderStyle     =   0  'なし
         Height          =   2295
         Index           =   0
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   5325
         Begin VSFlex8LCtl.VSFlexGrid flexTab 
            Height          =   1125
            Index           =   0
            Left            =   300
            TabIndex        =   4
            Top             =   120
            Width           =   2025
            _cx             =   3572
            _cy             =   1984
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483636
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   ""
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   0
            AutoSearch      =   0
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   0
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   0
            ShowComboButton =   1
            WordWrap        =   0   'False
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            ComboSearch     =   3
            AutoSizeMouse   =   -1  'True
            FrozenRows      =   0
            FrozenCols      =   0
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin VSFlex8LCtl.VSFlexGrid flexTab 
            Height          =   1125
            Index           =   1
            Left            =   2430
            TabIndex        =   5
            Top             =   210
            Width           =   2025
            _cx             =   3572
            _cy             =   1984
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MousePointer    =   0
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483636
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   ""
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   0
            AutoSearch      =   0
            AutoSearchDelay =   2
            MultiTotals     =   -1  'True
            SubtotalPosition=   1
            OutlineBar      =   0
            OutlineCol      =   0
            Ellipsis        =   0
            ExplorerBar     =   0
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   0
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   0
            ShowComboButton =   1
            WordWrap        =   0   'False
            TextStyle       =   0
            TextStyleFixed  =   0
            OleDragMode     =   0
            OleDropMode     =   0
            ComboSearch     =   3
            AutoSizeMouse   =   -1  'True
            FrozenRows      =   0
            FrozenCols      =   0
            AllowUserFreezing=   0
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
      End
   End
End
Attribute VB_Name = "ctlVHC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' @(h) ctlVHC.ctl
'
' @(s)
'   坂路一覧 表示コントロール
'
Option Explicit

Public Event ChangeTo(strViewerName As String, strKey As String)
Public Event WindowTitle(strKey As String)

Private mVB As clsViewerBase

Public mstrTitle As String


Public Property Let Key(strKey As String)
'    Label1.Caption = strKey
End Property

Public Property Get Title() As String
    Title = mstrTitle
End Property

Public Property Let Title(strTitle As String)
    mstrTitle = strTitle
    RaiseEvent WindowTitle(mstrTitle)
End Property

Private Sub UserControl_Initialize()
On Error GoTo ErrorHandler
    gApp.InitLog Me

    Dim i As Long
    Set mVB = New clsViewerBase
        
    mstrTitle = "坂路一覧"
    
    Call snap
    
    'FlexGrid設定
    For i = flexTab.LBound To flexTab.UBound
        Call mVB.FlexGridCommonSetting(flexTab(i))
    Next i
    
    ' Color Assign
    BackColor = gApp.ColBG
    mstTab.BackColor = gApp.ColBG
    fraTop.BackColor = gApp.ColDarkBG
    
    lblInfo(0).BackColor = gApp.ColDarkBG
    lblInfo(0).ForeColor = Contrast(gApp.ColDarkBG)
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub

Private Sub snap()
    Dim i As Integer
    
    With flexTab(0)
       .Rows = 4
       .Cols = 9
    '   .CellAlignment = flexAlignLeftCenter
       .ColWidth(0) = 0
        i = 0
            .TextMatrix(i, 1) = "調教年月日"
            .TextMatrix(i, 2) = "トレセン"
            .TextMatrix(i, 3) = "調教時刻"
            .TextMatrix(i, 4) = "競走馬"
            .TextMatrix(i, 5) = "4ﾊﾛﾝ(ﾗｯﾌﾟ)"
            .TextMatrix(i, 6) = "3ﾊﾛﾝ(ﾗｯﾌﾟ)"
            .TextMatrix(i, 7) = "2ﾊﾛﾝ(ﾗｯﾌﾟ)"
            .TextMatrix(i, 8) = "1ﾊﾛﾝ(ﾗｯﾌﾟ)"
        i = 1
            .TextMatrix(i, 1) = "2003/04/15"
            .TextMatrix(i, 2) = "   美浦    "
            .TextMatrix(i, 3) = "99:99  "
            .TextMatrix(i, 4) = "シルヴァーアーチャー　　　　　　　　                  "
            .TextMatrix(i, 5) = "99.9(99.9) "
            .TextMatrix(i, 6) = "99.9(99.9) "
            .TextMatrix(i, 7) = "99.9(99.9) "
            .TextMatrix(i, 8) = "99.9(99.9)"
        i = 2
            .TextMatrix(i, 1) = "2003/04/15"
            .TextMatrix(i, 2) = "   美浦    "
            .TextMatrix(i, 3) = "99:99  "
            .TextMatrix(i, 4) = "ティウチェフ　　　　　　　　　　　　                  "
            .TextMatrix(i, 5) = "99.9(99.9) "
            .TextMatrix(i, 6) = "99.9(99.9) "
            .TextMatrix(i, 7) = "99.9(99.9) "
            .TextMatrix(i, 8) = "99.9(99.9)"
        i = 3
            .TextMatrix(i, 1) = "2003/04/15"
            .TextMatrix(i, 2) = "   美浦    "
            .TextMatrix(i, 3) = "99:99  "
            .TextMatrix(i, 4) = "ゴールデンレリーフ　　　　　　　　　                  "
            .TextMatrix(i, 5) = "99.9(99.9) "
            .TextMatrix(i, 6) = "99.9(99.9) "
            .TextMatrix(i, 7) = "99.9(99.9) "
            .TextMatrix(i, 8) = "99.9(99.9)"
    End With
End Sub

Private Sub UserControl_Resize()
On Error GoTo ErrorHandler
    Dim i As Integer
    
    fraTop.Width = Bigger(1, ScaleWidth - fraTop.Left * 2)
    With mstTab
        .Width = Bigger(1, ScaleWidth - .Left * 2)
        .Height = Bigger(1, ScaleHeight - .Top - .Left)
    End With ' mstTab
    
    With fraTab.item(mstTab.Tab)
        .Width = Bigger(1, mstTab.Width - .Left * 2)
        '.Top = mstTab.TabHeight
        .Height = Bigger(1, (mstTab.Height - .Top) - (.Top - mstTab.TabHeight))
    End With ' fraTab.Item(mstTab.Index)
    
    For i = flexTab.LBound To flexTab.UBound
        With flexTab(i)
            .Width = Bigger(1, fraTab(i).Width - .Left)
            .Height = Bigger(1, fraTab(i).Height - .Top)
        End With ' flexTab(i)
    Next i
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub

Private Sub UserControl_Terminate()
On Error GoTo ErrorHandler
    gApp.TermLog Me
    
    Set mVB = Nothing
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub
