VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.UserControl ctlVFind 
   BackColor       =   &H00E0EEEE&
   ClientHeight    =   7410
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9360
   ScaleHeight     =   7410
   ScaleWidth      =   9360
   Begin VB.Timer tmrFetch 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   8640
      Top             =   120
   End
   Begin TabDlg.SSTab mstTab 
      Height          =   6405
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   11298
      _Version        =   393216
      Style           =   1
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      BackColor       =   14741230
      TabCaption(0)   =   "競走馬"
      TabPicture(0)   =   "ctlVFind.ctx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "picTab(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "騎手"
      TabPicture(1)   =   "ctlVFind.ctx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "picTab(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "調教師"
      TabPicture(2)   =   "ctlVFind.ctx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "picTab(2)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "馬主"
      TabPicture(3)   =   "ctlVFind.ctx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "picTab(3)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "生産者"
      TabPicture(4)   =   "ctlVFind.ctx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "picTab(4)"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "繁殖馬"
      TabPicture(5)   =   "ctlVFind.ctx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "picTab(5)"
      Tab(5).ControlCount=   1
      Begin VB.PictureBox picTab 
         BorderStyle     =   0  'なし
         Height          =   4095
         Index           =   5
         Left            =   -74940
         ScaleHeight     =   4095
         ScaleWidth      =   6735
         TabIndex        =   37
         Top             =   360
         Width           =   6735
         Begin VB.ComboBox cboMethod 
            Height          =   300
            Index           =   5
            ItemData        =   "ctlVFind.ctx":00A8
            Left            =   3630
            List            =   "ctlVFind.ctx":00B5
            Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
            TabIndex        =   23
            Top             =   0
            Width           =   1335
         End
         Begin VB.TextBox txtName 
            Height          =   300
            Index           =   5
            Left            =   960
            TabIndex        =   22
            Top             =   0
            Width           =   2595
         End
         Begin VB.CommandButton cmdDo 
            Caption         =   "検索実行"
            Height          =   300
            Index           =   5
            Left            =   5040
            TabIndex        =   24
            Top             =   0
            Width           =   975
         End
         Begin Umakichi.ctlPane paneTab 
            Height          =   3495
            Index           =   5
            Left            =   0
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   360
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   6165
            Begin Umakichi.ctlWrappedGrid flexTab 
               Height          =   1800
               Index           =   5
               Left            =   0
               TabIndex        =   25
               Top             =   0
               Width           =   2400
               _ExtentX        =   4233
               _ExtentY        =   3175
            End
         End
         Begin VB.Label lblFix 
            AutoSize        =   -1  'True
            Caption         =   "繁殖馬名："
            Height          =   180
            Index           =   6
            Left            =   60
            TabIndex        =   38
            Top             =   60
            Width           =   810
         End
      End
      Begin VB.PictureBox picTab 
         BorderStyle     =   0  'なし
         Height          =   4095
         Index           =   4
         Left            =   -74940
         ScaleHeight     =   4095
         ScaleWidth      =   6735
         TabIndex        =   35
         Top             =   360
         Width           =   6735
         Begin VB.CommandButton cmdDo 
            Caption         =   "検索実行"
            Height          =   300
            Index           =   4
            Left            =   5040
            TabIndex        =   20
            Top             =   0
            Width           =   975
         End
         Begin VB.TextBox txtName 
            Height          =   300
            Index           =   4
            Left            =   960
            TabIndex        =   18
            Top             =   0
            Width           =   2595
         End
         Begin VB.ComboBox cboMethod 
            Height          =   300
            Index           =   4
            ItemData        =   "ctlVFind.ctx":00D3
            Left            =   3600
            List            =   "ctlVFind.ctx":00E0
            Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
            TabIndex        =   19
            Top             =   0
            Width           =   1335
         End
         Begin Umakichi.ctlPane paneTab 
            Height          =   3495
            Index           =   4
            Left            =   0
            TabIndex        =   43
            TabStop         =   0   'False
            Top             =   360
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   6165
            Begin Umakichi.ctlWrappedGrid flexTab 
               Height          =   1800
               Index           =   4
               Left            =   0
               TabIndex        =   21
               Top             =   0
               Width           =   2400
               _ExtentX        =   4233
               _ExtentY        =   3175
            End
         End
         Begin VB.Label lblFix 
            AutoSize        =   -1  'True
            Caption         =   "生産者名："
            Height          =   180
            Index           =   5
            Left            =   60
            TabIndex        =   36
            Top             =   60
            Width           =   810
         End
      End
      Begin VB.PictureBox picTab 
         BorderStyle     =   0  'なし
         Height          =   4095
         Index           =   3
         Left            =   -74940
         ScaleHeight     =   4095
         ScaleWidth      =   6735
         TabIndex        =   33
         Top             =   360
         Width           =   6735
         Begin VB.ComboBox cboMethod 
            Height          =   300
            Index           =   3
            ItemData        =   "ctlVFind.ctx":00FE
            Left            =   3360
            List            =   "ctlVFind.ctx":010B
            Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
            TabIndex        =   15
            Top             =   0
            Width           =   1335
         End
         Begin VB.TextBox txtName 
            Height          =   300
            Index           =   3
            Left            =   720
            TabIndex        =   14
            Top             =   0
            Width           =   2595
         End
         Begin VB.CommandButton cmdDo 
            Caption         =   "検索実行"
            Height          =   300
            Index           =   3
            Left            =   4800
            TabIndex        =   16
            Top             =   0
            Width           =   975
         End
         Begin Umakichi.ctlPane paneTab 
            Height          =   3495
            Index           =   3
            Left            =   0
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   360
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   6165
            Begin Umakichi.ctlWrappedGrid flexTab 
               Height          =   1800
               Index           =   3
               Left            =   0
               TabIndex        =   17
               Top             =   0
               Width           =   2400
               _ExtentX        =   4233
               _ExtentY        =   3175
            End
         End
         Begin VB.Label lblFix 
            AutoSize        =   -1  'True
            Caption         =   "馬主名："
            Height          =   180
            Index           =   4
            Left            =   60
            TabIndex        =   34
            Top             =   60
            Width           =   630
         End
      End
      Begin VB.PictureBox picTab 
         BorderStyle     =   0  'なし
         Height          =   4095
         Index           =   2
         Left            =   -74940
         ScaleHeight     =   4095
         ScaleWidth      =   6735
         TabIndex        =   31
         Top             =   360
         Width           =   6735
         Begin VB.ComboBox cboMethod 
            Height          =   300
            Index           =   2
            ItemData        =   "ctlVFind.ctx":0129
            Left            =   3600
            List            =   "ctlVFind.ctx":0136
            Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
            TabIndex        =   11
            Top             =   0
            Width           =   1335
         End
         Begin VB.TextBox txtName 
            Height          =   300
            Index           =   2
            Left            =   960
            TabIndex        =   10
            Top             =   0
            Width           =   2595
         End
         Begin VB.CommandButton cmdDo 
            Caption         =   "検索実行"
            Height          =   300
            Index           =   2
            Left            =   5040
            TabIndex        =   12
            Top             =   0
            Width           =   975
         End
         Begin Umakichi.ctlPane paneTab 
            Height          =   3495
            Index           =   2
            Left            =   0
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   360
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   6165
            Begin Umakichi.ctlWrappedGrid flexTab 
               Height          =   1800
               Index           =   2
               Left            =   0
               TabIndex        =   13
               Top             =   0
               Width           =   2400
               _ExtentX        =   4233
               _ExtentY        =   3175
            End
         End
         Begin VB.Label lblFix 
            AutoSize        =   -1  'True
            Caption         =   "調教師名："
            Height          =   180
            Index           =   3
            Left            =   60
            TabIndex        =   32
            Top             =   60
            Width           =   810
         End
      End
      Begin VB.PictureBox picTab 
         BorderStyle     =   0  'なし
         Height          =   4095
         Index           =   1
         Left            =   -74940
         ScaleHeight     =   4095
         ScaleWidth      =   6735
         TabIndex        =   29
         Top             =   360
         Width           =   6735
         Begin VB.CommandButton cmdDo 
            Caption         =   "検索実行"
            Height          =   300
            Index           =   1
            Left            =   4800
            TabIndex        =   8
            Top             =   0
            Width           =   975
         End
         Begin VB.TextBox txtName 
            Height          =   300
            Index           =   1
            Left            =   720
            TabIndex        =   6
            Top             =   0
            Width           =   2595
         End
         Begin VB.ComboBox cboMethod 
            Height          =   300
            Index           =   1
            ItemData        =   "ctlVFind.ctx":0154
            Left            =   3360
            List            =   "ctlVFind.ctx":0161
            Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
            TabIndex        =   7
            Top             =   0
            Width           =   1335
         End
         Begin Umakichi.ctlPane paneTab 
            Height          =   3495
            Index           =   1
            Left            =   0
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   360
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   6165
            Begin Umakichi.ctlWrappedGrid flexTab 
               Height          =   1800
               Index           =   1
               Left            =   0
               TabIndex        =   9
               Top             =   0
               Width           =   2400
               _ExtentX        =   4233
               _ExtentY        =   3175
            End
         End
         Begin VB.Label lblFix 
            AutoSize        =   -1  'True
            Caption         =   "騎手名："
            Height          =   180
            Index           =   2
            Left            =   60
            TabIndex        =   30
            Top             =   60
            Width           =   630
         End
      End
      Begin VB.PictureBox picTab 
         BorderStyle     =   0  'なし
         Height          =   4095
         Index           =   0
         Left            =   60
         ScaleHeight     =   4095
         ScaleWidth      =   6735
         TabIndex        =   26
         Top             =   360
         Width           =   6735
         Begin VB.ComboBox cboMethod 
            Height          =   300
            Index           =   0
            ItemData        =   "ctlVFind.ctx":017F
            Left            =   3210
            List            =   "ctlVFind.ctx":018C
            Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
            TabIndex        =   3
            Top             =   0
            Width           =   1335
         End
         Begin VB.TextBox txtName 
            Height          =   300
            Index           =   0
            Left            =   540
            TabIndex        =   2
            Top             =   0
            Width           =   2595
         End
         Begin VB.CommandButton cmdDo 
            Caption         =   "検索実行"
            Height          =   300
            Index           =   0
            Left            =   4620
            TabIndex        =   4
            Top             =   0
            Width           =   975
         End
         Begin Umakichi.ctlPane paneTab 
            Height          =   3495
            Index           =   0
            Left            =   0
            TabIndex        =   40
            TabStop         =   0   'False
            Top             =   360
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   6165
            Begin Umakichi.ctlWrappedGrid flexTab 
               Height          =   1800
               Index           =   0
               Left            =   0
               TabIndex        =   5
               Top             =   0
               Width           =   2400
               _ExtentX        =   4233
               _ExtentY        =   3175
            End
         End
         Begin VB.Label lblFix 
            AutoSize        =   -1  'True
            Caption         =   "馬名："
            Height          =   180
            Index           =   1
            Left            =   60
            TabIndex        =   28
            Top             =   60
            Width           =   450
         End
      End
   End
   Begin VB.Frame fraTop 
      BackColor       =   &H00C0CCCC&
      BorderStyle     =   0  'なし
      Caption         =   "fraTop"
      Height          =   405
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   8955
      Begin VB.Label lblFix 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0CCCC&
         Caption         =   "マスター情報"
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
         Left            =   180
         TabIndex        =   27
         Top             =   90
         Width           =   1395
      End
   End
End
Attribute VB_Name = "ctlVFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   検索表示コントロール
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数(イベント)
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Public Event ChangeTo(strViewerName As String, strKey As String)    '' Vierer変更イベント
Public Event WindowTitle(strKey As String)                          '' ウインドウタイトル変更イベント
Public Event LinkContextMenu(strViewerName As String, strKey As String)

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mVB As clsViewerBase
Private mViewerState As clsVSFind
Private mblnNoData As Boolean           '' データ無しフラグ

Private mstrTitle As String              '' ウインドウタイトル
Private mstrKey As String                '' キー
Private WithEvents mData As clsDataFind  '' データ取得オブジェクト
Attribute mData.VB_VarHelpID = -1
Private mIndex As Integer

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部定数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Const MINIMUMWIDTH  As Long = 6400
Const MINIMUMHEIGHT As Long = 6400

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   プロパティ
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8


'
'   機能: キー設定プロパティ
'
'   備考: Viewer必須プロパティ
'
Public Property Let Key(strKey As String)
On Error GoTo ErrorHandler
    Dim i As Long
    
    If mViewerState.IsNoTouch Then
        mstrKey = strKey
        mstTab.Tab = CLng(strKey)
    Else
        mstrKey = mViewerState.LastTabNumber
        mstTab.Tab = mViewerState.LastTabNumber
        For i = 0 To 5
            txtName(i).Text = mViewerState.KeyWord(i)
            cboMethod(i).ListIndex = mViewerState.FindMode(i)
        Next i
    End If
    Exit Property
ErrorHandler:
    gApp.ErrLog
    Resume Next
End Property


'
'   機能: タイトル取得プロパティ
'
'   備考: Viewer必須プロパティ、　Browser が呼びます
'
Public Property Get Title() As String
    Title = mstrTitle
End Property


'
'   機能: タイトル設定プロパティ
'
'   備考: ブラウザに変更通知のイベントを投げます
'
Public Property Let Title(strTitle As String)
    mstrTitle = strTitle
    RaiseEvent WindowTitle(mstrTitle)
End Property


'
'   機能: Viewer状態提供
'
'   備考: なし
'
Public Property Get ViewerState() As clsVSFind
    Set ViewerState = mViewerState
End Property


'
'   機能: Viewer状態受け取り
'
'   備考: なし
'
Public Property Set ViewerState(RHS As clsVSFind)
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
'   機能: 検索実行ボタンクリックイベント
'
'   備考: なし
'
Private Sub cmdDo_Click(Index As Integer)
On Error GoTo ErrorHandler
    Dim i As Long
    If txtName(Index).Text <> Empty Then
        
        ' パネルのモード変更
        For i = 0 To 5
            If i = Index Then
                ' 検索対象は取得中に設定
                paneTab(i).Mode = ukcpNowFetching
            ElseIf paneTab(i).Mode = ukcpNowFetching Then
                ' 取得中だったものは非表示に設定
                paneTab(i).Mode = ukcpHideControls
            End If
        Next i
        
        mIndex = Index
        If mData.NowFetching = True Then
            Call mData.CancelFind
        End If
        tmrFetch.tag = txtName(Index).Text
        tmrFetch.Enabled = True
    End If
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 検索モード変更
'
'   備考: なし
'
Private Sub cboMethod_click(Index As Integer)
On Error GoTo ErrorHandler
    If cboMethod(Index).Enabled = True Then
        mViewerState.FindMode(CLng(Index)) = cboMethod(Index).ListIndex
    End If
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: Fetchタイマーイベント
'
'   備考: なし
'
Private Sub tmrFetch_Timer()
On Error GoTo ErrorHandler
    gApp.Log "Cancel Waiting: " & tmrFetch.tag
    
    If Not mData.NowFetching Then
        tmrFetch.Enabled = False
        Call mData.Fetch(CLng(mIndex), tmrFetch.tag, cboMethod(mIndex).ListIndex)
    End If
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 検索キーワード変更
'
'   備考: なし
'
Private Sub txtName_Change(Index As Integer)
On Error GoTo ErrorHandler
    mViewerState.LastTabNumber = mstTab.Tab
    mViewerState.KeyWord(CLng(Index)) = txtName(Index).Text
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


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
'   機能: マウスの下がリンク可能なグリッドならば反応する為のイベント
'
'   備考: 標準的な反応は、clsGridData.MouseMoveDrivenプロシージャに任せる
'
Private Sub flexTab_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrorHandler
    
    Call flexTab(Index).ReflexiveMouseMoveDriven(True)
    
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
'   機能: 馬主タブのデータがそろった
'
'   備考: なし
'
Private Sub mData_FetchCompleteBANUSI(gd As clsGridData)
On Error GoTo ErrorHandler
    Call flexTab(3).InsertGrid(gd)
    flexTab(3).Grid.col = 2
    Call flexTab(3).AutoSize(0, flexTab(3).Grid.Cols - 1)
    
    If flexTab(3).Grid.Rows < 2 Then
        Call flexTab(3).FlexDisable
    Else
        Call flexTab(3).FlexDisable(False)
    End If
        
    paneTab(3).Mode = ukcpShowControls
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 調教師タブのデータがそろった
'
'   備考: なし
'
Private Sub mData_FetchCompleteCHOKYO(gd As clsGridData)
On Error GoTo ErrorHandler
    Call flexTab(2).InsertGrid(gd)
    flexTab(2).Grid.col = 1
    Call flexTab(2).AutoSize(0, flexTab(2).Grid.Cols - 1)
    
    If flexTab(2).Grid.Rows < 2 Then
        Call flexTab(2).FlexDisable
    Else
        Call flexTab(2).FlexDisable(False)
    End If
        
    paneTab(2).Mode = ukcpShowControls
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 繁殖馬タブのデータがそろった
'
'   備考: なし
'
Private Sub mData_FetchCompleteHANSYOKU(gd As clsGridData)
On Error GoTo ErrorHandler
    Call flexTab(5).InsertGrid(gd)
    flexTab(5).Grid.col = 1
    Call flexTab(5).AutoSize(0, flexTab(5).Grid.Cols - 1)
    
    If flexTab(5).Grid.Rows < 2 Then
        Call flexTab(5).FlexDisable
    Else
        Call flexTab(5).FlexDisable(False)
    End If
    
    paneTab(5).Mode = ukcpShowControls
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 騎手タブのデータがそろった
'
'   備考: なし
'
Private Sub mData_FetchCompleteKISHU(gd As clsGridData)
On Error GoTo ErrorHandler
    Call flexTab(1).InsertGrid(gd)
    flexTab(1).Grid.col = 1
    Call flexTab(1).AutoSize(0, flexTab(1).Grid.Cols - 1)
    
    If flexTab(1).Grid.Rows < 2 Then
        Call flexTab(1).FlexDisable
    Else
        Call flexTab(1).FlexDisable(False)
    End If
    
    paneTab(1).Mode = ukcpShowControls
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 生産者タブのデータがそろった
'
'   備考: なし
'
Private Sub mData_FetchCompleteSEISAN(gd As clsGridData)
On Error GoTo ErrorHandler
    Call flexTab(4).InsertGrid(gd)
    flexTab(4).Grid.col = 2
    Call flexTab(4).AutoSize(0, flexTab(4).Grid.Cols - 1)
    
    If flexTab(4).Grid.Rows < 2 Then
        Call flexTab(4).FlexDisable
    Else
        Call flexTab(4).FlexDisable(False)
    End If
    
    paneTab(4).Mode = ukcpShowControls
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 競走馬タブのデータがそろった
'
'   備考: なし
'
Private Sub mData_FetchCompleteUMA(gd As clsGridData)
On Error GoTo ErrorHandler
    Call flexTab(0).InsertGrid(gd)
    flexTab(0).Grid.col = 0
    Call flexTab(0).AutoSize(0, flexTab(0).Grid.Cols - 1)
    
    If flexTab(0).Grid.Rows < 2 Then
        Call flexTab(0).FlexDisable
    Else
        Call flexTab(0).FlexDisable(False)
    End If
    
    paneTab(0).Mode = ukcpShowControls
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 馬主のデータがない
'
'   備考: なし
'
Private Sub mData_NoDataBANUSI()
On Error GoTo ErrorHandler
    paneTab(3).Mode = ukcpNoData
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 調教師のデータがない
'
'   備考: なし
'
Private Sub mData_NoDataCHOKYO()
On Error GoTo ErrorHandler
    paneTab(2).Mode = ukcpNoData
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 繁殖馬のデータがない
'
'   備考: なし
'
Private Sub mData_NoDataHANSYOKU()
On Error GoTo ErrorHandler
    paneTab(5).Mode = ukcpNoData
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 騎手のデータがない
'
'   備考: なし
'
Private Sub mData_NoDataKISHU()
On Error GoTo ErrorHandler
    paneTab(1).Mode = ukcpNoData
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub

'
'   機能: 生産者のデータがない
'
'   備考: なし
'
Private Sub mData_NoDataSEISAN()
On Error GoTo ErrorHandler
    paneTab(4).Mode = ukcpNoData
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 競走馬のデータがない
'
'   備考: なし
'
Private Sub mData_NoDataUMA()
On Error GoTo ErrorHandler
    paneTab(0).Mode = ukcpNoData
    
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
    
    mViewerState.LastTabNumber = mstTab.Tab
    
    For i = picTab.LBound To picTab.UBound
        picTab(i).Visible = (i = mstTab.Tab)
    Next i
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 検索語の入力イベント
'
'   備考: なし
'
Private Sub txtName_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error GoTo ErrorHandler
    Select Case KeyCode
        Case vbKeyReturn:
            Call cmdDo_Click(Index)
    End Select
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
    Set mData = New clsDataFind
    Set mVB = New clsViewerBase
    Set mViewerState = New clsVSFind
        
    'cboMethod初期化
    For i = cboMethod.LBound To cboMethod.UBound
        cboMethod(i).Enabled = False
        cboMethod(i).ListIndex = 0
        cboMethod(i).Enabled = True
    Next i
    
    ' FlexGrid設定
    For i = flexTab.LBound To flexTab.UBound
        flexTab(i).Grid.FixedRows = 1
        flexTab(i).Grid.FixedCols = 0
        flexTab(i).Grid.Cols = 0
        flexTab(i).Grid.Rows = 1
        Call mVB.FlexGridCommonSetting(flexTab(i).Grid)
        
    Next i
    
    For i = picTab.LBound To picTab.UBound
        picTab(i).Visible = (i = mstTab.Tab)
    Next i
    
    ' すべてのペインを、データ取得中に設定する。
    For i = paneTab.LBound To paneTab.UBound
        paneTab(i).Mode = ukcpHideControls
        With paneTab(i)
            .width = Bigger(1, picTab(i).width - .Left)
            .Height = Bigger(1, picTab(i).Height - .Top)
        End With ' paneTab(mstTab.Index)
    Next i
    
    ' Color Assign
    BackColor = gApp.ColBG
    mstTab.BackColor = gApp.ColBG
    fraTop.BackColor = gApp.ColDarkBG
    lblFix(0).BackColor = gApp.ColDarkBG
    lblFix(0).ForeColor = Contrast(gApp.ColDarkBG)
    
    mstrTitle = "検索"
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: ユーザコントロールのりサイズイベント
'
'   備考: なし
'
Private Sub UserControl_Resize()
On Error GoTo ErrorHandler
    Dim i As Long
    
    ' 最小幅設定
    With UserControl
        .width = Bigger(.width, MINIMUMWIDTH)
        .Height = Bigger(.Height, MINIMUMHEIGHT)
    End With
    
    ' 共通UIリサイズ
    fraTop.width = ScaleWidth - fraTop.Left * 2
    mstTab.width = ScaleWidth - mstTab.Left * 2
    mstTab.Height = ScaleHeight - mstTab.Top - fraTop.Top
    
    For i = 0 To picTab.count - 1
        With picTab(i)
            .Top = mstTab.TabHeight + 60
            .Left = 60
            .width = Bigger(1, mstTab.width - .Left * 2)
            .Height = Bigger(1, (mstTab.Height - .Top) - (.Top - mstTab.TabHeight))
        End With ' fraTab(mstTab.Index)
        
        With paneTab(i)
            .width = Bigger(1, picTab(i).width - .Left)
            .Height = Bigger(1, picTab(i).Height - .Top)
        End With ' paneTab(mstTab.Index)
        
        With flexTab(i)
            .width = Bigger(1, paneTab(i).width - .Left)
            .Height = Bigger(1, paneTab(i).Height - .Top)
        End With ' flexTab(mstTab.Index)
    Next i
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: ユーザコントロール終了イベント
'
'   備考: なし
'
Private Sub UserControl_Terminate()
On Error GoTo ErrorHandler
    gApp.TermLog Me

    Set mData = Nothing
    Set mVB = Nothing
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
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
    If Not (mData Is Nothing) Then mData.CancelFetching
    Dim i As Integer
    For i = flexTab.LBound To flexTab.UBound
        Call DestroyFlexGrid(flexTab(i))
    Next i
    Set mData = Nothing
End Sub


