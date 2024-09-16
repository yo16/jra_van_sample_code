VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.UserControl ctlVCH 
   BackColor       =   &H00E0EEEE&
   ClientHeight    =   5610
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9330
   ScaleHeight     =   5610
   ScaleWidth      =   9330
   Begin VB.Frame fraHeader 
      BackColor       =   &H00E0EEEE&
      BorderStyle     =   0  'なし
      Caption         =   "fraHeader"
      Height          =   855
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   6555
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H00E0EEEE&
         BorderStyle     =   0  'なし
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Index           =   1
         Left            =   3390
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   6
         Text            =   "ctlVCH.ctx":0000
         Top             =   30
         Width           =   3105
      End
      Begin VB.TextBox txtInfo 
         Appearance      =   0  'ﾌﾗｯﾄ
         BackColor       =   &H00E0EEEE&
         BorderStyle     =   0  'なし
         BeginProperty Font 
            Name            =   "ＭＳ ゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Index           =   0
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   5
         Text            =   "ctlVCH.ctx":001A
         Top             =   30
         Width           =   3315
      End
   End
   Begin VB.Frame fraTop 
      BackColor       =   &H00C0CCCC&
      BorderStyle     =   0  'なし
      Caption         =   "Frame1"
      Height          =   600
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8955
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0CCCC&
         Caption         =   "."
         Height          =   180
         Index           =   1
         Left            =   270
         TabIndex        =   7
         Top             =   60
         Width           =   30
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
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   5010
      End
   End
   Begin TabDlg.SSTab mstTab 
      Height          =   2985
      Left            =   180
      TabIndex        =   2
      Top             =   1800
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   5265
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      BackColor       =   14741230
      TabCaption(0)   =   "総合成績"
      TabPicture(0)   =   "ctlVCH.ctx":0061
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "paneTab(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "距離別成績"
      TabPicture(1)   =   "ctlVCH.ctx":007D
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "paneTab(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "競馬場別成績"
      TabPicture(2)   =   "ctlVCH.ctx":0099
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "paneTab(2)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "主要勝鞍"
      TabPicture(3)   =   "ctlVCH.ctx":00B5
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "paneTab(3)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "競走馬"
      TabPicture(4)   =   "ctlVCH.ctx":00D1
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "paneTab(4)"
      Tab(4).ControlCount=   1
      Begin Umakichi.ctlPane paneTab 
         Height          =   1695
         Index           =   0
         Left            =   60
         TabIndex        =   8
         Top             =   360
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   2990
         Begin Umakichi.ctlWrappedGrid flexTab 
            Height          =   1455
            Index           =   0
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   2566
         End
      End
      Begin Umakichi.ctlPane paneTab 
         Height          =   1695
         Index           =   1
         Left            =   -74520
         TabIndex        =   10
         Top             =   600
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   2990
         Begin Umakichi.ctlWrappedGrid flexTab 
            Height          =   1455
            Index           =   1
            Left            =   0
            TabIndex        =   11
            Top             =   0
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   2566
         End
      End
      Begin Umakichi.ctlPane paneTab 
         Height          =   1695
         Index           =   2
         Left            =   -74760
         TabIndex        =   12
         Top             =   600
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   2990
         Begin Umakichi.ctlWrappedGrid flexTab 
            Height          =   1455
            Index           =   2
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   2566
         End
      End
      Begin Umakichi.ctlPane paneTab 
         Height          =   1695
         Index           =   3
         Left            =   -74760
         TabIndex        =   14
         Top             =   600
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   2990
         Begin Umakichi.ctlWrappedGrid flexTab 
            Height          =   1455
            Index           =   3
            Left            =   0
            TabIndex        =   15
            Top             =   0
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   2566
         End
      End
      Begin Umakichi.ctlPane paneTab 
         Height          =   1695
         Index           =   4
         Left            =   -74760
         TabIndex        =   16
         Top             =   480
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   2990
         Begin Umakichi.ctlWrappedGrid flexTab 
            Height          =   1455
            Index           =   4
            Left            =   0
            TabIndex        =   17
            Top             =   0
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   2566
         End
      End
   End
   Begin VB.Label lblMakeDate 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0EEEE&
      Caption         =   "データ作成日: 9999年99月99日"
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   6720
      TabIndex        =   3
      Top             =   60
      Width           =   2175
   End
End
Attribute VB_Name = "ctlVCH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   調教師マスタ 表示コントロール
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数(イベント)
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Public Event ChangeTo(strViewerName As String, strKey As String)        '' Vierer変更イベント
Public Event WindowTitle(strKey As String)                              '' ウインドウタイトル変更イベント
Public Event LinkContextMenu(strViewerName As String, strKey As String) '' 右クリックメニュー表示イベント

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mVB As clsViewerBase            '' Viewer Base
Private mViewerState As clsVSTabOnly    '' Viewer State

Private WithEvents mData As clsDataCH   '' データ取得オブジェクト
Attribute mData.VB_VarHelpID = -1
Private mstrTitle As String             '' ウインドウタイトル
Private mKey As clsKeyCH                '' キー
Private mblnNoData As Boolean           '' データ無しフラグ

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部定数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

' 画面最小幅値
Const MINIMUMWIDTH  As Long = 4000
Const MINIMUMHEIGHT As Long = 4000


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   プロパティ
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8


'
'   機能: キー設定プロパティ
'
'   備考: Viewer必須プロパティ
'
Public Property Let Key(strKey As String)
    gApp.Log "CH: " & strKey
    mKey.str = strKey
    Call Update
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
Public Property Get ViewerState() As clsVSTabOnly
    Set ViewerState = mViewerState
End Property


'
'   機能: Viewer状態受け取り
'
'   備考: なし
'
Public Property Set ViewerState(RHS As clsVSTabOnly)
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
    
    If 3 = Index Then
        Call flexTab(Index).ReflexiveMouseMoveDriven(True)
    ElseIf 4 = Index Then
        Call flexTab(Index).ReflexiveMouseMoveDriven(True)
    Else
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
    Dim i As Integer
    With paneTab.item(mstTab.Tab)
        .Top = mstTab.TabHeight + 60
        .Left = 60
        .width = Bigger(1, mstTab.width - .Left * 2)
        .Height = Bigger(1, (mstTab.Height - .Top) - (.Top - mstTab.TabHeight))
    End With ' paneTab.Item(mstTab.Index)
    For i = flexTab.LBound To flexTab.UBound
        With flexTab(i)
            .width = Bigger(1, paneTab(i).width - .Left)
            .Height = Bigger(1, paneTab(i).Height - .Top)
        End With ' flexTab(i)
    Next i
    
    mViewerState.LastTabNumber = mstTab.Tab
    
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
    Set mKey = New clsKeyCH
    Set mData = New clsDataCH
    Set mVB = New clsViewerBase
    Set mViewerState = New clsVSTabOnly
    
    mstrTitle = "調教師"

    ' 最小幅設定
    With UserControl
        .width = Bigger(.width, MINIMUMWIDTH)
        .Height = Bigger(.Height, MINIMUMHEIGHT)
    End With
    
    ' 共通UI設定
    Call mVB.InitGUI(UserControl.width, UserControl.Height, fraTop, lblMakeDate, fraHeader, mstTab)
    
    ' FlexGrid設定
    For i = flexTab.LBound To flexTab.UBound
        Call mVB.FlexGridCommonSetting(flexTab(i).Grid)
    Next i
        
    ' Color Assign
    BackColor = gApp.ColBG
    
    For i = lblInfo.LBound To lblInfo.UBound
        lblInfo(i).BackColor = gApp.ColDarkBG
        lblInfo(i).ForeColor = Contrast(gApp.ColDarkBG)
    Next i
    
    For i = txtInfo.LBound To lblInfo.UBound
        txtInfo(i).BackColor = gApp.ColBG
        txtInfo(i).ForeColor = Contrast(gApp.ColBG)
    Next i
    
    ' ctlPane の初期化
    For i = 3 To 4
        paneTab(i).Mode = 0 ' データ取得中に
        mstTab.TabEnabled(i) = False
    Next i
    
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
    Dim i As Integer
    
    ' 最小幅設定
    With UserControl
        .width = Bigger(.width, MINIMUMWIDTH)
        .Height = Bigger(.Height, MINIMUMHEIGHT)
    End With
    
    ' 共通UIリサイズ
    Call mVB.ResizeGUI(UserControl.width, UserControl.Height, fraTop, lblMakeDate, fraHeader, mstTab)
    
    With paneTab.item(mstTab.Tab)
        .width = Bigger(1, mstTab.width - .Left * 2)
        .Height = Bigger(1, (mstTab.Height - .Top) - (.Top - mstTab.TabHeight))
    End With ' paneTab.Item(mstTab.Index)
    
    For i = flexTab.LBound To flexTab.UBound
        With flexTab(i)
            .width = Bigger(1, paneTab(i).width - .Left)
            .Height = Bigger(1, paneTab(i).Height - .Top)
        End With ' flexTab(i)
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
    
    Set mKey = Nothing
    Set mData = Nothing
    Set mVB = Nothing
    
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: データが無い
'
'   備考: なし
'
Private Sub mData_NoData()
On Error GoTo ErrorHandler
    gApp.Log "d:該当レコードはありません。CHの存在するkeyを指定してください。" & vbCrLf _
            & "呼び出し元をチェックしましょう＞＞開発者"
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub

'
'   機能: 競走馬タブのデータがない
'
'   備考: なし
'
Private Sub mData_NoDataUMA()
On Error GoTo ErrorHandler
    paneTab(4).Mode = 1
    mstTab.TabEnabled(4) = True
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: 主要勝鞍タブのデータがそろった
'
'   備考: なし
'
Private Sub mData_FetchCompleteSYUYOU(gd As clsGridData)
On Error GoTo ErrorHandler
    Call flexTab(3).InsertGrid(gd)
    With flexTab(3).Grid
        .FixedCols = 0
        Call flexTab(3).AutoSize(0, .Cols - 1)
        Dim r As Long
        For r = 1 To .Rows - 1
            .row = r
            .col = 2
            .CellAlignment = flexAlignRightCenter
        Next r
    End With
    paneTab(3).Mode = 2
    mstTab.TabEnabled(3) = True
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
    Dim i As Long
    Dim sumWidth As Long
    
    Call flexTab(4).InsertGrid(gd)
    
    With flexTab(4).Grid
        .FixedCols = 0
        ' カラム幅最大値制限付のオートサイズ  長すぎる馬主名などへの対応の為
        Call flexTab(4).AutoSize(0, .Cols - 1)
    End With
    
    paneTab(4).Mode = 2
    mstTab.TabEnabled(4) = True
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: データを取得する
'
'   備考: なし
'
Private Sub Update()
    Dim i As Integer
    Dim j As Integer
    
    ' データを取得してもらう
    If Not mData.Fetch(mKey) Then
        mblnNoData = True
        Exit Sub
    End If
    
    ' ラベルをもらう
    lblMakeDate = mData.Labels(0)
    lblInfo(0) = ReplaceAmpersand(mData.Labels(1))
    lblInfo(1) = ReplaceAmpersand(mData.Labels(2))
    txtInfo(0) = mData.Labels(3)
    txtInfo(1) = mData.Labels(4)
        
    '履歴用文字列追加
    mstrTitle = mstrTitle & " " & mData.Labels(1)
    
    ' グリッドを反映する
    For i = 0 To 3
        Call flexTab(i).InsertGrid(mData.GridDatas(i))
    Next i
    
    'マージ、寄せ、幅の調整
    For i = 0 To 2
        With flexTab(i).Grid
            Call flexTab(i).AutoSize(0, .Cols - 1)
        End With
    Next i
    
    Dim r As Long, c As Long
    For i = 0 To 2
        With flexTab(i).Grid
            For r = 0 To 1
                For c = 0 To .Cols - 1
                    .row = r
                    .col = c
                    .CellAlignment = flexAlignCenterCenter
                Next c
            Next r
            For r = 2 To .Rows - 1
                For c = 0 To .Cols - 1
                    .row = r
                    .col = c
                    .CellAlignment = flexAlignRightCenter
                Next c
            Next r
        End With
    Next i
    With flexTab(1).Grid
        Call flexTab(1).AutoSize(1, .Cols - 1, False, False, 1)
        .FixedRows = 2
        .MergeRow(0) = True
        .MergeRow(1) = True
        .MergeCol(0) = True
        .MergeCells = flexMergeFree
    End With
    With flexTab(2).Grid
        Call flexTab(2).AutoSize(1, .Cols - 1, False, False, 2)
        .FixedRows = 3
        .MergeRow(0) = True
        .MergeRow(1) = True
        .MergeCol(0) = True
        .MergeCells = flexMergeFree
    End With
    
    ' 最初に表示するタブを設定する
    If mViewerState.IsNoTouch Then
        mstTab.Tab = 0
    Else
        mstTab.Tab = mViewerState.LastTabNumber
    End If
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

