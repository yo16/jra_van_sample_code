VERSION 5.00
Begin VB.UserControl ctlClickLabel 
   BackColor       =   &H000000FF&
   ClientHeight    =   3570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4980
   ScaleHeight     =   3570
   ScaleWidth      =   4980
   Begin VB.Timer tmrMouse 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   3720
      Top             =   30
   End
   Begin VB.Label lblWraped 
      BackColor       =   &H00E0EEEE&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2925
   End
End
Attribute VB_Name = "ctlClickLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   クリック可能なリンクラベルユーザーコントロール
'
'   ラベル１つをラップし、右クリックによるポップアップメニューを持ち
'   通常リンクオープン(ChangeViewer)、あるいは、新規ウインドウで開く(NewWindow)の
'   二種類のイベントを生成する。
'
Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private mlngForeColor As Long
Private mlngBackColor As Long

'プロパティ変数:
Dim m_Key As String
Dim m_ViewerName As String

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部定数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'既定のプロパティ値:
Const m_def_Key = "0"
Const m_def_ViewerName = ""

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数(イベント)
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

' イベント
Event ChangeViewer()
Event RightMouseDown()

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   プロパティ
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: AutoSizeプロパティを取得
'
'   備考: MappingInfo=lblWraped,lblWraped,-1,AutoSize
'
Public Property Get AutoSize() As Boolean
Attribute AutoSize.VB_Description = "ｺﾝﾄﾛｰﾙのｻｲｽﾞが、その内容にあわせて自動的に変更されるかどうかを設定します。値の取得も可能です。"
    AutoSize = lblWraped.AutoSize
    UserControl.width = lblWraped.width
    UserControl.Height = lblWraped.Height
End Property

'
'   機能: AutoSizeプロパティをセット
'
'   備考: MappingInfo=lblWraped,lblWraped,-1,AutoSize
'
Public Property Let AutoSize(ByVal New_AutoSize As Boolean)
    lblWraped.AutoSize() = New_AutoSize
    If lblWraped.AutoSize Then
        UserControl.width = lblWraped.width
        UserControl.Height = lblWraped.Height
    End If
    PropertyChanged "Height"
    PropertyChanged "Width"
    PropertyChanged "AutoSize"
End Property

'
'   機能: BackColorプロパティを取得
'
'   備考: MappingInfo=lblWraped,lblWraped,-1,BackColor
'
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "ｵﾌﾞｼﾞｪｸﾄ内の文字やｸﾞﾗﾌｨｯｸの表示で使用する背景色を設定します。値の取得も可能です。"
    BackColor = lblWraped.BackColor
End Property

'
'   機能: BackColorプロパティをセット
'
'   備考: MappingInfo=lblWraped,lblWraped,-1,BackColor
'
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    lblWraped.BackColor() = New_BackColor
    mlngBackColor = lblWraped.BackColor
    PropertyChanged "BackColor"
End Property

'
'   機能: Enabledプロパティを取得
'
'   備考: MappingInfo=lblWraped,lblWraped,-1,Enabled
'
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "ﾕｰｻﾞｰの操作で発生したｲﾍﾞﾝﾄを、ｵﾌﾞｼﾞｪｸﾄに認識させるかどうかを設定します。値の取得も可能です。"
    Enabled = lblWraped.Enabled
End Property

'
'   機能: Enabledプロパティをセット
'
'   備考: MappingInfo=lblWraped,lblWraped,-1,Enabled
'
Public Property Let Enabled(ByVal New_Enabled As Boolean)
    lblWraped.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'
'   機能: Fontプロパティを取得
'
'   備考: MappingInfo=lblWraped,lblWraped,-1,Font
'
Public Property Get Font() As Font
Attribute Font.VB_Description = "Font ｵﾌﾞｼﾞｪｸﾄを返します。"
Attribute Font.VB_UserMemId = -512
    Set Font = lblWraped.Font
End Property

'
'   機能: Fontプロパティをセット
'
'   備考: MappingInfo=lblWraped,lblWraped,-1,Font
'
Public Property Set Font(ByVal New_Font As Font)
    Set lblWraped.Font = New_Font
    PropertyChanged "Font"
End Property

'
'   機能: ForeColorプロパティを取得
'
'   備考: MappingInfo=lblWraped,lblWraped,-1,ForeColor
'
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "ｵﾌﾞｼﾞｪｸﾄ内の文字やｸﾞﾗﾌｨｯｸの表示で使用する前景色を設定します。値の取得も可能です。"
    ForeColor = lblWraped.ForeColor
End Property

'
'   機能: ForeColorプロパティをセット
'
'   備考: MappingInfo=lblWraped,lblWraped,-1,ForeColor
'
Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    lblWraped.ForeColor() = New_ForeColor
    mlngForeColor = lblWraped.ForeColor
    PropertyChanged "ForeColor"
End Property

'
'   機能: WordWrapプロパティを取得
'
'   備考: MappingInfo=lblWraped,lblWraped,-1,WordWrap
'
Public Property Get WordWrap() As Boolean
Attribute WordWrap.VB_Description = "ｷｬﾌﾟｼｮﾝの文字列の長さに応じて、ｺﾝﾄﾛｰﾙが広がるかどうかを示す値を設定します。値の取得も可能です。"
    WordWrap = lblWraped.WordWrap
End Property

'
'   機能: WordWrapプロパティをセット
'
'   備考: MappingInfo=lblWraped,lblWraped,-1,WordWrap
'
Public Property Let WordWrap(ByVal New_WordWrap As Boolean)
    lblWraped.WordWrap() = New_WordWrap
    PropertyChanged "WordWrap"
End Property

'
'   機能: Keyプロパティを取得
'
'   備考: MemberInfo=13,0,0,0
'
Public Property Get Key() As String
    Key = m_Key
End Property

'
'   機能: Keyプロパティをセット
'
'   備考: MemberInfo=13,0,0,0
'
Public Property Let Key(ByVal New_Key As String)
    m_Key = New_Key
    PropertyChanged "Key"
    If m_Key <> "" Then
        lblWraped.MousePointer = 99
        lblWraped.MouseIcon = LoadResPicture(101, vbResCursor)
    End If
End Property

'
'   機能: ViewerNameプロパティを取得
'
'   備考: MemberInfo=13,0,0,
'
Public Property Get ViewerName() As String
    ViewerName = m_ViewerName
End Property

'
'   機能: ViewerNameプロパティをセット
'
'   備考: MemberInfo=13,0,0,
'
Public Property Let ViewerName(ByVal New_ViewerName As String)
    m_ViewerName = New_ViewerName
    PropertyChanged "ViewerName"
End Property

'
'   機能: Captionプロパティを取得
'
'   備考: MappingInfo=lblWraped,lblWraped,-1,Caption
'
Public Property Get Caption() As String
    Caption = lblWraped.Caption
End Property

'
'   機能: Captionプロパティをセット
'
'   備考: MappingInfo=lblWraped,lblWraped,-1,Caption
'
Public Property Let Caption(ByVal New_Caption As String)
    lblWraped.Caption() = New_Caption
    If lblWraped.AutoSize Then
        UserControl.width = lblWraped.width
        UserControl.Height = lblWraped.Height
    End If
    PropertyChanged "Height"
    PropertyChanged "Width"
    PropertyChanged "Caption"
End Property

'
'   機能: グリッドアイテムから、リンク情報を設定する
'
'   備考: なし
'
Public Property Set LinkItem(ByVal New_LinkItem As clsGridItem)
    m_Key = New_LinkItem.Key
    m_ViewerName = New_LinkItem.Link
    lblWraped.Caption() = ReplaceAmpersand(New_LinkItem.Text)
    lblWraped.ToolTipText = New_LinkItem.ToolTip
    If lblWraped.AutoSize Then
        UserControl.width = lblWraped.width
        UserControl.Height = lblWraped.Height
    End If
    PropertyChanged "Height"
    PropertyChanged "Width"
    PropertyChanged "Key"
    PropertyChanged "ViewerName"
    PropertyChanged "Caption"
End Property


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: クリックイベント
'
'   備考: なし
'
Private Sub lblWraped_Click()
On Error GoTo ErrorHandler
    If m_Key <> "" And m_ViewerName <> "" Then
        lblWraped.MousePointer = vbHourglass
        RaiseEvent ChangeViewer
    End If
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub

'
'   機能: マウスダウンイベント
'
'   備考: なし
'
Private Sub lblWraped_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrorHandler
    If Button = vbRightButton Then
        If m_Key <> "" And m_ViewerName <> "" Then
            RaiseEvent RightMouseDown
        End If
    End If
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub

'
'   機能: マウスムーブイベント
'
'   備考: なし
'
Private Sub lblWraped_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo ErrorHandler
    If m_Key <> "" And m_ViewerName <> "" Then
        lblWraped.MousePointer = 99
        lblWraped.MouseIcon = LoadResPicture(101, vbResCursor)
        tmrMouse.Enabled = True
    Else
        lblWraped.MousePointer = vbDefault
    End If
    ' MouseOut チェック開始
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub

'
'   機能: マウスタイマーイベント
'
'   備考: なし
'
Private Sub tmrMouse_Timer()
On Error GoTo ErrorHandler
    Dim WX1 As Long
    Dim WY1 As Long
    Dim WX2 As Long
    Dim WY2 As Long
    
    Dim MX As Long
    Dim MY As Long
    
    ' ユーザーコントロール位置を取得
    Call GetWindowRect(UserControl.hwnd, WX1, WY1, WX2, WY2)
    
    ' マウスカーソル位置を取得
    Call GetCursorPos(MX, MY)
    
    ' カーソルがユーザーコントロールの内側か外側かを判定
    If MX >= WX1 And MX <= WX2 And _
        MY >= WY1 And MY <= WY2 Then
        ' 内側
        Call SetHot
    Else
        ' 外側
        Call SetCool
        tmrMouse.Enabled = False
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
Private Sub UserControl_Initialize()
On Error GoTo ErrorHandler
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub

'
'   機能: プロパティの初期化
'
'   備考: なし
'
Private Sub UserControl_InitProperties()
On Error GoTo ErrorHandler
    m_Key = m_def_Key
    m_ViewerName = m_def_ViewerName
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: プロパティの読み込み
'
'   備考: なし
'
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error GoTo ErrorHandler

    lblWraped.AutoSize = PropBag.ReadProperty("AutoSize", False)
    lblWraped.BackColor = PropBag.ReadProperty("BackColor", &HE0EEEE)
    lblWraped.Enabled = PropBag.ReadProperty("Enabled", True)
    Set lblWraped.Font = PropBag.ReadProperty("Font", Ambient.Font)
    lblWraped.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    lblWraped.WordWrap = PropBag.ReadProperty("WordWrap", False)
    m_Key = PropBag.ReadProperty("Key", m_def_Key)
    m_ViewerName = PropBag.ReadProperty("ViewerName", m_def_ViewerName)
    lblWraped.Caption = PropBag.ReadProperty("Caption", "Label1")
    
    mlngForeColor = lblWraped.ForeColor
    mlngBackColor = lblWraped.BackColor
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: リサイズイベント
'
'   備考: なし
'
Private Sub UserControl_Resize()
On Error GoTo ErrorHandler
    lblWraped.width = ScaleWidth
    lblWraped.Height = ScaleHeight
    lblWraped.AutoSize = lblWraped.AutoSize
    UserControl.width = lblWraped.width
    UserControl.Height = lblWraped.Height
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: プロパティの書き込み
'
'   備考: なし
'
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error GoTo ErrorHandler

    Call PropBag.WriteProperty("AutoSize", lblWraped.AutoSize, False)
    Call PropBag.WriteProperty("BackColor", lblWraped.BackColor, &HE0EEEE)
    Call PropBag.WriteProperty("Enabled", lblWraped.Enabled, True)
    Call PropBag.WriteProperty("Font", lblWraped.Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", lblWraped.ForeColor, &H80000012)
    Call PropBag.WriteProperty("WordWrap", lblWraped.WordWrap, False)
    Call PropBag.WriteProperty("Key", m_Key, m_def_Key)
    Call PropBag.WriteProperty("ViewerName", m_ViewerName, m_def_ViewerName)
    Call PropBag.WriteProperty("Caption", lblWraped.Caption, "Label1")
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: フォアカラーにプロパティ値をセット
'
'   備考: なし
'
Private Sub SetCool()
    lblWraped.ForeColor = mlngForeColor
End Sub


'
'   機能: フォアカラーをColorLinkedに設定
'
'   備考: なし
'
Private Sub SetHot()
    lblWraped.ForeColor = ColorLinked
End Sub
