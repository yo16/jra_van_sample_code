VERSION 5.00
Begin VB.UserControl ctlTitleBand 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.ComboBox cboRace 
      CausesValidation=   0   'False
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
      Left            =   1080
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   2
      Top             =   600
      Width           =   2520
   End
   Begin VB.ComboBox cboLocation 
      CausesValidation=   0   'False
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
      Left            =   1095
      Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
      TabIndex        =   1
      Top             =   120
      Width           =   1425
   End
   Begin VB.Label lblWrapped 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
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
      Left            =   60
      TabIndex        =   0
      Top             =   30
      Width           =   750
   End
End
Attribute VB_Name = "ctlTitleBand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'   タイトルバンド
'

Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数(イベント)
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

' RaceChanger が ユーザーによって変更された時に発生するイベント
Public Event Change(Key As clsKeyRA)

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

Private WithEvents mData As clsDataRaceChanger
Attribute mData.VB_VarHelpID = -1
Private mKey As clsKeyRASel
Private mblnTKMode As Boolean

Private mlngPreLocationIndex As Long    'cboLocationの直前のListIndex
Private mlngPreRaceIndex As Long

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   プロパティ
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: Captionプロパティの取得
'
'   備考: なし
'
Public Property Get Caption() As String
    Caption = lblWrapped.Caption
End Property

'
'   機能: Captionプロパティのセット
'
'   備考: なし
'
Public Property Let Caption(ByVal RHS As String)
    lblWrapped.Visible = False
    lblWrapped.Caption = ReplaceAmpersand(RHS)
End Property


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: RaceChangerの表示リクエスト
'
'   備考: 引き数 Key - レース選択画面用のキークラス
'                blnTKMode - 特別登録場モード、特別登録場画面の時はTrueで呼ばれる
'
'         Key が Nothingの時は非表示にする
'
Public Sub ShowRaceChanger(Key As clsKeyRASel, Optional blnTKMode As Boolean = False)
On Error GoTo ErrorHandler
    Dim i As Long   '' LocationIndex
    Dim j As Long   '' RaceIndex
    Dim lngLocationIndex As Long    '' mdata.RaceKeyのインデックス
    Dim lngRaceIndex As Long        '' mdata.RaceKeyのインデックス
    
    
    If Key Is Nothing Then
        ' 出馬表以外の画面では、Nothingで呼ばれる。
        cboLocation.Visible = False '非表示
        cboRace.Visible = False
    Else
        cboLocation.Visible = False  '表示
        cboRace.Visible = False
        cboLocation.Enabled = False 'Disable    ListIndexプロパティに代入するとClickイベントが発生するので
        cboRace.Enabled = False     '           イベント側ではEnabledプロパティがFalseのときはイベント処理しない
        If (Key.JyoCD > "10") Then      'JRA以外はDisableのまま
            ' 地方、海外レースでは、コンボボックスは表示するが空欄で選択不可とする。
            cboLocation.Clear
            cboRace.Clear
            Set mKey = New clsKeyRASel
        Else                            'JRAはcboを再セット・選択してEnableにする
        
            ' 場選択コンボボックスを表示するとき
            ' もともとタイトルバンドに表示していた回場日と重複する為
            ' 出馬表タイトル文字列の曜日の ")" 以降を削除する
            ' 　※出馬表タイトル文字列はウインドウタイトル、履歴にも用いている為
            ' 　　タイトルバンドでのみの例外処理として、ここで削除している。
            Me.Caption = Left(Me.Caption, InStr(Me.Caption, ")"))

            If (Left(Key.str, 8) <> Left(mKey.str, 8)) Or mblnTKMode <> blnTKMode Then '開催日が変わったら再取得
                Set mData = New clsDataRaceChanger
                Call mData.Fetch(Key, blnTKMode)
            Else                                            '開催日が同じなら
                Call GetRaceKeyIndex(Key, lngLocationIndex, lngRaceIndex)
                If Key.JyoCD <> mKey.JyoCD Then             '開催場が変わったら
                    cboLocation.ListIndex = lngLocationIndex    'cboLocation選択
                    Call SetCboRace(lngLocationIndex)           'cboRaceを入れ替え
                End If
                cboRace.ListIndex = lngRaceIndex    'cboRace選択
            End If
            Set mKey = Key  'Keyを保存
            mblnTKMode = blnTKMode
            cboLocation.Enabled = True  'Enable
            cboRace.Enabled = True
        End If
        cboLocation.Visible = True  '表示
        cboRace.Visible = True
    End If

    Call resize
    lblWrapped.Visible = True
    Exit Sub

ErrorHandler:
    gApp.ErrLog
End Sub

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   内部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: Keyと一致するRaceChangerクラスのKey配列プロパティのインデックスを取得する
'
'   備考: なし
'
Private Sub GetRaceKeyIndex(Key As clsKeyRASel, lngLocationIndex, lngRaceIndex)
    Dim i As Long
    Dim j As Long
    
    For i = 0 To mData.LocationCount - 1
        For j = 0 To mData.RaceCount(i) - 1
            If mData.RaceKey(i, j) = Key.str Then
                lngLocationIndex = i
                lngRaceIndex = j
                Exit Sub
            End If
        Next j
    Next i

End Sub


'
'   機能: 場コンボにRaceChangerクラスからアイテムをセットする
'
'   備考: なし
'
Private Sub SetCboLocation()
    Dim i As Long
    
    cboLocation.Clear
    For i = 0 To mData.LocationCount - 1
        cboLocation.AddItem mData.LocationName(i)
    Next

End Sub


'
'   機能: レースコンボにRaceChangerクラスからアイテムをセットする
'
'   備考: なし
'
Private Sub SetCboRace(lngLocationIndex)
    Dim i As Long
    
    cboRace.Clear
    For i = 0 To mData.RaceCount(lngLocationIndex) - 1
        cboRace.AddItem mData.RaceName(lngLocationIndex, i)
    Next
    
End Sub

'
'   機能: コンボアイテム選択イベント
'
'   備考: レースコンボのアイテムを選択した場に入れ替え
'
Private Sub cboLocation_Click()
On Error GoTo ErrorHandler
    Dim i As Long
    Dim lngNewLocationIndex As Long
    Dim strPreRaceNum As String                     '現在のレース番号
    Dim lngNewRaceIndex As Long                     'cboRaceにセットするインデックス
    Dim blnFlag As Boolean
    Dim Key As clsKeyRA
    Dim lngSmallerRaceNumIndex As Long
    Dim lngLargerRaceNumIndex As Long
    
    If cboLocation.Enabled = False Then             '履歴などで操作されたとき
        mlngPreLocationIndex = cboLocation.ListIndex
        Exit Sub
    Else                                            'コンボボックスを操作されたとき
        
        lngNewLocationIndex = cboLocation.ListIndex '選択した場
        
        If (lngNewLocationIndex <> mlngPreLocationIndex) Then   '場を変更した
            
            strPreRaceNum = Left(cboRace.Text, 2)       '選択していたレース番号
            
            Call SetCboRace(lngNewLocationIndex)        'レースコンボを入れ替え
        
            '場を変更したとき同じレース番号がない場合、最も近いレース番号を選択する
            lngSmallerRaceNumIndex = -1
            lngLargerRaceNumIndex = -1
            blnFlag = False
            For i = 0 To mData.RaceCount(lngNewLocationIndex) - 1
                If (strPreRaceNum = Left(cboRace.List(i), 2)) Then  '同じレース番号があった
                    lngNewRaceIndex = i     '同じレース番号
                    blnFlag = True
                    Exit For
                ElseIf (strPreRaceNum > Left(cboRace.List(i), 2)) Then  '前のレース
                    lngSmallerRaceNumIndex = i                          '
                ElseIf (strPreRaceNum < Left(cboRace.List(i), 2)) Then  '後のレース
                    lngLargerRaceNumIndex = i                           '
                    Exit For                                            '
                End If
            Next
            If blnFlag = False Then         '同じレース番号がないとき
                If (lngSmallerRaceNumIndex = -1) Then   '前のレースが無い
                    If (lngLargerRaceNumIndex = -1) Then    '後のレースが無い
                        lngNewRaceIndex = 0                     'エラー、最初のレースをセットする
                    Else                                    '後のレースがある
                        lngNewRaceIndex = lngLargerRaceNumIndex '後のレース番号
                    End If
                Else                                    '前のレースがある
                    If (lngLargerRaceNumIndex = -1) Then    '後のレースが無い
                        lngNewRaceIndex = lngSmallerRaceNumIndex    '前のレース番号
                    Else                                    '後ろのレースがある
                        If strPreRaceNum - Left(cboRace.List(lngSmallerRaceNumIndex), 2) >= Left(cboRace.List(lngLargerRaceNumIndex), 2) - strPreRaceNum Then   'どちらが近いか
                            lngNewRaceIndex = lngLargerRaceNumIndex     '同じが前が遠いときは後のレース番号
                        Else
                            lngNewRaceIndex = lngSmallerRaceNumIndex    '前が近いときは前のレース番号
                        End If
                    End If
                End If
            End If
            cboRace.ListIndex = lngNewRaceIndex     'cboRace_Click発生
        End If
    End If
    Exit Sub

ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: コンボアイテム選択イベント
'
'   備考: 選択したレースで画面を更新
'
Private Sub cboRace_Click()
On Error GoTo ErrorHandler
    Dim Key As clsKeyRA
    
    If cboRace.Enabled = False Then             '履歴などで操作されたとき
        mlngPreRaceIndex = cboRace.ListIndex
    Else                                        'コンボボックスを操作されたとき
        If (mlngPreLocationIndex <> cboLocation.ListIndex) Or (mlngPreRaceIndex <> cboRace.ListIndex) Then
            Set Key = New clsKeyRA
            Key.str = mData.RaceKey(cboLocation.ListIndex, cboRace.ListIndex)
    
            RaiseEvent Change(Key)              '出馬表を表示
            
            mlngPreLocationIndex = cboLocation.ListIndex    'コンボのインデックスを保存
            mlngPreRaceIndex = cboRace.ListIndex
        End If
    End If
    Exit Sub
    
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: データ取得終了イベント、コンボに取得した値をセットする
'
'   備考: なし
'
'
Private Sub mData_FetchComplete(locationIndex As Long, raceIndex As Long)
On Error GoTo ErrorHandler
    
    cboLocation.Enabled = False     'Disable
    cboRace.Enabled = False
    
    Call SetCboLocation             'アイテムをセット
    Call SetCboRace(locationIndex)
    
    cboLocation.ListIndex = locationIndex   'アイテムを選択
    cboRace.ListIndex = raceIndex
    
    cboLocation.Enabled = True      'Enable
    cboRace.Enabled = True
    Exit Sub

ErrorHandler:
    gApp.ErrLog
End Sub

'
'   機能: ユーザコントロールの初期化
'
'   備考: なし
'
'
Private Sub UserControl_Initialize()

    Set mKey = New clsKeyRASel
    Set mData = New clsDataRaceChanger
    
End Sub


'
'   機能: ユーザコントロールのプロパティ取得
'
'   備考: なし
'
'
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
On Error GoTo ErrorHandler
    lblWrapped.Caption = ReplaceAmpersand(PropBag.ReadProperty("Caption", "Label1"))
    Call resize
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub

'
'   機能: ユーザコントロールのリサイズイベント
'
'   備考: なし
'
'
Private Sub UserControl_Resize()
On Error GoTo ErrorHandler
    Call resize
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: ユーザコントロールのプロパティセット
'
'   備考: なし
'
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
On Error GoTo ErrorHandler
    Call PropBag.WriteProperty("Caption", lblWrapped.Caption, "Label1")
    Exit Sub
ErrorHandler:
    gApp.ErrLog
End Sub


'
'   機能: コントロールのレイアウト
'
'   備考: キャプションが設定され、ラベルがAutoSizeで変更されたとき
'
Private Sub resize()
    Dim p As Single
    
    p = lblWrapped.Left + lblWrapped.width
    
    'ラベルとコンボの距離
    p = p + 200
    
    ' 場コンボの再配置
    cboLocation.Move p, 0
    p = p + cboLocation.width
    
    '場コンボとRaceコンボの距離
    p = p + 100
    
    ' Raceコンボの再配置
    cboRace.Move p, 0
    p = p + cboRace.width
    
    UserControl.width = p
    UserControl.Height = lblWrapped.Height + lblWrapped.Top * 2
End Sub
