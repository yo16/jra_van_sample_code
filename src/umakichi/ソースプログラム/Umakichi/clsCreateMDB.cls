VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCreateMDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "MDB作成、テーブル定義をするクラス"
'
'   MDB作成、テーブル定義をするクラス
'
Option Explicit

'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部変数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

' ADOコネクションをプロパティにセットしてからメソッドをよぶ
Public mConnection As ADODB.Connection


'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8
'   外部関数
'---+----1----+----2----+----3----+----4----+----5----+----6----+----7----+----8

'
'   機能: テーブルで切り分ける
'
'   備考: なし
'
Public Function createMDB() As Boolean
On Error GoTo Errorhandler
    Dim fso     As New Scripting.FileSystemObject   '' ファイルシステムオブジェクト
    Dim ts      As Scripting.TextStream
    Dim strTbl  As String                           '' 全テーブルの定義用
    Dim db()    As String                           '' 子MDB名配列
    Dim i       As Integer                          '' ループカウンタ
    Dim j       As Integer                          '' ループカウンタ
    Dim cat     As New ADOX.Catalog                 '' 子MDB用カタログオブジェクト
    Dim cat2    As New ADOX.Catalog                 '' 親MDB用カタログオブジェクト
    Dim tbl     As ADOX.Table                       '' リンクテーブル作成用テーブルオブジェクト

    ' 全テーブル定義
    strTbl = strTbl & "TOKU_RACE"
    strTbl = strTbl & ",TOKU"
    strTbl = strTbl & ",RACE"
    strTbl = strTbl & ",UMA_RACE_A"
    strTbl = strTbl & ",UMA_RACE_B"
    strTbl = strTbl & ",HARAI"
    strTbl = strTbl & ",ODDS_TANPUKUWAKU"
    strTbl = strTbl & ",ODDS_UMAREN"
    strTbl = strTbl & ",ODDS_WIDE"
    strTbl = strTbl & ",ODDS_UMATAN0"
    strTbl = strTbl & ",ODDS_UMATAN1"
    strTbl = strTbl & ",ODDS_UMATAN2"
    strTbl = strTbl & ",ODDS_UMATAN3"
    strTbl = strTbl & ",ODDS_UMATAN4"
    strTbl = strTbl & ",ODDS_UMATAN5"
    strTbl = strTbl & ",ODDS_UMATAN6"
    strTbl = strTbl & ",ODDS_UMATAN7"
    strTbl = strTbl & ",ODDS_UMATAN8"
    strTbl = strTbl & ",ODDS_UMATAN9"
    strTbl = strTbl & ",ODDS_SANREN0"
    strTbl = strTbl & ",ODDS_SANREN1"
    strTbl = strTbl & ",ODDS_SANREN2"
    strTbl = strTbl & ",ODDS_SANREN3"
    strTbl = strTbl & ",ODDS_SANREN4"
    strTbl = strTbl & ",ODDS_SANREN5"
    strTbl = strTbl & ",ODDS_SANREN6"
    strTbl = strTbl & ",ODDS_SANREN7"
    strTbl = strTbl & ",ODDS_SANREN8"
    strTbl = strTbl & ",ODDS_SANREN9"
    strTbl = strTbl & ",UMA"
    strTbl = strTbl & ",KISHU"
    strTbl = strTbl & ",KISHU_SEISEKI"
    strTbl = strTbl & ",CHOKYO"
    strTbl = strTbl & ",CHOKYO_SEISEKI"
    strTbl = strTbl & ",SEISAN"
    strTbl = strTbl & ",BANUSI"
    strTbl = strTbl & ",HANSYOKU"
    strTbl = strTbl & ",SANKU"
    strTbl = strTbl & ",RECORD"
    strTbl = strTbl & ",HANRO"
    strTbl = strTbl & ",BATAIJYU"
    strTbl = strTbl & ",TENKO_BABA"
    strTbl = strTbl & ",TORIKESI_JYOGAI"
    strTbl = strTbl & ",KISHU_CHANGE"
    strTbl = strTbl & ",MINING"
    strTbl = strTbl & ",SCHEDULE"
    strTbl = strTbl & ",RAKaiSel"
    strTbl = strTbl & ",HASSOU_CHANGE"
    strTbl = strTbl & ",COURSE_CHANGE"

    ' 子MDBはテーブル毎
    db = Split(strTbl, ",")
    
    ' 親MDB生成
    cat2.Create ("Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=" & gApp.R_DBPath & "\LinkTables.mdb;" & _
        "Jet OLEDB:Engine Type=5;") '' Engine Type=5 : Access2000
    

    ' 子MDBの生成,子MDBテーブル作成, 親MDBリンクテーブル作成
    For i = 0 To UBound(db)
        ' 子MDBの生成
        cat.Create ("Provider=Microsoft.Jet.OLEDB.4.0;" & _
            "Data Source=" & gApp.R_DBPath & "\sub" & db(i) & ".mdb;" & _
            "Jet OLEDB:Engine Type=5;") '' Engine Type=5 : Access2000

        ' 子MDBへのコネクション
        Set Me.mConnection = cat.ActiveConnection
        
        ' 子MDBへのテーブル作成
        ' 連番付の場合、連番を消したテーブル作成関数を呼び出す
        If Right$(db(i), 1) >= "0" And Right$(db(i), 1) <= "9" Then
            Call CallByName(Me, "createTable_" & Left$(db(i), Len(db(i)) - 1), VbMethod)
            Call CallByName(Me, "createIndex_" & Left$(db(i), Len(db(i)) - 1), VbMethod)
        Else
            Call CallByName(Me, "createTable_" & db(i), VbMethod)
            Call CallByName(Me, "createIndex_" & db(i), VbMethod)
        End If
        
        ' 親MDBにリンクテーブルを作成
        Set tbl = New ADOX.Table
        With tbl
            .Name = db(i)
            Set .ParentCatalog = cat2

            .Properties("Jet OLEDB:Create Link") = True
            .Properties("Jet OLEDB:Link Datasource") = gApp.R_DBPath & "\sub" & db(i) & ".mdb"
            '連番付の場合、連番を消したテーブル名を設定
            If Right$(db(i), 1) >= "0" And Right$(db(i), 1) <= "9" Then
                .Properties("Jet OLEDB:Remote Table Name") = Left$(db(i), Len(db(i)) - 1)
            Else
                .Properties("Jet OLEDB:Remote Table Name") = db(i)
            End If
        
        End With
        cat2.Tables.Append tbl
        
        Set tbl = Nothing

        Me.mConnection.Close
    Next i    ' 子MDBの生成,子MDBテーブル作成, 親MDBリンクテーブル作成
    
    SetIniData "DatabaseInfo", "Version", "1", gApp.R_DBPath & "\UmakichiDB.ini"
    
    createMDB = True ' 成功
    
    Exit Function
Errorhandler:
    gApp.ErrLog
    createMDB = False ' 失敗
End Function


'
'   機能: DBにTOKU_RACE(特別レース)テーブルを作る
'
'   備考: なし
'
Public Function createTable_TOKU_RACE() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String

    strSQL = "CREATE TABLE TOKU_RACE("
    strSQL = strSQL & "[RecordSpec]                    CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[DataKubun]                     CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[MakeDate]                      CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[Year]                          CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[MonthDay]                      CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[JyoCD]                         CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kaiji]                         CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Nichiji]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[RaceNum]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[YoubiCD]                       CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[TokuNum]                       CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[Hondai]                        CHAR(60)  WITH COMPRESSION,"
    strSQL = strSQL & "[Fukudai]                       CHAR(60)  WITH COMPRESSION,"
    strSQL = strSQL & "[Kakko]                         CHAR(60)  WITH COMPRESSION,"
    strSQL = strSQL & "[HondaiEng]                     CHAR(120) WITH COMPRESSION,"
    strSQL = strSQL & "[FukudaiEng]                    CHAR(120) WITH COMPRESSION,"
    strSQL = strSQL & "[KakkoEng]                      CHAR(120) WITH COMPRESSION,"
    strSQL = strSQL & "[Ryakusyo10]                    CHAR(20)  WITH COMPRESSION,"
    strSQL = strSQL & "[Ryakusyo6]                     CHAR(12)  WITH COMPRESSION,"
    strSQL = strSQL & "[Ryakusyo3]                     CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kubun]                         CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[Nkai]                          CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[GradeCD]                       CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[SyubetuCD]                     CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[KigoCD]                        CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[JyuryoCD]                      CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[JyokenCD1]                     CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[JyokenCD2]                     CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[JyokenCD3]                     CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[JyokenCD4]                     CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[JyokenCD5]                     CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori]                         CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[TrackCD]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[CourseKubunCD]                 CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[HandiDate]                     CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[TorokuTosu]                    CHAR(3)   WITH COMPRESSION"
    strSQL = strSQL & ")"

    mConnection.Execute strSQL

    createTable_TOKU_RACE = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createTable_TOKU_RACE = False
End Function ' createTable_TOKU_RACE


'
'   機能: DBにTOKU(特別登録馬)テーブルを作る
'
'   備考: なし
'
Public Function createTable_TOKU() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String

    strSQL = "CREATE TABLE TOKU("
    strSQL = strSQL & "[Year]                          CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[MonthDay]                      CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[JyoCD]                         CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kaiji]                         CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Nichiji]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[RaceNum]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Num]                           CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[KettoNum]                      CHAR(10)  WITH COMPRESSION,"
    strSQL = strSQL & "[Bamei]                         CHAR(36)  WITH COMPRESSION,"
    strSQL = strSQL & "[UmaKigoCD]                     CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[SexCD]                         CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[TozaiCD]                       CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[ChokyosiCode]                  CHAR(5)   WITH COMPRESSION,"
    strSQL = strSQL & "[ChokyosiRyakusyo]              CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[Futan]                         CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Koryu]                         CHAR(1)   WITH COMPRESSION"
    strSQL = strSQL & ")"

    mConnection.Execute strSQL

    createTable_TOKU = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createTable_TOKU = False
End Function ' createTable_TOKU


'
'   機能: DBにRACE(レース詳細)テーブルを作る
'
'   備考: なし
'
Public Function createTable_RACE() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String

    strSQL = "CREATE TABLE RACE("
    strSQL = strSQL & "[RecordSpec]                    CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[DataKubun]                     CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[MakeDate]                      CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[Year]                          CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[MonthDay]                      CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[JyoCD]                         CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kaiji]                         CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Nichiji]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[RaceNum]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[YoubiCD]                       CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[TokuNum]                       CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[Hondai]                        CHAR(60)  WITH COMPRESSION,"
    strSQL = strSQL & "[Fukudai]                       CHAR(60)  WITH COMPRESSION,"
    strSQL = strSQL & "[Kakko]                         CHAR(60)  WITH COMPRESSION,"
    strSQL = strSQL & "[HondaiEng]                     CHAR(120) WITH COMPRESSION,"
    strSQL = strSQL & "[FukudaiEng]                    CHAR(120) WITH COMPRESSION,"
    strSQL = strSQL & "[KakkoEng]                      CHAR(120) WITH COMPRESSION,"
    strSQL = strSQL & "[Ryakusyo10]                    CHAR(20)  WITH COMPRESSION,"
    strSQL = strSQL & "[Ryakusyo6]                     CHAR(12)  WITH COMPRESSION,"
    strSQL = strSQL & "[Ryakusyo3]                     CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kubun]                         CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[Nkai]                          CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[GradeCD]                       CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[GradeCDBefore]                 CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[SyubetuCD]                     CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[KigoCD]                        CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[JyuryoCD]                      CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[JyokenCD1]                     CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[JyokenCD2]                     CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[JyokenCD3]                     CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[JyokenCD4]                     CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[JyokenCD5]                     CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[JyokenName]                    CHAR(60)  WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori]                         CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[KyoriBefore]                   CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[TrackCD]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[TrackCDBefore]                 CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[CourseKubunCD]                 CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[CourseKubunCDBefore]           CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Honsyokin1]                    CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[Honsyokin2]                    CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[Honsyokin3]                    CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[Honsyokin4]                    CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[Honsyokin5]                    CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[Honsyokin6]                    CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[Honsyokin7]                    CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[HonsyokinBefore1]              CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[HonsyokinBefore2]              CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[HonsyokinBefore3]              CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[HonsyokinBefore4]              CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[HonsyokinBefore5]              CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[Fukasyokin1]                   CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[Fukasyokin2]                   CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[Fukasyokin3]                   CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[Fukasyokin4]                   CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[Fukasyokin5]                   CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[FukasyokinBefore1]             CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[FukasyokinBefore2]             CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[FukasyokinBefore3]             CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[HassoTime]                     CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[HassoTimeBefore]               CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[TorokuTosu]                    CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[SyussoTosu]                    CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[NyusenTosu]                    CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[TenkoCD]                       CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[SibaBabaCD]                    CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[DirtBabaCD]                    CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[LapTime1]                      CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[LapTime2]                      CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[LapTime3]                      CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[LapTime4]                      CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[LapTime5]                      CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[LapTime6]                      CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[LapTime7]                      CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[LapTime8]                      CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[LapTime9]                      CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[LapTime10]                     CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[LapTime11]                     CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[LapTime12]                     CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[LapTime13]                     CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[LapTime14]                     CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[LapTime15]                     CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[LapTime16]                     CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[LapTime17]                     CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[LapTime18]                     CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[LapTime19]                     CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[LapTime20]                     CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[LapTime21]                     CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[LapTime22]                     CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[LapTime23]                     CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[LapTime24]                     CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[LapTime25]                     CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[SyogaiMileTime]                CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[HaronTimeS3]                   CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[HaronTimeS4]                   CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[HaronTimeL3]                   CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[HaronTimeL4]                   CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Corner1]                       CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[Syukaisu1]                     CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyuni1]                        CHAR(70)  WITH COMPRESSION,"
    strSQL = strSQL & "[Corner2]                       CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[Syukaisu2]                     CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyuni2]                        CHAR(70)  WITH COMPRESSION,"
    strSQL = strSQL & "[Corner3]                       CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[Syukaisu3]                     CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyuni3]                        CHAR(70)  WITH COMPRESSION,"
    strSQL = strSQL & "[Corner4]                       CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[Syukaisu4]                     CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyuni4]                        CHAR(70)  WITH COMPRESSION,"
    strSQL = strSQL & "[RecordUpKubun]                 CHAR(1)   WITH COMPRESSION"
    strSQL = strSQL & ")"

    mConnection.Execute strSQL

    createTable_RACE = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createTable_RACE = False
End Function ' createTable_RACE


'
'   機能: DBにUMA_RACE_A(馬毎レース情報_前半)テーブルを作る
'
'   備考: なし
'
Public Function createTable_UMA_RACE_A() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String

    strSQL = "CREATE TABLE UMA_RACE_A("
    strSQL = strSQL & "[RecordSpec]                    CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[DataKubun]                     CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[MakeDate]                      CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[Year]                          CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[MonthDay]                      CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[JyoCD]                         CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kaiji]                         CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Nichiji]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[RaceNum]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Wakuban]                       CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[Umaban]                        CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[KettoNum]                      CHAR(10)  WITH COMPRESSION,"
    strSQL = strSQL & "[Bamei]                         CHAR(36)  WITH COMPRESSION,"
    strSQL = strSQL & "[UmaKigoCD]                     CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[SexCD]                         CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HinsyuCD]                      CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[KeiroCD]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Barei]                         CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[TozaiCD]                       CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[ChokyosiCode]                  CHAR(5)   WITH COMPRESSION,"
    strSQL = strSQL & "[ChokyosiRyakusyo]              CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[BanusiCode]                    CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[BanusiName]                    CHAR(64)  WITH COMPRESSION,"
    strSQL = strSQL & "[Fukusyoku]                     CHAR(60)  WITH COMPRESSION,"
    strSQL = strSQL & "[reserved1]                     CHAR(60)  WITH COMPRESSION"
    strSQL = strSQL & ")"

    mConnection.Execute strSQL

    createTable_UMA_RACE_A = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createTable_UMA_RACE_A = False
End Function ' createTable_UMA_RACE_A


'
'   機能: DBにUMA_RACE_B(馬毎レース情報_後半)テーブルを作る
'
'   備考: なし
'
Public Function createTable_UMA_RACE_B() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String

    strSQL = "CREATE TABLE UMA_RACE_B("
    strSQL = strSQL & "[DataKubun]                     CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[B_Year]                        CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[B_MonthDay]                    CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[B_JyoCD]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[B_RaceNum]                     CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[B_Umaban]                      CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[B_KettoNum]                    CHAR(10)  WITH COMPRESSION,"
    strSQL = strSQL & "[Futan]                         CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[FutanBefore]                   CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Blinker]                       CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[reserved2]                     CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[KisyuCode]                     CHAR(5)   WITH COMPRESSION,"
    strSQL = strSQL & "[KisyuCodeBefore]               CHAR(5)   WITH COMPRESSION,"
    strSQL = strSQL & "[KisyuRyakusyo]                 CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[KisyuRyakusyoBefore]           CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[MinaraiCD]                     CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[MinaraiCDBefore]               CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[BaTaijyu]                      CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[ZogenFugo]                     CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[ZogenSa]                       CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[IJyoCD]                        CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[NyusenJyuni]                   CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[KakuteiJyuni]                  CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[DochakuKubun]                  CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[DochakuTosu]                   CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[Time]                          CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[ChakusaCD]                     CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[ChakusaCDP]                    CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[ChakusaCDPP]                   CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyuni1c]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyuni2c]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyuni3c]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyuni4c]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Odds]                          CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ninki]                         CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Honsyokin]                     CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[Fukasyokin]                    CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[reserved3]                     CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[reserved4]                     CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[HaronTimeL4]                   CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[HaronTimeL3]                   CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[KettoNum1]                     CHAR(10)  WITH COMPRESSION,"
    strSQL = strSQL & "[Bamei1]                        CHAR(36)  WITH COMPRESSION,"
    strSQL = strSQL & "[KettoNum2]                     CHAR(10)  WITH COMPRESSION,"
    strSQL = strSQL & "[Bamei2]                        CHAR(36)  WITH COMPRESSION,"
    strSQL = strSQL & "[KettoNum3]                     CHAR(10)  WITH COMPRESSION,"
    strSQL = strSQL & "[Bamei3]                        CHAR(36)  WITH COMPRESSION,"
    strSQL = strSQL & "[TimeDiff]                      CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[RecordUpKubun]                 CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMKubun]                       CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMTime]                        CHAR(5)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMGosaP]                       CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMGosaM]                       CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMJyuni]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[KyakusituKubun]                CHAR(1)   WITH COMPRESSION"
    strSQL = strSQL & ")"

    mConnection.Execute strSQL

    createTable_UMA_RACE_B = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createTable_UMA_RACE_B = False
End Function ' createTable_UMA_RACE_B


'
'   機能: DBにHARAI(払戻)テーブルを作る
'
'   備考: なし
'
Public Function createTable_HARAI() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String

    strSQL = "CREATE TABLE HARAI("
    strSQL = strSQL & "[RecordSpec]                    CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[DataKubun]                     CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[MakeDate]                      CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[Year]                          CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[MonthDay]                      CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[JyoCD]                         CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kaiji]                         CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Nichiji]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[RaceNum]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[TorokuTosu]                    CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[SyussoTosu]                    CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[FuseirituFlag1]                CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[FuseirituFlag2]                CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[FuseirituFlag3]                CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[FuseirituFlag4]                CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[FuseirituFlag5]                CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[FuseirituFlag6]                CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[FuseirituFlag7]                CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[FuseirituFlag8]                CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[FuseirituFlag9]                CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[TokubaraiFlag1]                CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[TokubaraiFlag2]                CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[TokubaraiFlag3]                CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[TokubaraiFlag4]                CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[TokubaraiFlag5]                CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[TokubaraiFlag6]                CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[TokubaraiFlag7]                CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[TokubaraiFlag8]                CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[TokubaraiFlag9]                CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HenkanFlag1]                   CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HenkanFlag2]                   CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HenkanFlag3]                   CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HenkanFlag4]                   CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HenkanFlag5]                   CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HenkanFlag6]                   CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HenkanFlag7]                   CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HenkanFlag8]                   CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HenkanFlag9]                   CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HenkanUma1]                    CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HenkanUma2]                    CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HenkanUma3]                    CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HenkanUma4]                    CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HenkanUma5]                    CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HenkanUma6]                    CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HenkanUma7]                    CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HenkanUma8]                    CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HenkanUma9]                    CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HenkanUma10]                   CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HenkanUma11]                   CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HenkanUma12]                   CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HenkanUma13]                   CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HenkanUma14]                   CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HenkanUma15]                   CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HenkanUma16]                   CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HenkanUma17]                   CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HenkanUma18]                   CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HenkanUma19]                   CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HenkanUma20]                   CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HenkanUma21]                   CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HenkanUma22]                   CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HenkanUma23]                   CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HenkanUma24]                   CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HenkanUma25]                   CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HenkanUma26]                   CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HenkanUma27]                   CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HenkanUma28]                   CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HenkanWaku1]                   CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HenkanWaku2]                   CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HenkanWaku3]                   CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HenkanWaku4]                   CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HenkanWaku5]                   CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HenkanWaku6]                   CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HenkanWaku7]                   CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HenkanWaku8]                   CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HenkanDoWaku1]                 CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HenkanDoWaku2]                 CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HenkanDoWaku3]                 CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HenkanDoWaku4]                 CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HenkanDoWaku5]                 CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HenkanDoWaku6]                 CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HenkanDoWaku7]                 CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HenkanDoWaku8]                 CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayTansyoUmaban1]              CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayTansyoPay1]                 CHAR(9)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayTansyoNinki1]               CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayTansyoUmaban2]              CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayTansyoPay2]                 CHAR(9)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayTansyoNinki2]               CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayTansyoUmaban3]              CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayTansyoPay3]                 CHAR(9)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayTansyoNinki3]               CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayFukusyoUmaban1]             CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayFukusyoPay1]                CHAR(9)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayFukusyoNinki1]              CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayFukusyoUmaban2]             CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayFukusyoPay2]                CHAR(9)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayFukusyoNinki2]              CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayFukusyoUmaban3]             CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayFukusyoPay3]                CHAR(9)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayFukusyoNinki3]              CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayFukusyoUmaban4]             CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayFukusyoPay4]                CHAR(9)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayFukusyoNinki4]              CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayFukusyoUmaban5]             CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayFukusyoPay5]                CHAR(9)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayFukusyoNinki5]              CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayWakurenKumiban1]            CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayWakurenPay1]                CHAR(9)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayWakurenNinki1]              CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayWakurenKumiban2]            CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayWakurenPay2]                CHAR(9)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayWakurenNinki2]              CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayWakurenKumiban3]            CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayWakurenPay3]                CHAR(9)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayWakurenNinki3]              CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayUmarenKumiban1]             CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayUmarenPay1]                 CHAR(9)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayUmarenNinki1]               CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayUmarenKumiban2]             CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayUmarenPay2]                 CHAR(9)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayUmarenNinki2]               CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayUmarenKumiban3]             CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayUmarenPay3]                 CHAR(9)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayUmarenNinki3]               CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayWideKumiban1]               CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayWidePay1]                   CHAR(9)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayWideNinki1]                 CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayWideKumiban2]               CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayWidePay2]                   CHAR(9)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayWideNinki2]                 CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayWideKumiban3]               CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayWidePay3]                   CHAR(9)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayWideNinki3]                 CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayWideKumiban4]               CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayWidePay4]                   CHAR(9)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayWideNinki4]                 CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayWideKumiban5]               CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayWidePay5]                   CHAR(9)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayWideNinki5]                 CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayWideKumiban6]               CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayWidePay6]                   CHAR(9)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayWideNinki6]                 CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayWideKumiban7]               CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayWidePay7]                   CHAR(9)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayWideNinki7]                 CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayReserved1Kumiban1]          CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayReserved1Pay1]              CHAR(9)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayReserved1Ninki1]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayReserved1Kumiban2]          CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayReserved1Pay2]              CHAR(9)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayReserved1Ninki2]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayReserved1Kumiban3]          CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayReserved1Pay3]              CHAR(9)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayReserved1Ninki3]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayUmatanKumiban1]             CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayUmatanPay1]                 CHAR(9)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayUmatanNinki1]               CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayUmatanKumiban2]             CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayUmatanPay2]                 CHAR(9)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayUmatanNinki2]               CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayUmatanKumiban3]             CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayUmatanPay3]                 CHAR(9)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayUmatanNinki3]               CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayUmatanKumiban4]             CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayUmatanPay4]                 CHAR(9)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayUmatanNinki4]               CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayUmatanKumiban5]             CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayUmatanPay5]                 CHAR(9)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayUmatanNinki5]               CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayUmatanKumiban6]             CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayUmatanPay6]                 CHAR(9)   WITH COMPRESSION,"
    strSQL = strSQL & "[PayUmatanNinki6]               CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[PaySanrenpukuKumiban1]         CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[PaySanrenpukuPay1]             CHAR(9)   WITH COMPRESSION,"
    strSQL = strSQL & "[PaySanrenpukuNinki1]           CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[PaySanrenpukuKumiban2]         CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[PaySanrenpukuPay2]             CHAR(9)   WITH COMPRESSION,"
    strSQL = strSQL & "[PaySanrenpukuNinki2]           CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[PaySanrenpukuKumiban3]         CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[PaySanrenpukuPay3]             CHAR(9)   WITH COMPRESSION,"
    strSQL = strSQL & "[PaySanrenpukuNinki3]           CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[PaySanrentanKumiban1]          CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[PaySanrentanPay1]              CHAR(9)   WITH COMPRESSION,"
    strSQL = strSQL & "[PaySanrentanNinki1]            CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[PaySanrentanKumiban2]          CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[PaySanrentanPay2]              CHAR(9)   WITH COMPRESSION,"
    strSQL = strSQL & "[PaySanrentanNinki2]            CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[PaySanrentanKumiban3]          CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[PaySanrentanPay3]              CHAR(9)   WITH COMPRESSION,"
    strSQL = strSQL & "[PaySanrentanNinki3]            CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[PaySanrentanKumiban4]          CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[PaySanrentanPay4]              CHAR(9)   WITH COMPRESSION,"
    strSQL = strSQL & "[PaySanrentanNinki4]            CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[PaySanrentanKumiban5]          CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[PaySanrentanPay5]              CHAR(9)   WITH COMPRESSION,"
    strSQL = strSQL & "[PaySanrentanNinki5]            CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[PaySanrentanKumiban6]          CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[PaySanrentanPay6]              CHAR(9)   WITH COMPRESSION,"
    strSQL = strSQL & "[PaySanrentanNinki6]            CHAR(4)   WITH COMPRESSION"
    strSQL = strSQL & ")"

    mConnection.Execute strSQL

    createTable_HARAI = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createTable_HARAI = False
End Function ' createTable_HARAI


'
'   機能: DBにODDS_TANPUKUWAKU(オッズ_単複枠)テーブルを作る
'
'   備考: なし
'
Public Function createTable_ODDS_TANPUKUWAKU() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String

    strSQL = "CREATE TABLE ODDS_TANPUKUWAKU("
    strSQL = strSQL & "[RecordSpec]                    CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[DataKubun]                     CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[MakeDate]                      CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[Year]                          CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[MonthDay]                      CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[JyoCD]                         CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kaiji]                         CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Nichiji]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[RaceNum]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[HappyoTime]                    CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[TorokuTosu]                    CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[SyussoTosu]                    CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[TansyoFlag]                    CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[FukusyoFlag]                   CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[WakurenFlag]                   CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[FukuChakuBaraiKey]             CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[TotalHyosuTansyo]              CHAR(11)  WITH COMPRESSION,"
    strSQL = strSQL & "[TotalHyosuFukusyo]             CHAR(11)  WITH COMPRESSION,"
    strSQL = strSQL & "[TotalHyosuWakuren]             CHAR(11)  WITH COMPRESSION,"
    strSQL = strSQL & "[TanData]                       TEXT      WITH COMPRESSION,"
    strSQL = strSQL & "[FukuData]                      TEXT      WITH COMPRESSION,"
    strSQL = strSQL & "[WakuData]                      TEXT      WITH COMPRESSION"
    strSQL = strSQL & ")"

    mConnection.Execute strSQL

    createTable_ODDS_TANPUKUWAKU = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createTable_ODDS_TANPUKUWAKU = False
End Function ' createTable_ODDS_TANPUKUWAKU


'
'   機能: DBにODDS_UMAREN(オッズ_馬連)テーブルを作る
'
'   備考: なし
'
Public Function createTable_ODDS_UMAREN() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String

    strSQL = "CREATE TABLE ODDS_UMAREN("
    strSQL = strSQL & "[RecordSpec]                    CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[DataKubun]                     CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[MakeDate]                      CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[Year]                          CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[MonthDay]                      CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[JyoCD]                         CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kaiji]                         CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Nichiji]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[RaceNum]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[HappyoTime]                    CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[TorokuTosu]                    CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[SyussoTosu]                    CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[UmarenFlag]                    CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[TotalHyosuUmaren]              CHAR(11)  WITH COMPRESSION,"
    strSQL = strSQL & "[Data]                          TEXT      WITH COMPRESSION"
    strSQL = strSQL & ")"

    mConnection.Execute strSQL

    createTable_ODDS_UMAREN = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createTable_ODDS_UMAREN = False
End Function ' createTable_ODDS_UMAREN


'
'   機能: DBにODDS_WIDE(オッズ_ワイド)テーブルを作る
'
'   備考: なし
'
Public Function createTable_ODDS_WIDE() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String

    strSQL = "CREATE TABLE ODDS_WIDE("
    strSQL = strSQL & "[RecordSpec]                    CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[DataKubun]                     CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[MakeDate]                      CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[Year]                          CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[MonthDay]                      CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[JyoCD]                         CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kaiji]                         CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Nichiji]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[RaceNum]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[HappyoTime]                    CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[TorokuTosu]                    CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[SyussoTosu]                    CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[WideFlag]                      CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[TotalHyosuWide]                CHAR(11)  WITH COMPRESSION,"
    strSQL = strSQL & "[Data]                          TEXT      WITH COMPRESSION"
    strSQL = strSQL & ")"

    mConnection.Execute strSQL

    createTable_ODDS_WIDE = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createTable_ODDS_WIDE = False
End Function ' createTable_ODDS_WIDE


'
'   機能: DBにODDS_UMATAN(オッズ_馬単)テーブルを作る
'
'   備考: なし
'
Public Function createTable_ODDS_UMATAN() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String

    strSQL = "CREATE TABLE ODDS_UMATAN("
    strSQL = strSQL & "[RecordSpec]                    CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[DataKubun]                     CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[MakeDate]                      CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[Year]                          CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[MonthDay]                      CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[JyoCD]                         CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kaiji]                         CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Nichiji]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[RaceNum]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[HappyoTime]                    CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[TorokuTosu]                    CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[SyussoTosu]                    CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[UmatanFlag]                    CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[TotalHyosuUmatan]              CHAR(11)  WITH COMPRESSION,"
    strSQL = strSQL & "[Data]                          TEXT      WITH COMPRESSION"
    strSQL = strSQL & ")"

    mConnection.Execute strSQL

    createTable_ODDS_UMATAN = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createTable_ODDS_UMATAN = False
End Function ' createTable_ODDS_UMATAN


'
'   機能: DBにODDS_SANREN(オッズ_3連複)テーブルを作る
'
'   備考: なし
'
Public Function createTable_ODDS_SANREN() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String

    strSQL = "CREATE TABLE ODDS_SANREN("
    strSQL = strSQL & "[RecordSpec]                    CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[DataKubun]                     CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[MakeDate]                      CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[Year]                          CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[MonthDay]                      CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[JyoCD]                         CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kaiji]                         CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Nichiji]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[RaceNum]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[HappyoTime]                    CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[TorokuTosu]                    CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[SyussoTosu]                    CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[SanrenFlag]                    CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[TotalHyosuSanren]              CHAR(11)  WITH COMPRESSION,"
    strSQL = strSQL & "[Data]                          TEXT      WITH COMPRESSION"
    strSQL = strSQL & ")"

    mConnection.Execute strSQL

    createTable_ODDS_SANREN = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createTable_ODDS_SANREN = False
End Function ' createTable_ODDS_SANREN


'
'   機能: DBにUMA(競走馬マスタ)テーブルを作る
'
'   備考: なし
'
Public Function createTable_UMA() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String

    strSQL = "CREATE TABLE UMA("
    strSQL = strSQL & "[RecordSpec]                    CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[DataKubun]                     CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[MakeDate]                      CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[KettoNum]                      CHAR(10)  WITH COMPRESSION,"
    strSQL = strSQL & "[DelKubun]                      CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[RegDate]                       CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[DelDate]                       CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[BirthDate]                     CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[Bamei]                         CHAR(36)  WITH COMPRESSION,"
    strSQL = strSQL & "[BameiKana]                     CHAR(36)  WITH COMPRESSION,"
    strSQL = strSQL & "[BameiEng]                      CHAR(80)  WITH COMPRESSION,"
    strSQL = strSQL & "[UmaKigoCD]                     CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[SexCD]                         CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HinsyuCD]                      CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[KeiroCD]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ketto3InfoHansyokuNum1]        CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ketto3InfoBamei1]              CHAR(36)  WITH COMPRESSION,"
    strSQL = strSQL & "[Ketto3InfoHansyokuNum2]        CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ketto3InfoBamei2]              CHAR(36)  WITH COMPRESSION,"
    strSQL = strSQL & "[Ketto3InfoHansyokuNum3]        CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ketto3InfoBamei3]              CHAR(36)  WITH COMPRESSION,"
    strSQL = strSQL & "[Ketto3InfoHansyokuNum4]        CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ketto3InfoBamei4]              CHAR(36)  WITH COMPRESSION,"
    strSQL = strSQL & "[Ketto3InfoHansyokuNum5]        CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ketto3InfoBamei5]              CHAR(36)  WITH COMPRESSION,"
    strSQL = strSQL & "[Ketto3InfoHansyokuNum6]        CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ketto3InfoBamei6]              CHAR(36)  WITH COMPRESSION,"
    strSQL = strSQL & "[Ketto3InfoHansyokuNum7]        CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ketto3InfoBamei7]              CHAR(36)  WITH COMPRESSION,"
    strSQL = strSQL & "[Ketto3InfoHansyokuNum8]        CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ketto3InfoBamei8]              CHAR(36)  WITH COMPRESSION,"
    strSQL = strSQL & "[Ketto3InfoHansyokuNum9]        CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ketto3InfoBamei9]              CHAR(36)  WITH COMPRESSION,"
    strSQL = strSQL & "[Ketto3InfoHansyokuNum10]       CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ketto3InfoBamei10]             CHAR(36)  WITH COMPRESSION,"
    strSQL = strSQL & "[Ketto3InfoHansyokuNum11]       CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ketto3InfoBamei11]             CHAR(36)  WITH COMPRESSION,"
    strSQL = strSQL & "[Ketto3InfoHansyokuNum12]       CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ketto3InfoBamei12]             CHAR(36)  WITH COMPRESSION,"
    strSQL = strSQL & "[Ketto3InfoHansyokuNum13]       CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ketto3InfoBamei13]             CHAR(36)  WITH COMPRESSION,"
    strSQL = strSQL & "[Ketto3InfoHansyokuNum14]       CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ketto3InfoBamei14]             CHAR(36)  WITH COMPRESSION,"
    strSQL = strSQL & "[TozaiCD]                       CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[ChokyosiCode]                  CHAR(5)   WITH COMPRESSION,"
    strSQL = strSQL & "[ChokyosiRyakusyo]              CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[Syotai]                        CHAR(20)  WITH COMPRESSION,"
    strSQL = strSQL & "[BreederCode]                   CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[BreederName]                   CHAR(70)  WITH COMPRESSION,"
    strSQL = strSQL & "[SanchiName]                    CHAR(20)  WITH COMPRESSION,"
    strSQL = strSQL & "[BanusiCode]                    CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[BanusiName]                    CHAR(64)  WITH COMPRESSION,"
    strSQL = strSQL & "[RuikeiHonsyoHeiti]             CHAR(9)   WITH COMPRESSION,"
    strSQL = strSQL & "[RuikeiHonsyoSyogai]            CHAR(9)   WITH COMPRESSION,"
    strSQL = strSQL & "[RuikeiFukaHeichi]              CHAR(9)   WITH COMPRESSION,"
    strSQL = strSQL & "[RuikeiFukaSyogai]              CHAR(9)   WITH COMPRESSION,"
    strSQL = strSQL & "[RuikeiSyutokuHeichi]           CHAR(9)   WITH COMPRESSION,"
    strSQL = strSQL & "[RuikeiSyutokuSyogai]           CHAR(9)   WITH COMPRESSION,"
    strSQL = strSQL & "[SogoChakukaisu1]               CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[SogoChakukaisu2]               CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[SogoChakukaisu3]               CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[SogoChakukaisu4]               CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[SogoChakukaisu5]               CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[SogoChakukaisu6]               CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[ChuoChakukaisu1]               CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[ChuoChakukaisu2]               CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[ChuoChakukaisu3]               CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[ChuoChakukaisu4]               CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[ChuoChakukaisu5]               CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[ChuoChakukaisu6]               CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ba1Chakukaisu1]                CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ba1Chakukaisu2]                CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ba1Chakukaisu3]                CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ba1Chakukaisu4]                CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ba1Chakukaisu5]                CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ba1Chakukaisu6]                CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ba2Chakukaisu1]                CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ba2Chakukaisu2]                CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ba2Chakukaisu3]                CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ba2Chakukaisu4]                CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ba2Chakukaisu5]                CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ba2Chakukaisu6]                CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ba3Chakukaisu1]                CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ba3Chakukaisu2]                CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ba3Chakukaisu3]                CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ba3Chakukaisu4]                CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ba3Chakukaisu5]                CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ba3Chakukaisu6]                CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ba4Chakukaisu1]                CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ba4Chakukaisu2]                CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ba4Chakukaisu3]                CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ba4Chakukaisu4]                CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ba4Chakukaisu5]                CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ba4Chakukaisu6]                CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ba5Chakukaisu1]                CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ba5Chakukaisu2]                CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ba5Chakukaisu3]                CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ba5Chakukaisu4]                CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ba5Chakukaisu5]                CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ba5Chakukaisu6]                CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ba6Chakukaisu1]                CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ba6Chakukaisu2]                CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ba6Chakukaisu3]                CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ba6Chakukaisu4]                CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ba6Chakukaisu5]                CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ba6Chakukaisu6]                CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ba7Chakukaisu1]                CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ba7Chakukaisu2]                CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ba7Chakukaisu3]                CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ba7Chakukaisu4]                CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ba7Chakukaisu5]                CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Ba7Chakukaisu6]                CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai1Chakukaisu1]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai1Chakukaisu2]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai1Chakukaisu3]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai1Chakukaisu4]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai1Chakukaisu5]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai1Chakukaisu6]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai2Chakukaisu1]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai2Chakukaisu2]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai2Chakukaisu3]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai2Chakukaisu4]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai2Chakukaisu5]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai2Chakukaisu6]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai3Chakukaisu1]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai3Chakukaisu2]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai3Chakukaisu3]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai3Chakukaisu4]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai3Chakukaisu5]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai3Chakukaisu6]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai4Chakukaisu1]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai4Chakukaisu2]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai4Chakukaisu3]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai4Chakukaisu4]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai4Chakukaisu5]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai4Chakukaisu6]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai5Chakukaisu1]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai5Chakukaisu2]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai5Chakukaisu3]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai5Chakukaisu4]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai5Chakukaisu5]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai5Chakukaisu6]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai6Chakukaisu1]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai6Chakukaisu2]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai6Chakukaisu3]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai6Chakukaisu4]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai6Chakukaisu5]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai6Chakukaisu6]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai7Chakukaisu1]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai7Chakukaisu2]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai7Chakukaisu3]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai7Chakukaisu4]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai7Chakukaisu5]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai7Chakukaisu6]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai8Chakukaisu1]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai8Chakukaisu2]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai8Chakukaisu3]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai8Chakukaisu4]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai8Chakukaisu5]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai8Chakukaisu6]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai9Chakukaisu1]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai9Chakukaisu2]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai9Chakukaisu3]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai9Chakukaisu4]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai9Chakukaisu5]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai9Chakukaisu6]            CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai10Chakukaisu1]           CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai10Chakukaisu2]           CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai10Chakukaisu3]           CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai10Chakukaisu4]           CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai10Chakukaisu5]           CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai10Chakukaisu6]           CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai11Chakukaisu1]           CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai11Chakukaisu2]           CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai11Chakukaisu3]           CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai11Chakukaisu4]           CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai11Chakukaisu5]           CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai11Chakukaisu6]           CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai12Chakukaisu1]           CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai12Chakukaisu2]           CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai12Chakukaisu3]           CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai12Chakukaisu4]           CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai12Chakukaisu5]           CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyotai12Chakukaisu6]           CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori1Chakukaisu1]             CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori1Chakukaisu2]             CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori1Chakukaisu3]             CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori1Chakukaisu4]             CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori1Chakukaisu5]             CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori1Chakukaisu6]             CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori2Chakukaisu1]             CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori2Chakukaisu2]             CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori2Chakukaisu3]             CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori2Chakukaisu4]             CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori2Chakukaisu5]             CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori2Chakukaisu6]             CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori3Chakukaisu1]             CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori3Chakukaisu2]             CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori3Chakukaisu3]             CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori3Chakukaisu4]             CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori3Chakukaisu5]             CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori3Chakukaisu6]             CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori4Chakukaisu1]             CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori4Chakukaisu2]             CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori4Chakukaisu3]             CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori4Chakukaisu4]             CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori4Chakukaisu5]             CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori4Chakukaisu6]             CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori5Chakukaisu1]             CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori5Chakukaisu2]             CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori5Chakukaisu3]             CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori5Chakukaisu4]             CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori5Chakukaisu5]             CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori5Chakukaisu6]             CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori6Chakukaisu1]             CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori6Chakukaisu2]             CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori6Chakukaisu3]             CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori6Chakukaisu4]             CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori6Chakukaisu5]             CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori6Chakukaisu6]             CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyakusitu1]                    CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyakusitu2]                    CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyakusitu3]                    CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyakusitu4]                    CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[RaceCount]                     CHAR(3)   WITH COMPRESSION"
    strSQL = strSQL & ")"

    mConnection.Execute strSQL

    createTable_UMA = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createTable_UMA = False
End Function ' createTable_UMA


'
'   機能: DBにKISHU(騎手マスタ)テーブルを作る
'
'   備考: なし
'
Public Function createTable_KISHU() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String

    strSQL = "CREATE TABLE KISHU("
    strSQL = strSQL & "[RecordSpec]                    CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[DataKubun]                     CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[MakeDate]                      CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[KisyuCode]                     CHAR(5)   WITH COMPRESSION,"
    strSQL = strSQL & "[DelKubun]                      CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[IssueDate]                     CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[DelDate]                       CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[BirthDate]                     CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[KisyuName]                     CHAR(34)  WITH COMPRESSION,"
    strSQL = strSQL & "[reserved]                      CHAR(34)  WITH COMPRESSION,"
    strSQL = strSQL & "[KisyuNameKana]                 CHAR(30)  WITH COMPRESSION,"
    strSQL = strSQL & "[KisyuRyakusyo]                 CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[KisyuNameEng]                  CHAR(80)  WITH COMPRESSION,"
    strSQL = strSQL & "[SexCD]                         CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[SikakuCD]                      CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[MinaraiCD]                     CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[TozaiCD]                       CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[Syotai]                        CHAR(20)  WITH COMPRESSION,"
    strSQL = strSQL & "[ChokyosiCode]                  CHAR(5)   WITH COMPRESSION,"
    strSQL = strSQL & "[ChokyosiRyakusyo]              CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[HatuKiJyo1Hatukijyoid]         CHAR(16)  WITH COMPRESSION,"
    strSQL = strSQL & "[HatuKiJyo1SyussoTosu]          CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[HatuKiJyo1KettoNum]            CHAR(10)  WITH COMPRESSION,"
    strSQL = strSQL & "[HatuKiJyo1Bamei]               CHAR(36)  WITH COMPRESSION,"
    strSQL = strSQL & "[HatuKiJyo1KakuteiJyuni]        CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[HatuKiJyo1IJyoCD]              CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HatuKiJyo2Hatukijyoid]         CHAR(16)  WITH COMPRESSION,"
    strSQL = strSQL & "[HatuKiJyo2SyussoTosu]          CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[HatuKiJyo2KettoNum]            CHAR(10)  WITH COMPRESSION,"
    strSQL = strSQL & "[HatuKiJyo2Bamei]               CHAR(36)  WITH COMPRESSION,"
    strSQL = strSQL & "[HatuKiJyo2KakuteiJyuni]        CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[HatuKiJyo2IJyoCD]              CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HatuSyori1Hatusyoriid]         CHAR(16)  WITH COMPRESSION,"
    strSQL = strSQL & "[HatuSyori1SyussoTosu]          CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[HatuSyori1KettoNum]            CHAR(10)  WITH COMPRESSION,"
    strSQL = strSQL & "[HatuSyori1Bamei]               CHAR(36)  WITH COMPRESSION,"
    strSQL = strSQL & "[HatuSyori2Hatusyoriid]         CHAR(16)  WITH COMPRESSION,"
    strSQL = strSQL & "[HatuSyori2SyussoTosu]          CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[HatuSyori2KettoNum]            CHAR(10)  WITH COMPRESSION,"
    strSQL = strSQL & "[HatuSyori2Bamei]               CHAR(36)  WITH COMPRESSION,"
    strSQL = strSQL & "[SaikinJyusyo1SaikinJyusyoid]   CHAR(16)  WITH COMPRESSION,"
    strSQL = strSQL & "[SaikinJyusyo1Hondai]           CHAR(60)  WITH COMPRESSION,"
    strSQL = strSQL & "[SaikinJyusyo1Ryakusyo10]       CHAR(20)  WITH COMPRESSION,"
    strSQL = strSQL & "[SaikinJyusyo1Ryakusyo6]        CHAR(12)  WITH COMPRESSION,"
    strSQL = strSQL & "[SaikinJyusyo1Ryakusyo3]        CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[SaikinJyusyo1GradeCD]          CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[SaikinJyusyo1SyussoTosu]       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[SaikinJyusyo1KettoNum]         CHAR(10)  WITH COMPRESSION,"
    strSQL = strSQL & "[SaikinJyusyo1Bamei]            CHAR(36)  WITH COMPRESSION,"
    strSQL = strSQL & "[SaikinJyusyo2SaikinJyusyoid]   CHAR(16)  WITH COMPRESSION,"
    strSQL = strSQL & "[SaikinJyusyo2Hondai]           CHAR(60)  WITH COMPRESSION,"
    strSQL = strSQL & "[SaikinJyusyo2Ryakusyo10]       CHAR(20)  WITH COMPRESSION,"
    strSQL = strSQL & "[SaikinJyusyo2Ryakusyo6]        CHAR(12)  WITH COMPRESSION,"
    strSQL = strSQL & "[SaikinJyusyo2Ryakusyo3]        CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[SaikinJyusyo2GradeCD]          CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[SaikinJyusyo2SyussoTosu]       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[SaikinJyusyo2KettoNum]         CHAR(10)  WITH COMPRESSION,"
    strSQL = strSQL & "[SaikinJyusyo2Bamei]            CHAR(36)  WITH COMPRESSION,"
    strSQL = strSQL & "[SaikinJyusyo3SaikinJyusyoid]   CHAR(16)  WITH COMPRESSION,"
    strSQL = strSQL & "[SaikinJyusyo3Hondai]           CHAR(60)  WITH COMPRESSION,"
    strSQL = strSQL & "[SaikinJyusyo3Ryakusyo10]       CHAR(20)  WITH COMPRESSION,"
    strSQL = strSQL & "[SaikinJyusyo3Ryakusyo6]        CHAR(12)  WITH COMPRESSION,"
    strSQL = strSQL & "[SaikinJyusyo3Ryakusyo3]        CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[SaikinJyusyo3GradeCD]          CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[SaikinJyusyo3SyussoTosu]       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[SaikinJyusyo3KettoNum]         CHAR(10)  WITH COMPRESSION,"
    strSQL = strSQL & "[SaikinJyusyo3Bamei]            CHAR(36)  WITH COMPRESSION"
    strSQL = strSQL & ")"

    mConnection.Execute strSQL

    createTable_KISHU = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createTable_KISHU = False
End Function ' createTable_KISHU


'
'   機能: DBにKISHU_SEISEKI(騎手マスタ_成績)テーブルを作る
'
'   備考: なし
'
Public Function createTable_KISHU_SEISEKI() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String

    strSQL = "CREATE TABLE KISHU_SEISEKI("
    strSQL = strSQL & "[KisyuCode]                     CHAR(5)   WITH COMPRESSION,"
    strSQL = strSQL & "[Num]                           CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[SetYear]                       CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[HonSyokinHeichi]               CHAR(10)  WITH COMPRESSION,"
    strSQL = strSQL & "[HonSyokinSyogai]               CHAR(10)  WITH COMPRESSION,"
    strSQL = strSQL & "[FukaSyokinHeichi]              CHAR(10)  WITH COMPRESSION,"
    strSQL = strSQL & "[FukaSyokinSyogai]              CHAR(10)  WITH COMPRESSION,"
    strSQL = strSQL & "[HeichiChakukaisu1]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[HeichiChakukaisu2]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[HeichiChakukaisu3]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[HeichiChakukaisu4]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[HeichiChakukaisu5]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[HeichiChakukaisu6]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[SyogaiChakukaisu1]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[SyogaiChakukaisu2]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[SyogaiChakukaisu3]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[SyogaiChakukaisu4]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[SyogaiChakukaisu5]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[SyogaiChakukaisu6]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo1Chakukaisu1]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo1Chakukaisu2]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo1Chakukaisu3]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo1Chakukaisu4]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo1Chakukaisu5]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo1Chakukaisu6]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo2Chakukaisu1]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo2Chakukaisu2]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo2Chakukaisu3]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo2Chakukaisu4]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo2Chakukaisu5]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo2Chakukaisu6]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo3Chakukaisu1]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo3Chakukaisu2]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo3Chakukaisu3]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo3Chakukaisu4]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo3Chakukaisu5]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo3Chakukaisu6]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo4Chakukaisu1]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo4Chakukaisu2]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo4Chakukaisu3]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo4Chakukaisu4]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo4Chakukaisu5]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo4Chakukaisu6]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo5Chakukaisu1]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo5Chakukaisu2]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo5Chakukaisu3]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo5Chakukaisu4]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo5Chakukaisu5]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo5Chakukaisu6]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo6Chakukaisu1]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo6Chakukaisu2]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo6Chakukaisu3]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo6Chakukaisu4]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo6Chakukaisu5]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo6Chakukaisu6]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo7Chakukaisu1]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo7Chakukaisu2]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo7Chakukaisu3]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo7Chakukaisu4]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo7Chakukaisu5]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo7Chakukaisu6]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo8Chakukaisu1]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo8Chakukaisu2]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo8Chakukaisu3]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo8Chakukaisu4]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo8Chakukaisu5]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo8Chakukaisu6]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo9Chakukaisu1]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo9Chakukaisu2]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo9Chakukaisu3]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo9Chakukaisu4]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo9Chakukaisu5]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo9Chakukaisu6]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo10Chakukaisu1]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo10Chakukaisu2]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo10Chakukaisu3]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo10Chakukaisu4]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo10Chakukaisu5]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo10Chakukaisu6]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo11Chakukaisu1]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo11Chakukaisu2]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo11Chakukaisu3]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo11Chakukaisu4]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo11Chakukaisu5]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo11Chakukaisu6]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo12Chakukaisu1]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo12Chakukaisu2]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo12Chakukaisu3]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo12Chakukaisu4]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo12Chakukaisu5]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo12Chakukaisu6]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo13Chakukaisu1]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo13Chakukaisu2]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo13Chakukaisu3]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo13Chakukaisu4]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo13Chakukaisu5]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo13Chakukaisu6]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo14Chakukaisu1]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo14Chakukaisu2]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo14Chakukaisu3]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo14Chakukaisu4]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo14Chakukaisu5]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo14Chakukaisu6]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo15Chakukaisu1]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo15Chakukaisu2]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo15Chakukaisu3]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo15Chakukaisu4]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo15Chakukaisu5]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo15Chakukaisu6]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo16Chakukaisu1]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo16Chakukaisu2]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo16Chakukaisu3]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo16Chakukaisu4]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo16Chakukaisu5]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo16Chakukaisu6]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo17Chakukaisu1]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo17Chakukaisu2]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo17Chakukaisu3]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo17Chakukaisu4]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo17Chakukaisu5]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo17Chakukaisu6]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo18Chakukaisu1]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo18Chakukaisu2]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo18Chakukaisu3]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo18Chakukaisu4]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo18Chakukaisu5]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo18Chakukaisu6]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo19Chakukaisu1]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo19Chakukaisu2]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo19Chakukaisu3]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo19Chakukaisu4]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo19Chakukaisu5]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo19Chakukaisu6]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo20Chakukaisu1]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo20Chakukaisu2]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo20Chakukaisu3]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo20Chakukaisu4]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo20Chakukaisu5]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo20Chakukaisu6]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori1Chakukaisu1]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori1Chakukaisu2]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori1Chakukaisu3]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori1Chakukaisu4]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori1Chakukaisu5]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori1Chakukaisu6]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori2Chakukaisu1]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori2Chakukaisu2]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori2Chakukaisu3]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori2Chakukaisu4]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori2Chakukaisu5]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori2Chakukaisu6]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori3Chakukaisu1]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori3Chakukaisu2]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori3Chakukaisu3]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori3Chakukaisu4]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori3Chakukaisu5]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori3Chakukaisu6]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori4Chakukaisu1]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori4Chakukaisu2]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori4Chakukaisu3]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori4Chakukaisu4]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori4Chakukaisu5]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori4Chakukaisu6]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori5Chakukaisu1]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori5Chakukaisu2]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori5Chakukaisu3]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori5Chakukaisu4]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori5Chakukaisu5]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori5Chakukaisu6]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori6Chakukaisu1]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori6Chakukaisu2]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori6Chakukaisu3]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori6Chakukaisu4]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori6Chakukaisu5]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori6Chakukaisu6]             CHAR(6)   WITH COMPRESSION"
    strSQL = strSQL & ")"

    mConnection.Execute strSQL

    createTable_KISHU_SEISEKI = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createTable_KISHU_SEISEKI = False
End Function ' createTable_KISHU_SEISEKI


'
'   機能: DBにCHOKYO(調教師マスタ)テーブルを作る
'
'   備考: なし
'
Public Function createTable_CHOKYO() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String

    strSQL = "CREATE TABLE CHOKYO("
    strSQL = strSQL & "[RecordSpec]                    CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[DataKubun]                     CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[MakeDate]                      CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[ChokyosiCode]                  CHAR(5)   WITH COMPRESSION,"
    strSQL = strSQL & "[DelKubun]                      CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[IssueDate]                     CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[DelDate]                       CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[BirthDate]                     CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[ChokyosiName]                  CHAR(34)  WITH COMPRESSION,"
    strSQL = strSQL & "[ChokyosiNameKana]              CHAR(30)  WITH COMPRESSION,"
    strSQL = strSQL & "[ChokyosiRyakusyo]              CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[ChokyosiNameEng]               CHAR(80)  WITH COMPRESSION,"
    strSQL = strSQL & "[SexCD]                         CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[TozaiCD]                       CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[Syotai]                        CHAR(20)  WITH COMPRESSION,"
    strSQL = strSQL & "[SaikinJyusyo1SaikinJyusyoid]   CHAR(16)  WITH COMPRESSION,"
    strSQL = strSQL & "[SaikinJyusyo1Hondai]           CHAR(60)  WITH COMPRESSION,"
    strSQL = strSQL & "[SaikinJyusyo1Ryakusyo10]       CHAR(20)  WITH COMPRESSION,"
    strSQL = strSQL & "[SaikinJyusyo1Ryakusyo6]        CHAR(12)  WITH COMPRESSION,"
    strSQL = strSQL & "[SaikinJyusyo1Ryakusyo3]        CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[SaikinJyusyo1GradeCD]          CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[SaikinJyusyo1SyussoTosu]       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[SaikinJyusyo1KettoNum]         CHAR(10)  WITH COMPRESSION,"
    strSQL = strSQL & "[SaikinJyusyo1Bamei]            CHAR(36)  WITH COMPRESSION,"
    strSQL = strSQL & "[SaikinJyusyo2SaikinJyusyoid]   CHAR(16)  WITH COMPRESSION,"
    strSQL = strSQL & "[SaikinJyusyo2Hondai]           CHAR(60)  WITH COMPRESSION,"
    strSQL = strSQL & "[SaikinJyusyo2Ryakusyo10]       CHAR(20)  WITH COMPRESSION,"
    strSQL = strSQL & "[SaikinJyusyo2Ryakusyo6]        CHAR(12)  WITH COMPRESSION,"
    strSQL = strSQL & "[SaikinJyusyo2Ryakusyo3]        CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[SaikinJyusyo2GradeCD]          CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[SaikinJyusyo2SyussoTosu]       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[SaikinJyusyo2KettoNum]         CHAR(10)  WITH COMPRESSION,"
    strSQL = strSQL & "[SaikinJyusyo2Bamei]            CHAR(36)  WITH COMPRESSION,"
    strSQL = strSQL & "[SaikinJyusyo3SaikinJyusyoid]   CHAR(16)  WITH COMPRESSION,"
    strSQL = strSQL & "[SaikinJyusyo3Hondai]           CHAR(60)  WITH COMPRESSION,"
    strSQL = strSQL & "[SaikinJyusyo3Ryakusyo10]       CHAR(20)  WITH COMPRESSION,"
    strSQL = strSQL & "[SaikinJyusyo3Ryakusyo6]        CHAR(12)  WITH COMPRESSION,"
    strSQL = strSQL & "[SaikinJyusyo3Ryakusyo3]        CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[SaikinJyusyo3GradeCD]          CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[SaikinJyusyo3SyussoTosu]       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[SaikinJyusyo3KettoNum]         CHAR(10)  WITH COMPRESSION,"
    strSQL = strSQL & "[SaikinJyusyo3Bamei]            CHAR(36)  WITH COMPRESSION"
    strSQL = strSQL & ")"

    mConnection.Execute strSQL

    createTable_CHOKYO = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createTable_CHOKYO = False
End Function ' createTable_CHOKYO


'
'   機能: DBにCHOKYO_SEISEKI(調教師マスタ_成績)テーブルを作る
'
'   備考: なし
'
Public Function createTable_CHOKYO_SEISEKI() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String

    strSQL = "CREATE TABLE CHOKYO_SEISEKI("
    strSQL = strSQL & "[ChokyosiCode]                  CHAR(5)   WITH COMPRESSION,"
    strSQL = strSQL & "[Num]                           CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[SetYear]                       CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[HonSyokinHeichi]               CHAR(10)  WITH COMPRESSION,"
    strSQL = strSQL & "[HonSyokinSyogai]               CHAR(10)  WITH COMPRESSION,"
    strSQL = strSQL & "[FukaSyokinHeichi]              CHAR(10)  WITH COMPRESSION,"
    strSQL = strSQL & "[FukaSyokinSyogai]              CHAR(10)  WITH COMPRESSION,"
    strSQL = strSQL & "[HeichiChakukaisu1]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[HeichiChakukaisu2]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[HeichiChakukaisu3]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[HeichiChakukaisu4]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[HeichiChakukaisu5]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[HeichiChakukaisu6]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[SyogaiChakukaisu1]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[SyogaiChakukaisu2]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[SyogaiChakukaisu3]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[SyogaiChakukaisu4]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[SyogaiChakukaisu5]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[SyogaiChakukaisu6]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo1Chakukaisu1]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo1Chakukaisu2]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo1Chakukaisu3]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo1Chakukaisu4]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo1Chakukaisu5]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo1Chakukaisu6]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo2Chakukaisu1]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo2Chakukaisu2]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo2Chakukaisu3]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo2Chakukaisu4]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo2Chakukaisu5]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo2Chakukaisu6]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo3Chakukaisu1]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo3Chakukaisu2]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo3Chakukaisu3]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo3Chakukaisu4]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo3Chakukaisu5]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo3Chakukaisu6]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo4Chakukaisu1]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo4Chakukaisu2]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo4Chakukaisu3]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo4Chakukaisu4]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo4Chakukaisu5]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo4Chakukaisu6]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo5Chakukaisu1]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo5Chakukaisu2]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo5Chakukaisu3]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo5Chakukaisu4]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo5Chakukaisu5]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo5Chakukaisu6]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo6Chakukaisu1]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo6Chakukaisu2]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo6Chakukaisu3]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo6Chakukaisu4]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo6Chakukaisu5]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo6Chakukaisu6]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo7Chakukaisu1]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo7Chakukaisu2]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo7Chakukaisu3]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo7Chakukaisu4]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo7Chakukaisu5]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo7Chakukaisu6]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo8Chakukaisu1]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo8Chakukaisu2]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo8Chakukaisu3]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo8Chakukaisu4]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo8Chakukaisu5]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo8Chakukaisu6]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo9Chakukaisu1]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo9Chakukaisu2]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo9Chakukaisu3]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo9Chakukaisu4]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo9Chakukaisu5]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo9Chakukaisu6]               CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo10Chakukaisu1]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo10Chakukaisu2]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo10Chakukaisu3]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo10Chakukaisu4]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo10Chakukaisu5]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo10Chakukaisu6]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo11Chakukaisu1]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo11Chakukaisu2]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo11Chakukaisu3]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo11Chakukaisu4]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo11Chakukaisu5]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo11Chakukaisu6]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo12Chakukaisu1]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo12Chakukaisu2]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo12Chakukaisu3]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo12Chakukaisu4]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo12Chakukaisu5]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo12Chakukaisu6]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo13Chakukaisu1]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo13Chakukaisu2]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo13Chakukaisu3]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo13Chakukaisu4]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo13Chakukaisu5]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo13Chakukaisu6]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo14Chakukaisu1]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo14Chakukaisu2]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo14Chakukaisu3]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo14Chakukaisu4]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo14Chakukaisu5]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo14Chakukaisu6]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo15Chakukaisu1]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo15Chakukaisu2]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo15Chakukaisu3]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo15Chakukaisu4]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo15Chakukaisu5]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo15Chakukaisu6]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo16Chakukaisu1]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo16Chakukaisu2]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo16Chakukaisu3]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo16Chakukaisu4]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo16Chakukaisu5]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo16Chakukaisu6]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo17Chakukaisu1]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo17Chakukaisu2]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo17Chakukaisu3]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo17Chakukaisu4]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo17Chakukaisu5]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo17Chakukaisu6]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo18Chakukaisu1]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo18Chakukaisu2]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo18Chakukaisu3]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo18Chakukaisu4]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo18Chakukaisu5]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo18Chakukaisu6]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo19Chakukaisu1]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo19Chakukaisu2]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo19Chakukaisu3]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo19Chakukaisu4]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo19Chakukaisu5]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo19Chakukaisu6]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo20Chakukaisu1]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo20Chakukaisu2]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo20Chakukaisu3]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo20Chakukaisu4]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo20Chakukaisu5]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyo20Chakukaisu6]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori1Chakukaisu1]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori1Chakukaisu2]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori1Chakukaisu3]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori1Chakukaisu4]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori1Chakukaisu5]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori1Chakukaisu6]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori2Chakukaisu1]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori2Chakukaisu2]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori2Chakukaisu3]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori2Chakukaisu4]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori2Chakukaisu5]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori2Chakukaisu6]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori3Chakukaisu1]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori3Chakukaisu2]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori3Chakukaisu3]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori3Chakukaisu4]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori3Chakukaisu5]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori3Chakukaisu6]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori4Chakukaisu1]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori4Chakukaisu2]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori4Chakukaisu3]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori4Chakukaisu4]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori4Chakukaisu5]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori4Chakukaisu6]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori5Chakukaisu1]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori5Chakukaisu2]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori5Chakukaisu3]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori5Chakukaisu4]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori5Chakukaisu5]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori5Chakukaisu6]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori6Chakukaisu1]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori6Chakukaisu2]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori6Chakukaisu3]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori6Chakukaisu4]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori6Chakukaisu5]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori6Chakukaisu6]             CHAR(6)   WITH COMPRESSION"
    strSQL = strSQL & ")"

    mConnection.Execute strSQL

    createTable_CHOKYO_SEISEKI = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createTable_CHOKYO_SEISEKI = False
End Function ' createTable_CHOKYO_SEISEKI


'
'   機能: DBにSEISAN(生産者マスタ)テーブルを作る
'
'   備考: なし
'
Public Function createTable_SEISAN() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String

    strSQL = "CREATE TABLE SEISAN("
    strSQL = strSQL & "[RecordSpec]                    CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[DataKubun]                     CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[MakeDate]                      CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[BreederCode]                   CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[BreederName_Co]                CHAR(70)  WITH COMPRESSION,"
    strSQL = strSQL & "[BreederName]                   CHAR(70)  WITH COMPRESSION,"
    strSQL = strSQL & "[BreederNameKana]               CHAR(70)  WITH COMPRESSION,"
    strSQL = strSQL & "[BreederNameEng]                CHAR(168) WITH COMPRESSION,"
    strSQL = strSQL & "[Address]                       CHAR(20)  WITH COMPRESSION,"
    strSQL = strSQL & "[H_SetYear]                     CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[H_HonSyokinTotal]              CHAR(10)  WITH COMPRESSION,"
    strSQL = strSQL & "[H_FukaSyokin]                  CHAR(10)  WITH COMPRESSION,"
    strSQL = strSQL & "[H_ChakuKaisu1]                 CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[H_ChakuKaisu2]                 CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[H_ChakuKaisu3]                 CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[H_ChakuKaisu4]                 CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[H_ChakuKaisu5]                 CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[H_ChakuKaisu6]                 CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[R_SetYear]                     CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[R_HonSyokinTotal]              CHAR(10)  WITH COMPRESSION,"
    strSQL = strSQL & "[R_FukaSyokin]                  CHAR(10)  WITH COMPRESSION,"
    strSQL = strSQL & "[R_ChakuKaisu1]                 CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[R_ChakuKaisu2]                 CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[R_ChakuKaisu3]                 CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[R_ChakuKaisu4]                 CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[R_ChakuKaisu5]                 CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[R_ChakuKaisu6]                 CHAR(6)   WITH COMPRESSION"
    strSQL = strSQL & ")"

    mConnection.Execute strSQL

    createTable_SEISAN = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createTable_SEISAN = False
End Function ' createTable_SEISAN


'
'   機能: DBにBANUSI(馬主マスタ)テーブルを作る
'
'   備考: なし
'
Public Function createTable_BANUSI() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String

    strSQL = "CREATE TABLE BANUSI("
    strSQL = strSQL & "[RecordSpec]                    CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[DataKubun]                     CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[MakeDate]                      CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[BanusiCode]                    CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[BanusiName_Co]                 CHAR(64)  WITH COMPRESSION,"
    strSQL = strSQL & "[BanusiName]                    CHAR(64)  WITH COMPRESSION,"
    strSQL = strSQL & "[BanusiNameKana]                CHAR(50)  WITH COMPRESSION,"
    strSQL = strSQL & "[BanusiNameEng]                 CHAR(100) WITH COMPRESSION,"
    strSQL = strSQL & "[Fukusyoku]                     CHAR(60)  WITH COMPRESSION,"
    strSQL = strSQL & "[H_SetYear]                     CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[H_HonSyokinTotal]              CHAR(10)  WITH COMPRESSION,"
    strSQL = strSQL & "[H_FukaSyokin]                  CHAR(10)  WITH COMPRESSION,"
    strSQL = strSQL & "[H_ChakuKaisu1]                 CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[H_ChakuKaisu2]                 CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[H_ChakuKaisu3]                 CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[H_ChakuKaisu4]                 CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[H_ChakuKaisu5]                 CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[H_ChakuKaisu6]                 CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[R_SetYear]                     CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[R_HonSyokinTotal]              CHAR(10)  WITH COMPRESSION,"
    strSQL = strSQL & "[R_FukaSyokin]                  CHAR(10)  WITH COMPRESSION,"
    strSQL = strSQL & "[R_ChakuKaisu1]                 CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[R_ChakuKaisu2]                 CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[R_ChakuKaisu3]                 CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[R_ChakuKaisu4]                 CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[R_ChakuKaisu5]                 CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[R_ChakuKaisu6]                 CHAR(6)   WITH COMPRESSION"
    strSQL = strSQL & ")"

    mConnection.Execute strSQL

    createTable_BANUSI = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createTable_BANUSI = False
End Function ' createTable_BANUSI


'
'   機能: DBにHANSYOKU(繁殖馬マスタ)テーブルを作る
'
'   備考: なし
'
Public Function createTable_HANSYOKU() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String

    strSQL = "CREATE TABLE HANSYOKU("
    strSQL = strSQL & "[RecordSpec]                    CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[DataKubun]                     CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[MakeDate]                      CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[HansyokuNum]                   CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[reserved]                      CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[KettoNum]                      CHAR(10)  WITH COMPRESSION,"
    strSQL = strSQL & "[DelKubun]                      CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[Bamei]                         CHAR(36)  WITH COMPRESSION,"
    strSQL = strSQL & "[BameiKana]                     CHAR(40)  WITH COMPRESSION,"
    strSQL = strSQL & "[BameiEng]                      CHAR(80)  WITH COMPRESSION,"
    strSQL = strSQL & "[BirthYear]                     CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[SexCD]                         CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HinsyuCD]                      CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[KeiroCD]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[HansyokuMochiKubun]            CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[ImportYear]                    CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[SanchiName]                    CHAR(20)  WITH COMPRESSION,"
    strSQL = strSQL & "[HansyokuFNum]                  CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[HansyokuMNum]                  CHAR(8)   WITH COMPRESSION"
    strSQL = strSQL & ")"

    mConnection.Execute strSQL

    createTable_HANSYOKU = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createTable_HANSYOKU = False
End Function ' createTable_HANSYOKU


'
'   機能: DBにSANKU(産駒マスタ)テーブルを作る
'
'   備考: なし
'
Public Function createTable_SANKU() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String

    strSQL = "CREATE TABLE SANKU("
    strSQL = strSQL & "[RecordSpec]                    CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[DataKubun]                     CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[MakeDate]                      CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[KettoNum]                      CHAR(10)  WITH COMPRESSION,"
    strSQL = strSQL & "[BirthDate]                     CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[SexCD]                         CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[HinsyuCD]                      CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[KeiroCD]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[SankuMochiKubun]               CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[ImportYear]                    CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[BreederCode]                   CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[SanchiName]                    CHAR(20)  WITH COMPRESSION,"
    strSQL = strSQL & "[FNum]                          CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[MNum]                          CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[FFNum]                         CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[FMNum]                         CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[MFNum]                         CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[MMNum]                         CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[FFFNum]                        CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[FFMNum]                        CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[FMFNum]                        CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[FMMNum]                        CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[MFFNum]                        CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[MFMNum]                        CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[MMFNum]                        CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[MMMNum]                        CHAR(8)   WITH COMPRESSION"
    strSQL = strSQL & ")"

    mConnection.Execute strSQL

    createTable_SANKU = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createTable_SANKU = False
End Function ' createTable_SANKU


'
'   機能: DBにRECORD(レコードマスタ)テーブルを作る
'
'   備考: なし
'
Public Function createTable_RECORD() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String

    strSQL = "CREATE TABLE RECORD("
    strSQL = strSQL & "[RecordSpec]                    CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[DataKubun]                     CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[MakeDate]                      CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[RecInfoKubun]                  CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[Year]                          CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[MonthDay]                      CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[JyoCD]                         CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kaiji]                         CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Nichiji]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[RaceNum]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[TokuNum_SyubetuCD]             CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Hondai]                        CHAR(60)  WITH COMPRESSION,"
    strSQL = strSQL & "[GradeCD]                       CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kyori]                         CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[TrackCD]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[RecKubun]                      CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[RecTime]                       CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[TenkoCD]                       CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[SibaBabaCD]                    CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[DirtBabaCD]                    CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[RecUmaKettoNum1]               CHAR(10)  WITH COMPRESSION,"
    strSQL = strSQL & "[RecUmaBamei1]                  CHAR(36)  WITH COMPRESSION,"
    strSQL = strSQL & "[RecUmaUmaKigoCD1]              CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[RecUmaSexCD1]                  CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[RecUmaChokyosiCode1]           CHAR(5)   WITH COMPRESSION,"
    strSQL = strSQL & "[RecUmaChokyosiName1]           CHAR(34)  WITH COMPRESSION,"
    strSQL = strSQL & "[RecUmaFutan1]                  CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[RecUmaKisyuCode1]              CHAR(5)   WITH COMPRESSION,"
    strSQL = strSQL & "[RecUmaKisyuName1]              CHAR(34)  WITH COMPRESSION,"
    strSQL = strSQL & "[RecUmaKettoNum2]               CHAR(10)  WITH COMPRESSION,"
    strSQL = strSQL & "[RecUmaBamei2]                  CHAR(36)  WITH COMPRESSION,"
    strSQL = strSQL & "[RecUmaUmaKigoCD2]              CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[RecUmaSexCD2]                  CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[RecUmaChokyosiCode2]           CHAR(5)   WITH COMPRESSION,"
    strSQL = strSQL & "[RecUmaChokyosiName2]           CHAR(34)  WITH COMPRESSION,"
    strSQL = strSQL & "[RecUmaFutan2]                  CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[RecUmaKisyuCode2]              CHAR(5)   WITH COMPRESSION,"
    strSQL = strSQL & "[RecUmaKisyuName2]              CHAR(34)  WITH COMPRESSION,"
    strSQL = strSQL & "[RecUmaKettoNum3]               CHAR(10)  WITH COMPRESSION,"
    strSQL = strSQL & "[RecUmaBamei3]                  CHAR(36)  WITH COMPRESSION,"
    strSQL = strSQL & "[RecUmaUmaKigoCD3]              CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[RecUmaSexCD3]                  CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[RecUmaChokyosiCode3]           CHAR(5)   WITH COMPRESSION,"
    strSQL = strSQL & "[RecUmaChokyosiName3]           CHAR(34)  WITH COMPRESSION,"
    strSQL = strSQL & "[RecUmaFutan3]                  CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[RecUmaKisyuCode3]              CHAR(5)   WITH COMPRESSION,"
    strSQL = strSQL & "[RecUmaKisyuName3]              CHAR(34)  WITH COMPRESSION"
    strSQL = strSQL & ")"

    mConnection.Execute strSQL

    createTable_RECORD = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createTable_RECORD = False
End Function ' createTable_RECORD


'
'   機能: DBにHANRO(坂路調教)テーブルを作る
'
'   備考: なし
'
Public Function createTable_HANRO() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String

    strSQL = "CREATE TABLE HANRO("
    strSQL = strSQL & "[RecordSpec]                    CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[DataKubun]                     CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[MakeDate]                      CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[TresenKubun]                   CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[ChokyoDate]                    CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[ChokyoTime]                    CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[KettoNum]                      CHAR(10)  WITH COMPRESSION,"
    strSQL = strSQL & "[HaronTime4]                    CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[LapTime4]                      CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[HaronTime3]                    CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[LapTime3]                      CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[HaronTime2]                    CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[LapTime2]                      CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[LapTime1]                      CHAR(3)   WITH COMPRESSION"
    strSQL = strSQL & ")"

    mConnection.Execute strSQL

    createTable_HANRO = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createTable_HANRO = False
End Function ' createTable_HANRO


'
'   機能: DBにBATAIJYU(馬体重)テーブルを作る
'
'   備考: なし
'
Public Function createTable_BATAIJYU() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String

    strSQL = "CREATE TABLE BATAIJYU("
    strSQL = strSQL & "[RecordSpec]                    CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[DataKubun]                     CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[MakeDate]                      CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[Year]                          CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[MonthDay]                      CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[JyoCD]                         CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kaiji]                         CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Nichiji]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[RaceNum]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[HappyoTime]                    CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[Umaban1]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Bamei1]                        CHAR(36)  WITH COMPRESSION,"
    strSQL = strSQL & "[BaTaijyu1]                     CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[ZogenFugo1]                    CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[ZogenSa1]                      CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Umaban2]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Bamei2]                        CHAR(36)  WITH COMPRESSION,"
    strSQL = strSQL & "[BaTaijyu2]                     CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[ZogenFugo2]                    CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[ZogenSa2]                      CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Umaban3]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Bamei3]                        CHAR(36)  WITH COMPRESSION,"
    strSQL = strSQL & "[BaTaijyu3]                     CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[ZogenFugo3]                    CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[ZogenSa3]                      CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Umaban4]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Bamei4]                        CHAR(36)  WITH COMPRESSION,"
    strSQL = strSQL & "[BaTaijyu4]                     CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[ZogenFugo4]                    CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[ZogenSa4]                      CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Umaban5]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Bamei5]                        CHAR(36)  WITH COMPRESSION,"
    strSQL = strSQL & "[BaTaijyu5]                     CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[ZogenFugo5]                    CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[ZogenSa5]                      CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Umaban6]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Bamei6]                        CHAR(36)  WITH COMPRESSION,"
    strSQL = strSQL & "[BaTaijyu6]                     CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[ZogenFugo6]                    CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[ZogenSa6]                      CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Umaban7]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Bamei7]                        CHAR(36)  WITH COMPRESSION,"
    strSQL = strSQL & "[BaTaijyu7]                     CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[ZogenFugo7]                    CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[ZogenSa7]                      CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Umaban8]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Bamei8]                        CHAR(36)  WITH COMPRESSION,"
    strSQL = strSQL & "[BaTaijyu8]                     CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[ZogenFugo8]                    CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[ZogenSa8]                      CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Umaban9]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Bamei9]                        CHAR(36)  WITH COMPRESSION,"
    strSQL = strSQL & "[BaTaijyu9]                     CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[ZogenFugo9]                    CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[ZogenSa9]                      CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Umaban10]                      CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Bamei10]                       CHAR(36)  WITH COMPRESSION,"
    strSQL = strSQL & "[BaTaijyu10]                    CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[ZogenFugo10]                   CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[ZogenSa10]                     CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Umaban11]                      CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Bamei11]                       CHAR(36)  WITH COMPRESSION,"
    strSQL = strSQL & "[BaTaijyu11]                    CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[ZogenFugo11]                   CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[ZogenSa11]                     CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Umaban12]                      CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Bamei12]                       CHAR(36)  WITH COMPRESSION,"
    strSQL = strSQL & "[BaTaijyu12]                    CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[ZogenFugo12]                   CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[ZogenSa12]                     CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Umaban13]                      CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Bamei13]                       CHAR(36)  WITH COMPRESSION,"
    strSQL = strSQL & "[BaTaijyu13]                    CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[ZogenFugo13]                   CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[ZogenSa13]                     CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Umaban14]                      CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Bamei14]                       CHAR(36)  WITH COMPRESSION,"
    strSQL = strSQL & "[BaTaijyu14]                    CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[ZogenFugo14]                   CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[ZogenSa14]                     CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Umaban15]                      CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Bamei15]                       CHAR(36)  WITH COMPRESSION,"
    strSQL = strSQL & "[BaTaijyu15]                    CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[ZogenFugo15]                   CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[ZogenSa15]                     CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Umaban16]                      CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Bamei16]                       CHAR(36)  WITH COMPRESSION,"
    strSQL = strSQL & "[BaTaijyu16]                    CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[ZogenFugo16]                   CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[ZogenSa16]                     CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Umaban17]                      CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Bamei17]                       CHAR(36)  WITH COMPRESSION,"
    strSQL = strSQL & "[BaTaijyu17]                    CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[ZogenFugo17]                   CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[ZogenSa17]                     CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Umaban18]                      CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Bamei18]                       CHAR(36)  WITH COMPRESSION,"
    strSQL = strSQL & "[BaTaijyu18]                    CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[ZogenFugo18]                   CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[ZogenSa18]                     CHAR(3)   WITH COMPRESSION"
    strSQL = strSQL & ")"

    mConnection.Execute strSQL

    createTable_BATAIJYU = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createTable_BATAIJYU = False
End Function ' createTable_BATAIJYU


'
'   機能: DBにTENKO_BABA(天候馬場状態)テーブルを作る
'
'   備考: なし
'
Public Function createTable_TENKO_BABA() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String

    strSQL = "CREATE TABLE TENKO_BABA("
    strSQL = strSQL & "[RecordSpec]                    CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[DataKubun]                     CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[MakeDate]                      CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[Year]                          CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[MonthDay]                      CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[JyoCD]                         CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kaiji]                         CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Nichiji]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[HappyoTime]                    CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[HenkoID]                       CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[AtoTenkoCD]                    CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[AtoSibaBabaCD]                 CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[AtoDirtBabaCD]                 CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[MaeTenkoCD]                    CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[MaeSibaBabaCD]                 CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[MaeDirtBabaCD]                 CHAR(1)   WITH COMPRESSION"
    strSQL = strSQL & ")"

    mConnection.Execute strSQL

    createTable_TENKO_BABA = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createTable_TENKO_BABA = False
End Function ' createTable_TENKO_BABA


'
'   機能: DBにTORIKESI_JYOGAI(出走取消・競走除外)テーブルを作る
'
'   備考: なし
'
Public Function createTable_TORIKESI_JYOGAI() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String

    strSQL = "CREATE TABLE TORIKESI_JYOGAI("
    strSQL = strSQL & "[RecordSpec]                    CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[DataKubun]                     CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[MakeDate]                      CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[Year]                          CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[MonthDay]                      CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[JyoCD]                         CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kaiji]                         CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Nichiji]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[RaceNum]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[HappyoTime]                    CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[Umaban]                        CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Bamei]                         CHAR(36)  WITH COMPRESSION,"
    strSQL = strSQL & "[JiyuKubun]                     CHAR(3)   WITH COMPRESSION"
    strSQL = strSQL & ")"

    mConnection.Execute strSQL

    createTable_TORIKESI_JYOGAI = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createTable_TORIKESI_JYOGAI = False
End Function ' createTable_TORIKESI_JYOGAI


'
'   機能: DBにKISHU_CHANGE(騎手変更)テーブルを作る
'
'   備考: なし
'
Public Function createTable_KISHU_CHANGE() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String

    strSQL = "CREATE TABLE KISHU_CHANGE("
    strSQL = strSQL & "[RecordSpec]                    CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[DataKubun]                     CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[MakeDate]                      CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[Year]                          CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[MonthDay]                      CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[JyoCD]                         CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kaiji]                         CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Nichiji]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[RaceNum]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[HappyoTime]                    CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[Umaban]                        CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Bamei]                         CHAR(36)  WITH COMPRESSION,"
    strSQL = strSQL & "[AtoFutan]                      CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[AtoKisyuCode]                  CHAR(5)   WITH COMPRESSION,"
    strSQL = strSQL & "[AtoKisyuName]                  CHAR(34)  WITH COMPRESSION,"
    strSQL = strSQL & "[AtoMinaraiCD]                  CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[MaeFutan]                      CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[MaeKisyuCode]                  CHAR(5)   WITH COMPRESSION,"
    strSQL = strSQL & "[MaeKisyuName]                  CHAR(34)  WITH COMPRESSION,"
    strSQL = strSQL & "[MaeMinaraiCD]                  CHAR(1)   WITH COMPRESSION"
    strSQL = strSQL & ")"

    mConnection.Execute strSQL

    createTable_KISHU_CHANGE = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createTable_KISHU_CHANGE = False
End Function ' createTable_KISHU_CHANGE


'
'   機能: DBにMINING(データマイニング予想)テーブルを作る
'
'   備考: なし
'
Public Function createTable_MINING() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String

    strSQL = "CREATE TABLE MINING("
    strSQL = strSQL & "[RecordSpec]                    CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[DataKubun]                     CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[MakeDate]                      CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[Year]                          CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[MonthDay]                      CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[JyoCD]                         CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kaiji]                         CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Nichiji]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[RaceNum]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[MakeHM]                        CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[Umaban1]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMTime1]                       CHAR(5)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMGosaP1]                      CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMGosaM1]                      CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[Umaban2]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMTime2]                       CHAR(5)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMGosaP2]                      CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMGosaM2]                      CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[Umaban3]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMTime3]                       CHAR(5)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMGosaP3]                      CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMGosaM3]                      CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[Umaban4]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMTime4]                       CHAR(5)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMGosaP4]                      CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMGosaM4]                      CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[Umaban5]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMTime5]                       CHAR(5)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMGosaP5]                      CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMGosaM5]                      CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[Umaban6]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMTime6]                       CHAR(5)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMGosaP6]                      CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMGosaM6]                      CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[Umaban7]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMTime7]                       CHAR(5)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMGosaP7]                      CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMGosaM7]                      CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[Umaban8]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMTime8]                       CHAR(5)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMGosaP8]                      CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMGosaM8]                      CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[Umaban9]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMTime9]                       CHAR(5)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMGosaP9]                      CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMGosaM9]                      CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[Umaban10]                      CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMTime10]                      CHAR(5)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMGosaP10]                     CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMGosaM10]                     CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[Umaban11]                      CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMTime11]                      CHAR(5)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMGosaP11]                     CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMGosaM11]                     CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[Umaban12]                      CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMTime12]                      CHAR(5)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMGosaP12]                     CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMGosaM12]                     CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[Umaban13]                      CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMTime13]                      CHAR(5)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMGosaP13]                     CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMGosaM13]                     CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[Umaban14]                      CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMTime14]                      CHAR(5)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMGosaP14]                     CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMGosaM14]                     CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[Umaban15]                      CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMTime15]                      CHAR(5)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMGosaP15]                     CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMGosaM15]                     CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[Umaban16]                      CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMTime16]                      CHAR(5)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMGosaP16]                     CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMGosaM16]                     CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[Umaban17]                      CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMTime17]                      CHAR(5)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMGosaP17]                     CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMGosaM17]                     CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[Umaban18]                      CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMTime18]                      CHAR(5)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMGosaP18]                     CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[DMGosaM18]                     CHAR(4)   WITH COMPRESSION"
    strSQL = strSQL & ")"

    mConnection.Execute strSQL

    createTable_MINING = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createTable_MINING = False
End Function ' createTable_MINING


'
'   機能: DBにSCHEDULE(開催スケジュール)テーブルを作る
'
'   備考: なし
'
Public Function createTable_SCHEDULE() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String

    strSQL = "CREATE TABLE SCHEDULE("
    strSQL = strSQL & "[RecordSpec]                    CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[DataKubun]                     CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[MakeDate]                      CHAR(8)   WITH COMPRESSION,"
    strSQL = strSQL & "[Year]                          CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[MonthDay]                      CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[JyoCD]                         CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Kaiji]                         CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Nichiji]                       CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[YoubiCD]                       CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyusyo1TokuNum]                CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyusyo1Hondai]                 CHAR(60)  WITH COMPRESSION,"
    strSQL = strSQL & "[Jyusyo1Ryakusyo10]             CHAR(20)  WITH COMPRESSION,"
    strSQL = strSQL & "[Jyusyo1Ryakusyo6]              CHAR(12)  WITH COMPRESSION,"
    strSQL = strSQL & "[Jyusyo1Ryakusyo3]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyusyo1Nkai]                   CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyusyo1GradeCD]                CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyusyo1SyubetuCD]              CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyusyo1KigoCD]                 CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyusyo1JyuryoCD]               CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyusyo1Kyori]                  CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyusyo1TrackCD]                CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyusyo2TokuNum]                CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyusyo2Hondai]                 CHAR(60)  WITH COMPRESSION,"
    strSQL = strSQL & "[Jyusyo2Ryakusyo10]             CHAR(20)  WITH COMPRESSION,"
    strSQL = strSQL & "[Jyusyo2Ryakusyo6]              CHAR(12)  WITH COMPRESSION,"
    strSQL = strSQL & "[Jyusyo2Ryakusyo3]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyusyo2Nkai]                   CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyusyo2GradeCD]                CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyusyo2SyubetuCD]              CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyusyo2KigoCD]                 CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyusyo2JyuryoCD]               CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyusyo2Kyori]                  CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyusyo2TrackCD]                CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyusyo3TokuNum]                CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyusyo3Hondai]                 CHAR(60)  WITH COMPRESSION,"
    strSQL = strSQL & "[Jyusyo3Ryakusyo10]             CHAR(20)  WITH COMPRESSION,"
    strSQL = strSQL & "[Jyusyo3Ryakusyo6]              CHAR(12)  WITH COMPRESSION,"
    strSQL = strSQL & "[Jyusyo3Ryakusyo3]              CHAR(6)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyusyo3Nkai]                   CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyusyo3GradeCD]                CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyusyo3SyubetuCD]              CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyusyo3KigoCD]                 CHAR(3)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyusyo3JyuryoCD]               CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyusyo3Kyori]                  CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[Jyusyo3TrackCD]                CHAR(2)   WITH COMPRESSION"
    strSQL = strSQL & ")"

    mConnection.Execute strSQL

    createTable_SCHEDULE = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createTable_SCHEDULE = False
End Function ' createTable_SCHEDULE


'
'   機能: DBにRAKaiSel(出馬表開催選択)テーブルを作る
'
'   備考: なし
'
Public Function createTable_RAKaiSel() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String

    strSQL = "CREATE TABLE RAKaiSel("
    strSQL = strSQL & "[Year]                          CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[MonthDay]                      CHAR(4)   WITH COMPRESSION,"
    strSQL = strSQL & "[JyoCD]                         CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[YoubiCD]                       CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[CanLink]                       CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[JyoCD0]                        CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[JyoCD1]                        CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[JyoCD2]                        CHAR(2)   WITH COMPRESSION,"
    strSQL = strSQL & "[Dai0]                          CHAR(12)  WITH COMPRESSION,"
    strSQL = strSQL & "[Dai1]                          CHAR(12)  WITH COMPRESSION,"
    strSQL = strSQL & "[Dai2]                          CHAR(12)  WITH COMPRESSION,"
    strSQL = strSQL & "[DaiToolTip0]                   CHAR(60)  WITH COMPRESSION,"
    strSQL = strSQL & "[DaiToolTip1]                   CHAR(60)  WITH COMPRESSION,"
    strSQL = strSQL & "[DaiToolTip2]                   CHAR(60)  WITH COMPRESSION,"
    strSQL = strSQL & "[GradeCD0]                      CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[GradeCD1]                      CHAR(1)   WITH COMPRESSION,"
    strSQL = strSQL & "[GradeCD2]                      CHAR(1)   WITH COMPRESSION"
    strSQL = strSQL & ")"

    mConnection.Execute strSQL

    createTable_RAKaiSel = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createTable_RAKaiSel = False
End Function ' createTable_RAKaiSel


'
'   機能: DBにODDS_SANRENTAN(オッズ3連単)テーブルを作る
'
'   備考: なし
'
Public Function createTable_ODDS_SANRENTAN() As Boolean
On Error GoTo Errorhandler

    Dim strSQL As String
    
    strSQL = "CREATE TABLE ODDS_SANRENTAN("
    strSQL = strSQL & "[RecordSpec]                     CHAR(2)     WITH COMPRESSION,"
    strSQL = strSQL & "[DataKubun]                      CHAR(1)     WITH COMPRESSION,"
    strSQL = strSQL & "[MakeDate]                       CHAR(8)     WITH COMPRESSION,"
    strSQL = strSQL & "[Year]                           CHAR(4)     WITH COMPRESSION,"
    strSQL = strSQL & "[MonthDay]                       CHAR(4)     WITH COMPRESSION,"
    strSQL = strSQL & "[JyoCD]                          CHAR(2)     WITH COMPRESSION,"
    strSQL = strSQL & "[Kaiji]                          CHAR(2)     WITH COMPRESSION,"
    strSQL = strSQL & "[Nichiji]                        CHAR(2)     WITH COMPRESSION,"
    strSQL = strSQL & "[RaceNum]                        CHAR(2)     WITH COMPRESSION,"
    strSQL = strSQL & "[HappyoTime]                     CHAR(8)     WITH COMPRESSION,"
    strSQL = strSQL & "[TorokuTosu]                     CHAR(2)     WITH COMPRESSION,"
    strSQL = strSQL & "[SyussoTosu]                     CHAR(2)     WITH COMPRESSION,"
    strSQL = strSQL & "[SanrentanFlag]                  CHAR(1)     WITH COMPRESSION,"
    strSQL = strSQL & "[TotalHyosuSanrentan]            CHAR(11)    WITH COMPRESSION,"
    strSQL = strSQL & "[Data]                           TEXT        WITH COMPRESSION"
    strSQL = strSQL & ")"
    
    mConnection.Execute strSQL
    
    createTable_ODDS_SANRENTAN = True
    Exit Function

Errorhandler:
    gApp.ErrLog
    createTable_ODDS_SANRENTAN = False
End Function    'createTable_ODDS_SANRENTAN


'
'   機能: DBにHASSOU_CHANGE(発走時刻変更)テーブルを作る
'
'   備考: なし
'
Public Function createTable_HASSOU_CHANGE() As Boolean
On Error GoTo Errorhandler

    Dim strSQL As String
    
    strSQL = "CREATE TABLE HASSOU_CHANGE("
    strSQL = strSQL & "[RecordSpec]                     CHAR(2)     WITH COMPRESSION,"
    strSQL = strSQL & "[DataKubun]                     CHAR(1)     WITH COMPRESSION,"
    strSQL = strSQL & "[MakeDate]                      CHAR(8)     WITH COMPRESSION,"
    strSQL = strSQL & "[Year]                          CHAR(4)     WITH COMPRESSION,"
    strSQL = strSQL & "[MonthDay]                      CHAR(4)     WITH COMPRESSION,"
    strSQL = strSQL & "[JyoCD]                         CHAR(2)     WITH COMPRESSION,"
    strSQL = strSQL & "[Kaiji]                         CHAR(2)     WITH COMPRESSION,"
    strSQL = strSQL & "[Nichiji]                       CHAR(2)     WITH COMPRESSION,"
    strSQL = strSQL & "[RaceNum]                       CHAR(2)     WITH COMPRESSION,"
    strSQL = strSQL & "[HappyoTime]                    CHAR(8)     WITH COMPRESSION,"
    strSQL = strSQL & "[AtoHassoTime]                  CHAR(4)     WITH COMPRESSION,"
    strSQL = strSQL & "[MaeHassoTime]                  CHAR(4)     WITH COMPRESSION"
    strSQL = strSQL & ")"

    mConnection.Execute strSQL
    
    createTable_HASSOU_CHANGE = True
    Exit Function

Errorhandler:
    gApp.ErrLog
    createTable_HASSOU_CHANGE = False
End Function    'createTable_HASSOU_CHANGE


'
'   機能: DBにCOURSE_CHANGE(コース変更)テーブルを作る
'
'   備考: なし
'
Public Function createTable_COURSE_CHANGE() As Boolean
On Error GoTo Errorhandler

    Dim strSQL As String
    
    strSQL = "CREATE TABLE COURSE_CHANGE("
    strSQL = strSQL & "[RecordSpec]                    CHAR(2)     WITH COMPRESSION,"
    strSQL = strSQL & "[DataKubun]                     CHAR(1)     WITH COMPRESSION,"
    strSQL = strSQL & "[MakeDate]                      CHAR(8)     WITH COMPRESSION,"
    strSQL = strSQL & "[Year]                          CHAR(4)     WITH COMPRESSION,"
    strSQL = strSQL & "[MonthDay]                      CHAR(4)     WITH COMPRESSION,"
    strSQL = strSQL & "[JyoCD]                         CHAR(2)     WITH COMPRESSION,"
    strSQL = strSQL & "[Kaiji]                         CHAR(2)     WITH COMPRESSION,"
    strSQL = strSQL & "[Nichiji]                       CHAR(2)     WITH COMPRESSION,"
    strSQL = strSQL & "[RaceNum]                       CHAR(2)     WITH COMPRESSION,"
    strSQL = strSQL & "[HappyoTime]                    CHAR(8)     WITH COMPRESSION,"
    strSQL = strSQL & "[AtoKyori]                      CHAR(4)     WITH COMPRESSION,"
    strSQL = strSQL & "[AtoTrackCD]                    CHAR(2)     WITH COMPRESSION,"
    strSQL = strSQL & "[MaeKyori]                      CHAR(4)     WITH COMPRESSION,"
    strSQL = strSQL & "[MaeTrackCD]                    CHAR(2)     WITH COMPRESSION,"
    strSQL = strSQL & "[JiyuKubun]                     CHAR(1)     WITH COMPRESSION"
    strSQL = strSQL & ")"
    
    mConnection.Execute strSQL
    
    createTable_COURSE_CHANGE = True
    Exit Function

Errorhandler:
    gApp.ErrLog
    createTable_COURSE_CHANGE = False
End Function    'createTable_COURSE_CHANGE



'
'   機能: DBにTOKU_RACE(特別レース)テーブルを作る
'
'   備考: なし
'
Public Function createIndex_TOKU_RACE() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String

    strSQL = "CREATE INDEX PrimaryKey ON TOKU_RACE ("
    strSQL = strSQL & "[Year],"
    strSQL = strSQL & "[MonthDay],"
    strSQL = strSQL & "[JyoCD],"
    strSQL = strSQL & "[Kaiji],"
    strSQL = strSQL & "[Nichiji],"
    strSQL = strSQL & "[RaceNum]"
    strSQL = strSQL & ") WITH PRIMARY"

    mConnection.Execute strSQL

    createIndex_TOKU_RACE = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createIndex_TOKU_RACE = False
End Function ' createIndex_TOKU_RACE


'
'   機能: DBにTOKU(特別登録馬)テーブルを作る
'
'   備考: なし
'
Public Function createIndex_TOKU() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String

    strSQL = "CREATE INDEX PrimaryKey ON TOKU ("
    strSQL = strSQL & "[Year],"
    strSQL = strSQL & "[MonthDay],"
    strSQL = strSQL & "[JyoCD],"
    strSQL = strSQL & "[Kaiji],"
    strSQL = strSQL & "[Nichiji],"
    strSQL = strSQL & "[RaceNum],"
    strSQL = strSQL & "[Num]"
    strSQL = strSQL & ") WITH PRIMARY"

    mConnection.Execute strSQL

    createIndex_TOKU = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createIndex_TOKU = False
End Function ' createIndex_TOKU


'
'   機能: DBにRACE(レース詳細)テーブルを作る
'
'   備考: なし
'
Public Function createIndex_RACE() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String

    strSQL = "CREATE INDEX PrimaryKey ON RACE ("
    strSQL = strSQL & "[Year],"
    strSQL = strSQL & "[MonthDay],"
    strSQL = strSQL & "[JyoCD],"
    strSQL = strSQL & "[Kaiji],"
    strSQL = strSQL & "[Nichiji],"
    strSQL = strSQL & "[RaceNum]"
    strSQL = strSQL & ") WITH PRIMARY"

    mConnection.Execute strSQL


    strSQL = "CREATE INDEX Kyori ON RACE ( [Kyori] )"
    mConnection.Execute strSQL
    
    createIndex_RACE = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createIndex_RACE = False
End Function ' createIndex_RACE


'
'   機能: DBにUMA_RACE_A(馬毎レース情報_前半)テーブルを作る
'
'   備考: なし
'
Public Function createIndex_UMA_RACE_A() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String



    strSQL = "CREATE INDEX PrimaryKey ON UMA_RACE_A ("
    strSQL = strSQL & "[Year],"
    strSQL = strSQL & "[MonthDay],"
    strSQL = strSQL & "[JyoCD],"
    strSQL = strSQL & "[Kaiji],"
    strSQL = strSQL & "[Nichiji],"
    strSQL = strSQL & "[RaceNum],"
    strSQL = strSQL & "[Umaban],"
    strSQL = strSQL & "[KettoNum]"
    strSQL = strSQL & ") WITH PRIMARY"

    mConnection.Execute strSQL

    strSQL = "CREATE INDEX KettoNum ON UMA_RACE_A ( [KettoNum] )"
    mConnection.Execute strSQL

    createIndex_UMA_RACE_A = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createIndex_UMA_RACE_A = False
End Function ' createIndex_UMA_RACE_A


'
'   機能: DBにUMA_RACE_B(馬毎レース情報_後半)テーブルを作る
'
'   備考: なし
'
Public Function createIndex_UMA_RACE_B() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String



    strSQL = "CREATE INDEX PrimaryKey ON UMA_RACE_B ("
    strSQL = strSQL & "[B_Year],"
    strSQL = strSQL & "[B_MonthDay],"
    strSQL = strSQL & "[B_JyoCD],"
    strSQL = strSQL & "[B_RaceNum],"
    strSQL = strSQL & "[B_Umaban],"
    strSQL = strSQL & "[B_KettoNum]"
    strSQL = strSQL & ") WITH PRIMARY"

    mConnection.Execute strSQL
    
    strSQL = "CREATE INDEX [Time] ON UMA_RACE_B ( [Time] )"
    mConnection.Execute strSQL

    createIndex_UMA_RACE_B = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createIndex_UMA_RACE_B = False
End Function ' createIndex_UMA_RACE_B


'
'   機能: DBにHARAI(払戻)テーブルを作る
'
'   備考: なし
'
Public Function createIndex_HARAI() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String

    strSQL = "CREATE INDEX PrimaryKey ON HARAI ("
    strSQL = strSQL & "[Year],"
    strSQL = strSQL & "[MonthDay],"
    strSQL = strSQL & "[JyoCD],"
    strSQL = strSQL & "[Kaiji],"
    strSQL = strSQL & "[Nichiji],"
    strSQL = strSQL & "[RaceNum]"
    strSQL = strSQL & ") WITH PRIMARY"

    mConnection.Execute strSQL

    createIndex_HARAI = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createIndex_HARAI = False
End Function ' createIndex_HARAI


'
'   機能: DBにODDS_TANPUKUWAKU(オッズ_単複枠)テーブルを作る
'
'   備考: なし
'
Public Function createIndex_ODDS_TANPUKUWAKU() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String

    strSQL = "CREATE INDEX PrimaryKey ON ODDS_TANPUKUWAKU ("
    strSQL = strSQL & "[Year],"
    strSQL = strSQL & "[MonthDay],"
    strSQL = strSQL & "[JyoCD],"
    strSQL = strSQL & "[Kaiji],"
    strSQL = strSQL & "[Nichiji],"
    strSQL = strSQL & "[RaceNum]"
    strSQL = strSQL & ") WITH PRIMARY"

    mConnection.Execute strSQL

    createIndex_ODDS_TANPUKUWAKU = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createIndex_ODDS_TANPUKUWAKU = False
End Function ' createIndex_ODDS_TANPUKUWAKU


'
'   機能: DBにODDS_UMAREN(オッズ_馬連)テーブルを作る
'
'   備考: なし
'
Public Function createIndex_ODDS_UMAREN() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String

    strSQL = "CREATE INDEX PrimaryKey ON ODDS_UMAREN ("
    strSQL = strSQL & "[Year],"
    strSQL = strSQL & "[MonthDay],"
    strSQL = strSQL & "[JyoCD],"
    strSQL = strSQL & "[Kaiji],"
    strSQL = strSQL & "[Nichiji],"
    strSQL = strSQL & "[RaceNum]"
    strSQL = strSQL & ") WITH PRIMARY"

    mConnection.Execute strSQL

    createIndex_ODDS_UMAREN = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createIndex_ODDS_UMAREN = False
End Function ' createIndex_ODDS_UMAREN


'
'   機能: DBにODDS_WIDE(オッズ_ワイド)テーブルを作る
'
'   備考: なし
'
Public Function createIndex_ODDS_WIDE() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String

    strSQL = "CREATE INDEX PrimaryKey ON ODDS_WIDE ("
    strSQL = strSQL & "[Year],"
    strSQL = strSQL & "[MonthDay],"
    strSQL = strSQL & "[JyoCD],"
    strSQL = strSQL & "[Kaiji],"
    strSQL = strSQL & "[Nichiji],"
    strSQL = strSQL & "[RaceNum]"
    strSQL = strSQL & ") WITH PRIMARY"

    mConnection.Execute strSQL

    createIndex_ODDS_WIDE = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createIndex_ODDS_WIDE = False
End Function ' createIndex_ODDS_WIDE


'
'   機能: DBにODDS_UMATAN(オッズ_馬単)テーブルを作る
'
'   備考: なし
'
Public Function createIndex_ODDS_UMATAN() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String


    strSQL = "CREATE INDEX PrimaryKey ON ODDS_UMATAN ("
    strSQL = strSQL & "[Year],"
    strSQL = strSQL & "[MonthDay],"
    strSQL = strSQL & "[JyoCD],"
    strSQL = strSQL & "[Kaiji],"
    strSQL = strSQL & "[Nichiji],"
    strSQL = strSQL & "[RaceNum]"
    strSQL = strSQL & ") WITH PRIMARY"

    mConnection.Execute strSQL

    createIndex_ODDS_UMATAN = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createIndex_ODDS_UMATAN = False
End Function ' createIndex_ODDS_UMATAN


'
'   機能: DBにODDS_SANREN(オッズ_3連複)テーブルを作る
'
'   備考: なし
'
Public Function createIndex_ODDS_SANREN() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String

    strSQL = "CREATE INDEX PrimaryKey ON ODDS_SANREN ("
    strSQL = strSQL & "[Year],"
    strSQL = strSQL & "[MonthDay],"
    strSQL = strSQL & "[JyoCD],"
    strSQL = strSQL & "[Kaiji],"
    strSQL = strSQL & "[Nichiji],"
    strSQL = strSQL & "[RaceNum]"
    strSQL = strSQL & ") WITH PRIMARY"

    mConnection.Execute strSQL

    createIndex_ODDS_SANREN = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createIndex_ODDS_SANREN = False
End Function ' createIndex_ODDS_SANREN


'
'   機能: DBにUMA(競走馬マスタ)テーブルを作る
'
'   備考: なし
'
Public Function createIndex_UMA() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String

    strSQL = "CREATE INDEX PrimaryKey ON UMA ("
    strSQL = strSQL & "[KettoNum]"
    strSQL = strSQL & ") WITH PRIMARY"

    mConnection.Execute strSQL
    
    strSQL = "CREATE INDEX Bamei ON UMA ( [Bamei] )"
    mConnection.Execute strSQL
    
    strSQL = "CREATE INDEX BameiKana ON UMA ( [BameiKana] )"
    mConnection.Execute strSQL
    
    strSQL = "CREATE INDEX BameiEng ON UMA ( [BameiEng] )"
    mConnection.Execute strSQL
    
    strSQL = "CREATE INDEX ChokyosiCode ON UMA ( [ChokyosiCode] )"
    mConnection.Execute strSQL
    
    strSQL = "CREATE INDEX BreederCode ON UMA ( [BreederCode] )"
    mConnection.Execute strSQL

    strSQL = "CREATE INDEX BanusiCode ON UMA ( [BanusiCode] )"
    mConnection.Execute strSQL

    createIndex_UMA = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createIndex_UMA = False
End Function ' createIndex_UMA


'
'   機能: DBにKISHU(騎手マスタ)テーブルを作る
'
'   備考: なし
'
Public Function createIndex_KISHU() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String

    strSQL = "CREATE INDEX PrimaryKey ON KISHU ("
    strSQL = strSQL & "[KisyuCode]"
    strSQL = strSQL & ") WITH PRIMARY"

    mConnection.Execute strSQL

    strSQL = "CREATE INDEX KisyuName ON KISHU ( [KisyuName] )"
    mConnection.Execute strSQL
    strSQL = "CREATE INDEX KisyuNameKana ON KISHU ( [KisyuNameKana] )"
    mConnection.Execute strSQL
    strSQL = "CREATE INDEX KisyuRyakusyo ON KISHU ( [KisyuRyakusyo] )"
    mConnection.Execute strSQL
    strSQL = "CREATE INDEX KisyuNameEng ON KISHU ( [KisyuNameEng] )"
    mConnection.Execute strSQL

    createIndex_KISHU = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createIndex_KISHU = False
End Function ' createIndex_KISHU


'
'   機能: DBにKISHU_SEISEKI(騎手マスタ_成績)テーブルを作る
'
'   備考: なし
'
Public Function createIndex_KISHU_SEISEKI() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String


    strSQL = "CREATE INDEX PrimaryKey ON KISHU_SEISEKI ("
    strSQL = strSQL & "[KisyuCode],"
    strSQL = strSQL & "[Num]"
    strSQL = strSQL & ") WITH PRIMARY"

    mConnection.Execute strSQL

    createIndex_KISHU_SEISEKI = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createIndex_KISHU_SEISEKI = False
End Function ' createIndex_KISHU_SEISEKI


'
'   機能: DBにCHOKYO(調教師マスタ)テーブルを作る
'
'   備考: なし
'
Public Function createIndex_CHOKYO() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String

    strSQL = "CREATE INDEX PrimaryKey ON CHOKYO ("
    strSQL = strSQL & "[ChokyosiCode]"
    strSQL = strSQL & ") WITH PRIMARY"
    mConnection.Execute strSQL
    
    strSQL = "CREATE INDEX ChokyosiNameKey ON CHOKYO ("
    strSQL = strSQL & "[ChokyosiName]"
    strSQL = strSQL & ") "
    mConnection.Execute strSQL

    strSQL = "CREATE INDEX ChokyosiNameKanaKey ON CHOKYO ("
    strSQL = strSQL & "[ChokyosiNameKana]"
    strSQL = strSQL & ") "
    mConnection.Execute strSQL

    strSQL = "CREATE INDEX ChokyosiRyakusyoKey ON CHOKYO ("
    strSQL = strSQL & "[ChokyosiRyakusyo]"
    strSQL = strSQL & ") "
    mConnection.Execute strSQL

    strSQL = "CREATE INDEX ChokyosiNameEngKey ON CHOKYO ("
    strSQL = strSQL & "[ChokyosiNameEng]"
    strSQL = strSQL & ") "
    mConnection.Execute strSQL

    createIndex_CHOKYO = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createIndex_CHOKYO = False
End Function ' createIndex_CHOKYO


'
'   機能: DBにCHOKYO_SEISEKI(調教師マスタ_成績)テーブルを作る
'
'   備考: なし
'
Public Function createIndex_CHOKYO_SEISEKI() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String


    strSQL = "CREATE INDEX PrimaryKey ON CHOKYO_SEISEKI ("
    strSQL = strSQL & "[ChokyosiCode],"
    strSQL = strSQL & "[Num]"
    strSQL = strSQL & ") WITH PRIMARY"

    mConnection.Execute strSQL

    createIndex_CHOKYO_SEISEKI = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createIndex_CHOKYO_SEISEKI = False
End Function ' createIndex_CHOKYO_SEISEKI


'
'   機能: DBにSEISAN(生産者マスタ)テーブルを作る
'
'   備考: なし
'
Public Function createIndex_SEISAN() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String

    strSQL = "CREATE INDEX PrimaryKey ON SEISAN ("
    strSQL = strSQL & "[BreederCode]"
    strSQL = strSQL & ") WITH PRIMARY"

    mConnection.Execute strSQL

    strSQL = "CREATE INDEX BreederName_CoKey ON SEISAN ([BreederName_Co]) "
    mConnection.Execute strSQL
    
    strSQL = "CREATE INDEX BreederNameKey ON SEISAN ([BreederName]) "
    mConnection.Execute strSQL
    
    strSQL = "CREATE INDEX BreederNameKanaKey ON SEISAN ([BreederNameKana]) "
    mConnection.Execute strSQL
    
    strSQL = "CREATE INDEX BreederNameEngKey ON SEISAN ([BreederNameEng]) "
    mConnection.Execute strSQL
    
    createIndex_SEISAN = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createIndex_SEISAN = False
End Function ' createIndex_SEISAN


'
'   機能: DBにBANUSI(馬主マスタ)テーブルを作る
'
'   備考: なし
'
Public Function createIndex_BANUSI() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String

    strSQL = "CREATE INDEX PrimaryKey ON BANUSI ("
    strSQL = strSQL & "[BanusiCode]"
    strSQL = strSQL & ") WITH PRIMARY"
    mConnection.Execute strSQL
    
    strSQL = "CREATE INDEX BanusiName_CoKey ON BANUSI ("
    strSQL = strSQL & "[BanusiName_Co]"
    strSQL = strSQL & ") "
    mConnection.Execute strSQL
    
    strSQL = "CREATE INDEX BanusiNameKey ON BANUSI ("
    strSQL = strSQL & "[BanusiName]"
    strSQL = strSQL & ") "
    mConnection.Execute strSQL
    
    strSQL = "CREATE INDEX BanusiNameKanaKey ON BANUSI ("
    strSQL = strSQL & "[BanusiNameKana]"
    strSQL = strSQL & ") "
    mConnection.Execute strSQL
    
    strSQL = "CREATE INDEX BanusiNameEngKey ON BANUSI ("
    strSQL = strSQL & "[BanusiNameEng]"
    strSQL = strSQL & ") "
    
    mConnection.Execute strSQL

    createIndex_BANUSI = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createIndex_BANUSI = False
End Function ' createIndex_BANUSI


'
'   機能: DBにHANSYOKU(繁殖馬マスタ)テーブルを作る
'
'   備考: なし
'
Public Function createIndex_HANSYOKU() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String

    strSQL = "CREATE INDEX PrimaryKey ON HANSYOKU ("
    strSQL = strSQL & "[HansyokuNum]"
    strSQL = strSQL & ") WITH PRIMARY"

    mConnection.Execute strSQL
    
    strSQL = "CREATE INDEX BameiKey ON HANSYOKU ("
    strSQL = strSQL & "[Bamei]"
    strSQL = strSQL & ") "
    mConnection.Execute strSQL
    
    strSQL = "CREATE INDEX BameiKanaKey ON HANSYOKU ("
    strSQL = strSQL & "[BameiKana]"
    strSQL = strSQL & ") "
    mConnection.Execute strSQL
    
    strSQL = "CREATE INDEX BameiEngKey ON HANSYOKU ("
    strSQL = strSQL & "[BameiEng]"
    strSQL = strSQL & ") "
    mConnection.Execute strSQL
    
    createIndex_HANSYOKU = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createIndex_HANSYOKU = False
End Function ' createIndex_HANSYOKU


'
'   機能: DBにSANKU(産駒マスタ)テーブルを作る
'
'   備考: なし
'
Public Function createIndex_SANKU() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String


    strSQL = "CREATE INDEX PrimaryKey ON SANKU ("
    strSQL = strSQL & "[KettoNum]"
    strSQL = strSQL & ") WITH PRIMARY"

    mConnection.Execute strSQL

    strSQL = "CREATE INDEX BreederCodeKey ON SANKU ("
    strSQL = strSQL & "[BreederCode]"
    strSQL = strSQL & ") "

    mConnection.Execute strSQL
    
    strSQL = "CREATE INDEX FNumKey ON SANKU ("
    strSQL = strSQL & "[FNum]"
    strSQL = strSQL & ") "

    mConnection.Execute strSQL
    
    strSQL = "CREATE INDEX MNumKey ON SANKU ("
    strSQL = strSQL & "[MNum]"
    strSQL = strSQL & ") "

    mConnection.Execute strSQL

    createIndex_SANKU = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createIndex_SANKU = False
End Function ' createIndex_SANKU


'
'   機能: DBにRECORD(レコードマスタ)テーブルを作る
'
'   備考: なし
'
Public Function createIndex_RECORD() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String

    strSQL = "CREATE INDEX PrimaryKey ON RECORD ("
    strSQL = strSQL & "[RecInfoKubun],"
    strSQL = strSQL & "[Year],"
    strSQL = strSQL & "[MonthDay],"
    strSQL = strSQL & "[JyoCD],"
    strSQL = strSQL & "[Kaiji],"
    strSQL = strSQL & "[Nichiji],"
    strSQL = strSQL & "[RaceNum],"
    strSQL = strSQL & "[TokuNum_SyubetuCD],"
    strSQL = strSQL & "[Kyori],"
    strSQL = strSQL & "[TrackCD]"
    strSQL = strSQL & ") WITH PRIMARY"

    mConnection.Execute strSQL

    createIndex_RECORD = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createIndex_RECORD = False
End Function ' createIndex_RECORD


'
'   機能: DBにHANRO(坂路調教)テーブルを作る
'
'   備考: なし
'
Public Function createIndex_HANRO() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String

    strSQL = "CREATE INDEX PrimaryKey ON HANRO ("
    strSQL = strSQL & "[TresenKubun],"
    strSQL = strSQL & "[ChokyoDate],"
    strSQL = strSQL & "[ChokyoTime],"
    strSQL = strSQL & "[KettoNum]"
    strSQL = strSQL & ") WITH PRIMARY"

    mConnection.Execute strSQL

    strSQL = "CREATE INDEX KettoNumKey ON HANRO ( [KettoNum] )"
    mConnection.Execute strSQL
    
    createIndex_HANRO = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createIndex_HANRO = False
End Function ' createIndex_HANRO


'
'   機能: DBにBATAIJYU(馬体重)テーブルを作る
'
'   備考: なし
'
Public Function createIndex_BATAIJYU() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String

    strSQL = "CREATE INDEX PrimaryKey ON BATAIJYU ("
    strSQL = strSQL & "[Year],"
    strSQL = strSQL & "[MonthDay],"
    strSQL = strSQL & "[JyoCD],"
    strSQL = strSQL & "[Kaiji],"
    strSQL = strSQL & "[Nichiji],"
    strSQL = strSQL & "[RaceNum]"
    strSQL = strSQL & ") WITH PRIMARY"

    mConnection.Execute strSQL

    createIndex_BATAIJYU = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createIndex_BATAIJYU = False
End Function ' createIndex_BATAIJYU


'
'   機能: DBにTENKO_BABA(天候馬場状態)テーブルを作る
'
'   備考: なし
'
Public Function createIndex_TENKO_BABA() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String

    strSQL = "CREATE INDEX PrimaryKey ON TENKO_BABA ("
    strSQL = strSQL & "[Year],"
    strSQL = strSQL & "[MonthDay],"
    strSQL = strSQL & "[JyoCD],"
    strSQL = strSQL & "[Kaiji],"
    strSQL = strSQL & "[Nichiji],"
    strSQL = strSQL & "[HappyoTime],"
    strSQL = strSQL & "[HenkoID]"
    strSQL = strSQL & ") WITH PRIMARY"

    mConnection.Execute strSQL

    createIndex_TENKO_BABA = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createIndex_TENKO_BABA = False
End Function ' createIndex_TENKO_BABA


'
'   機能: DBにTORIKESI_JYOGAI(出走取消・競走除外)テーブルを作る
'
'   備考: なし
'
Public Function createIndex_TORIKESI_JYOGAI() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String

    strSQL = "CREATE INDEX PrimaryKey ON TORIKESI_JYOGAI ("
    strSQL = strSQL & "[Year],"
    strSQL = strSQL & "[MonthDay],"
    strSQL = strSQL & "[JyoCD],"
    strSQL = strSQL & "[Kaiji],"
    strSQL = strSQL & "[Nichiji],"
    strSQL = strSQL & "[RaceNum],"
    strSQL = strSQL & "[Umaban]"
    strSQL = strSQL & ") WITH PRIMARY"

    mConnection.Execute strSQL

    createIndex_TORIKESI_JYOGAI = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createIndex_TORIKESI_JYOGAI = False
End Function ' createIndex_TORIKESI_JYOGAI


'
'   機能: DBにKISHU_CHANGE(騎手変更)テーブルを作る
'
'   備考: なし
'
Public Function createIndex_KISHU_CHANGE() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String

    strSQL = "CREATE INDEX PrimaryKey ON KISHU_CHANGE ("
    strSQL = strSQL & "[Year],"
    strSQL = strSQL & "[MonthDay],"
    strSQL = strSQL & "[JyoCD],"
    strSQL = strSQL & "[Kaiji],"
    strSQL = strSQL & "[Nichiji],"
    strSQL = strSQL & "[RaceNum],"
    strSQL = strSQL & "[HappyoTime],"
    strSQL = strSQL & "[Umaban]"
    strSQL = strSQL & ") WITH PRIMARY"

    mConnection.Execute strSQL

    createIndex_KISHU_CHANGE = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createIndex_KISHU_CHANGE = False
End Function ' createIndex_KISHU_CHANGE


'
'   機能: DBにMINING(データマイニング予想)テーブルを作る
'
'   備考: なし
'
Public Function createIndex_MINING() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String

    strSQL = "CREATE INDEX PrimaryKey ON MINING ("
    strSQL = strSQL & "[Year],"
    strSQL = strSQL & "[MonthDay],"
    strSQL = strSQL & "[JyoCD],"
    strSQL = strSQL & "[Kaiji],"
    strSQL = strSQL & "[Nichiji],"
    strSQL = strSQL & "[RaceNum]"
    strSQL = strSQL & ") WITH PRIMARY"

    mConnection.Execute strSQL

    createIndex_MINING = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createIndex_MINING = False
End Function ' createIndex_MINING


'
'   機能: DBにSCHEDULE(開催スケジュール)テーブルを作る
'
'   備考: なし
'
Public Function createIndex_SCHEDULE() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String

    strSQL = "CREATE INDEX PrimaryKey ON SCHEDULE ("
    strSQL = strSQL & "[Year],"
    strSQL = strSQL & "[MonthDay],"
    strSQL = strSQL & "[JyoCD],"
    strSQL = strSQL & "[Kaiji],"
    strSQL = strSQL & "[Nichiji]"
    strSQL = strSQL & ") WITH PRIMARY"

    mConnection.Execute strSQL

    createIndex_SCHEDULE = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createIndex_SCHEDULE = False
End Function ' createIndex_SCHEDULE



'
'   機能: DBにRAKaiSel(出馬表開催選択)テーブルを作る
'
'   備考: なし
'
Public Function createIndex_RAKaiSel() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String

    strSQL = "CREATE INDEX PrimaryKey ON RAKaiSel ("
    strSQL = strSQL & "[Year] DESC,"
    strSQL = strSQL & "[MonthDay] DESC,"
    strSQL = strSQL & "[JyoCD]"
    strSQL = strSQL & ") WITH PRIMARY"

    mConnection.Execute strSQL

    createIndex_RAKaiSel = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createIndex_RAKaiSel = False
End Function ' createIndex_RAKaiSel


'
'   機能: DBにODDS_SANRENTAN(オッズ_3連単)のインデックステーブルを作る
'
'   備考: なし
'
Public Function createIndex_ODDS_SANRENTAN() As Boolean
On Error GoTo Errorhandler
    Dim strSQL As String

    strSQL = "CREATE INDEX PrimaryKey ON ODDS_SANRENTAN ("
    strSQL = strSQL & "[Year],"
    strSQL = strSQL & "[MonthDay],"
    strSQL = strSQL & "[JyoCD],"
    strSQL = strSQL & "[Kaiji],"
    strSQL = strSQL & "[Nichiji],"
    strSQL = strSQL & "[RaceNum]"
    strSQL = strSQL & ") WITH PRIMARY"

    mConnection.Execute strSQL

    createIndex_ODDS_SANRENTAN = True
    Exit Function
Errorhandler:
    gApp.ErrLog
    createIndex_ODDS_SANRENTAN = False
End Function ' createIndex_ODDS_SANRENTAN


'
'   機能: DBにHASSOU_CHANGE(発走時刻変更)のインデックステーブルを作る
'
'   備考: なし
'
Public Function createIndex_HASSOU_CHANGE() As Boolean
On Error GoTo Errorhandler

    Dim strSQL As String
    
    strSQL = "CREATE INDEX PrimaryKey ON HASSOU_CHANGE ("
    strSQL = strSQL & "[Year],"
    strSQL = strSQL & "[MonthDay],"
    strSQL = strSQL & "[JyoCD],"
    strSQL = strSQL & "[Kaiji],"
    strSQL = strSQL & "[Nichiji],"
    strSQL = strSQL & "[RaceNum]"
    strSQL = strSQL & ") WITH PRIMARY"
    
    mConnection.Execute strSQL
    
    createIndex_HASSOU_CHANGE = True
    Exit Function
    
Errorhandler:
    gApp.ErrLog
    createIndex_HASSOU_CHANGE = False
End Function    'createIndex_HASSOU_CHANGE


'
'   機能: DBにCOURSE_CHANGE(コース変更)のインデックステーブルを作る
'
'   備考: なし
'
Public Function createIndex_COURSE_CHANGE() As Boolean
On Error GoTo Errorhandler

    Dim strSQL As String
    
    strSQL = "CREATE INDEX PrimaryKey ON COURSE_CHANGE ("
    strSQL = strSQL & "[Year],"
    strSQL = strSQL & "[MonthDay],"
    strSQL = strSQL & "[JyoCD],"
    strSQL = strSQL & "[Kaiji],"
    strSQL = strSQL & "[Nichiji],"
    strSQL = strSQL & "[RaceNum]"
    strSQL = strSQL & ") WITH PRIMARY"
    
    mConnection.Execute strSQL
    
    createIndex_COURSE_CHANGE = True
    Exit Function
    
Errorhandler:
    gApp.ErrLog
    createIndex_COURSE_CHANGE = False
End Function    'createIndex_COURSE_CHANGE

