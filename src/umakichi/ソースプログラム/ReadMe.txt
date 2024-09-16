―――――――――――――――――――――――――――――――――――――――

　　　　　　　　　馬吉ソース公開版 for Data Lab. Ver 1.0.1

―――――――――――――――――――――――――――――――――――――――
目次

　　１．はじめに
　　２．ライセンスについて
　　３．開発環境
　　４．動作環境
　　５．馬吉 Ver.5.3.3との違い
　　６．ファイル一覧

―――――――――――――――――――――――――――――――――――――――
１．はじめに

馬吉 for Data Lab.(以下、馬吉)のソースを公開いたします。
競馬ソフトを作成する際の参考プログラムとしてご利用下さい。

―――――――――――――――――――――――――――――――――――――――
２．ライセンスについて

┌───────────────────────────────────┐
│ Copyright (c) 2007, JRA SYSTEM SERVICE CO.,LTD. All rights reserved. │
└───────────────────────────────────┘

馬吉（以降本ソフトウェア）はソースコード形式かバイナリ形式か、変更するかしない
かを問わず、以下の条件を満たす場合に限り、再頒布および使用が許可されます。

（１）ソースコードを再頒布する場合、上記の著作権表示、本条件一覧、および下記免
　　　責条項を含めること。 

（２）バイナリ形式で再頒布する場合、頒布物に付属のドキュメント等の資料に、上記
　　　の著作権表示、本条件一覧、および下記免責条項を含めること。

（３）本ソフトウェアの使用は、JRAシステムサービス（株）が 提供する競馬
　　　データ取得サービス「JRA-VAN Data Lab.」の利用に限る。

（４）本ソフトウェアの改変に当り、下記文面を表示している箇所を削除、変更しては
　　　ならない。

　　　「本ソフトウェアはJRAシステムサービス（株）提供のソフトウェア（馬吉）
	を基に作成されました。」

（５）書面による特別の許可なしに、本ソフトウェアから派生した製品の宣伝または販
　　　売促進に、JRAシステムサービス（株）の名前またはJRA-VANの名前を使
　　　用してはならない。 

本ソフトウェアは、著作権者によって「現状のまま」提供されており、明示黙示を問わ
ず、商業的な使用可能性、および特定の目的に対する適合性に関する暗黙の保証も含め
またそれに限定されない、いかなる保証もありません。著作権者は、事由のいかんを問
わず、 損害発生の原因いかんを問わず、かつ責任の根拠が契約であるか厳格責任であ
るか（過失その他の）不法行為であるかを問わず、仮にそのような損害が発生する可能
性を知らされていたとしても、本ソフトウェアの使用によって発生した（代替品または
代用サービスの調達、使用の喪失、データの喪失、利益の喪失、業務の中断も含め、ま
たそれに限定されない）直接損害、間接損害、偶発的な損害、特別損害、懲罰的損害、
または結果損害について、一切責任を負わないものとします。

―――――――――――――――――――――――――――――――――――――――
３．開発環境

馬吉は Microsoft VisualBasic 6.0(SP6) Professional Editionを使用して開発されて
います。

※：一部コントロールにおいて、Learning Editionに含まれない開発ライセンスのコン
　　トロールを使用しております。
※：Visual Basic .NETでは動作いたしません。コンバージョン作業が必要になります。
　　（Visual Basicアップグレードウィザードのみでは動作いたしません。)

―――――――――――――――――――――――――――――――――――――――
４．動作環境

馬吉は以下の環境で動作します。
またJRA-VAN DataLab.でデータを入手するためにはインターネット接続環境が必要に
なります。詳しくはJRA-VAN Data Lab. データ入手方法をご覧下さい。

・対応ＯＳ 
Microsoft Windows(R)98 
Microsoft Windows(R)Me 
Microsoft Windows(R)2000 
Microsoft Windows(R)XP

・ハードディスク 
空き容量推奨6GB以上

・必要メモリ 
最低128MB、推奨256MB以上

・ディスプレイ 
解像度 800×600、推奨解像度 1024×768以上

―――――――――――――――――――――――――――――――――――――――
５．馬吉 Ver.5.3.3との違い

公開ソースは、Ver5.3.3をベースとし、以下の機能変更を実施しております。

(1)使用グリッドコントロールの変更
　馬吉 Ver5.3.3では、「GrapeCity VS-FlexGrid Pro Ver8.0」を利用しておりますが、
　標準では含まれない開発ライセンスを必要とするコントロールであるため、
　「Microsoft FlexGrid Control Ver6.0」に変更しております。
　これにより、該当コントロールに依存していた下記の機能は削除となりました。
　・列の固定機能
　　横スクロールをしても、特定の列を表示し続ける機能
　・グリッドのソート機能
　　グリッドの行見出しをクリックすることで、列のソートを実行する機能
(2)印刷機能の削除
　印刷機能についてはソース公開の対象外としております。
(3)テキストエキスポート機能の削除
(4)旧バージョン(3連単対応前)DBからのアップデート機能の削除

―――――――――――――――――――――――――――――――――――――――
６．Umakichi.vbpに組み込まれるファイル

basAPI.bas              API宣言モジュール
basEnum.bas             Enum宣言モジュール
basFlexgrid.bas         FlexGrid ラッパー
basIni.bas              INIファイルに関するモジュール
basMain.bas             起動モジュール いくつかのユーティリティーFunctionを含む
basReg.bas              レジストリに関するモジュール
basSetDataFromByte.bas  データセット関数
basSetDataFromRS.bas    構造体にレコードデータを取得するモジュール
clsApp.cls              アプリケーションクラス
clsBrowserMgr.cls       ブラウザマネージャ　Browser Manager すべてのブラウザを管理する
clsCodeConverter.cls    コード変換クラス
clsCreateMDB.cls        MDB作成、テーブル定義をするクラス
clsDatabaseMgr.cls      データベースマネージャ
clsDataBN.cls           馬主 データクラス
clsDataBR.cls           生産者 データクラス
clsDataCH.cls           調教師 データクラス
clsDataFind.cls         検索画面データ取得オブジェクト
clsDataHCSel.cls        坂路調教選択画面データクラス
clsDataHK.cls           変更情報 データクラス
clsDataHN.cls           繁殖馬 データクラス
clsDataKS.cls           騎手 データクラス
clsDataOD.cls           オッズ・票数 データクラス
clsDataRA.cls           RA - RACE データ取得オブジェクト
clsDataRaceChanger.cls  タイトルバンドのレースチェンジャーコンボボックスのデータを取得、保持する
clsDataRAKaiSel.cls     出馬表開催選択画面データクラス
clsDataRASel.cls        出馬表選択画面データクラス
clsDataRC.cls           レコード データクラス
clsDataRCSel.cls        レコード一覧画面 データクラス
clsDataSK.cls           産駒 データクラス
clsDataTK.cls           特別登録馬データ取得オブジェクト
clsDataTKKaiSel.cls     特別登録馬開催選択画面データクラス
clsDataTKSel.cls        特別登録馬選択画面データクラス
clsDataUM.cls           競走馬 データクラス
clsGridData.cls         グリッドデータクラス
clsGridItem.cls         グリッドアイテムクラス
clsH1Iterator.cls       H1集合体 クラス
clsHistoryItem.cls      履歴アイテム クラス
clsHistoryMgr.cls       履歴管理クラス
clsIImport.cls          データベース登録クラス共通インターフェイス
clsImportAV.cls         JVData "AV" データベース登録クラス
clsImportBN.cls         JVData "BN" データベース登録クラス
clsImportBR.cls         JVData "BR" データベース登録クラス
clsImportCC.cls         JVData "CC" データベース登録クラス
clsImportCH.cls         JVData "CH" データベース登録クラス
clsImportDM.cls         JVData "DM" データベース登録クラス
clsImportH1.cls         JVData "H1" データベース登録クラス
clsImportHC.cls         JVData "HC" データベース登録クラス
clsImportHN.cls         JVData "HN" データベース登録クラス
clsImportHR.cls         JVData "HR" データベース登録クラス
clsImportJC.cls         JVData "JC" データベース登録クラス
clsImportKS.cls         JVData "KS" データベース登録クラス
clsImportO1.cls         JVData "O1" データベース登録クラス
clsImportO2.cls         JVData "O2" データベース登録クラス
clsImportO3.cls         JVData "O3" データベース登録クラス
clsImportO4.cls         JVData "O4" データベース登録クラス 馬単オッズは、データ容量が多いため、MDBを10分割します。
clsImportO5.cls         JVData "O5" データベース登録クラス 3連複オッズは、データ容量が多いため、MDBを10分割します。
clsImportODDS.cls       JVData オッズ票数 データベース登録クラス
clsImportRA.cls         JVData "RA" データベース登録クラス
clsImportRC.cls         JVData "RC" データベース登録クラス
clsImportSE.cls         JVData "SE" データベース登録クラス 馬毎レース情報は、データ容量が多いため、MDBを2分割します。
clsImportSK.cls         JVData "SK" データベース登録クラス
clsImportTC.cls         JVData "TC" データベース登録クラス
clsImportTK.cls         JVData "TK" データベース登録クラス
clsImportUM.cls         JVData "UM" データベース登録クラス
clsImportWE.cls         JVData "WE" データベース登録クラス
clsImportWH.cls         JVData "WH" データベース登録クラス
clsImportYS.cls         JVData "YS" データベース登録クラス
clsIterator.cls         集合体 クラス
clsIViewerState.cls     ビュアーステート 共通インターフェイス
clsKeyBN.cls            馬主マスタ  keyクラス
clsKeyBR.cls            生産者マスタ    keyクラス
clsKeyCH.cls            競走馬  keyクラス
clsKeyHCSel.cls         坂路調教選択 keyクラス
clsKeyHN.cls            繁殖馬マスタ    keyクラス
clsKeyKS.cls            騎手マスタ keyクラス
clsKeyRA.cls            出馬表  keyクラス
clsKeyRAKaiSel.cls      開催選択画面用キー 年度と場所コードで絞り込む 場所コードが 00 の場合は、全競馬場
clsKeyRASel.cls         出馬表選択 keyクラス
clsKeyRC.cls            レコードマスタ  keyクラス
clsKeyRCSel.cls         レコード一覧画面
clsKeySK.cls            産駒マスタ  keyクラス 　産駒とペテルブルグ
clsKeyUM.cls            競走馬  keyクラス
clsMSFlexData.cls       グリッドのプロパティを保持するクラス
clsPointer.cls          ポインター クラス
clsRCSearch.cls         レコード検索 クラス
clsStringConverter.cls  文字列変換クラス
clsViewerBase.cls       Viewer基底クラス   ctlV* はすべて、このインスタンスを持つこと
clsVSDate.cls           Viewer State 日時行数の状態を持つViewer用(坂路)
clsVSDateJyo.cls        Viewer State 日時場所行数の状態を持つViewer用(坂路)
clsVSFind.cls           Viewer State VFind用
clsVSNothing.cls        Viewer State 状態変化を持たないViewer用
clsVSOdds.cls           Viewer State OD Viewer用
clsVSTabOnly.cls        Viewer State カレントタブ情報のみ持つViewer用
clsVSYearJyo.cls        Viewer State 年度場所行数の状態を持つViewer用(開催選択等)
ctlLabel.ctl            ラベル１つをラップし、右クリックによるポップアップメニューを持ち 通常リンクオープン(ChangeViewer)、あるいは、新規ウインドウで開く(NewWindow)の 二種類のイベントを生成。
ctlMenu.ctl             メニューパレット ユーザーコントロール
ctlPane.ctl             読み込み中、データがありません、有効、の３状態を持つコンテナ領域
ctlTitleBand.ctl        タイトルバンド
ctlToolBars.ctl         ツールバー
ctlVBN.ctl              馬主マスタ 表示コントロール
ctlVBR.ctl              生産者 表示コントロール
ctlVCH.ctl              調教師マスタ 表示コントロール
ctlVFind.ctl            検索表示コントロール
ctlVHC.ctl              坂路一覧 表示コントロール
ctlVHCSel.ctl           坂路一覧選択 表示コントロール
ctlVHK.ctl              変更情報表示ユーザーコントロール
ctlVHN.ctl              繁殖馬マスタ 表示コントロール
ctlVHome.ctl            ホーム画面 ホームボタンを押すと、この画面が表示される デフォルトの起動時画面もこの画面
ctlVKS.ctl              騎手マスタ 表示コントロール
ctlVOD.ctl              オッズ･票数 表示コントロール
ctlVRA.ctl              出馬表 表示コントロール   出馬表Viewer
ctlVRAKaiSel.ctl        出馬表開催選択 表示コントロール
ctlVRASel.ctl           出馬表選択 表示コントロール
ctlVRC.ctl              レコード 表示コントロール
ctlVRCSel.ctl           レコード選択 表示コントロール
ctlVSK.ctl              産駒マスタ 表示コントロール
ctlVTK.ctl              特別登録馬 表示コントロール
ctlVTKKaiSel.ctl        特別登録馬開催選択 表示コントロール
ctlVTKSel.ctl           特別登録馬  選択 表示コントロール
ctlVUM.ctl              競走馬  表示コントロール
ctlWrappedGrid.ctl      MSFlexGridをラップするユーザーコントロール
frmAbout.frm            ヘルプ フォーム
frmBrowser.frm          ブラウザーフォーム Viewerを乗せるコンテナになります。 WebBrowserに似たインターフェイスを持っています。 Viewerがはみ出す場合スクロールバーで制御します。
frmConfig.frm           環境設定画面
frmConfigFirst.frm      (初)データ取得設定 ダイアログ
frmDBMaintenance.frm    データベース最適化ダイアログ
frmDBPrompt.frm         速報データ取得 ダイアログ
frmDBRAKaiSel.frm       開催情報作成 ダイアログ
frmDBUpdate.frm         frmDBUpdate には、 Other
frmDirRef.frm           フォルダを選択するダイアログボックス 起動時に選択されているパスを、起動前に BeginingPath Property で、設定できる。 起動前に、Stringを Message Property に設定できる。 このフォームは、必ず .Show vbModal で呼び出す。 起動後に ReturnPath Property で選択したFolderPathを読める。
frmMenu.frm             メニューパレットフォーム
frmNewFolder.frm        フォルダ作成 ダイアログ
frmSplash.frm           スプラッシュウインドウ 馬吉の起動時に、著作権表記などを表示する 主にデータベースのチェック中の待ち時間に表示
frmWrappedJVLink.frm    JVLinkがインストールされているかどうかの判定の為の隠しフォーム Visible=Falseでユーザーからは隠匿されます。
JVData_Structure.bas    DataLab. 構造体
umakichi.res		リソースファイル

