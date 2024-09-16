README 2.2.0
―――――――――――――――――――――――――――――――――――――――
　　　　　　　　　　　　　　　　　
　　　　　　　　　　　　　　　JV-Data登録クラス

―――――――――――――――――――――――――――――――――――――――
目次

　　１．はじめに
　　２．ファイル一覧
　　３．動作環境
　　４．使い方
　　５．注意事項

―――――――――――――――――――――――――――――――――――――――
１．はじめに

JV-Data登録クラスは、JV-Linkのメソッド"JVRead"で読み込んだJV-Dataレコードを
データベースに登録するためのプログラミングパーツです。
このクラスは、同梱されているデータベースファイル（Data.accdb)をデータベースとして
使用します。また、パーツとして提供されているデータベース作成クラスを使って作成
することも可能です。
―――――――――――――――――――――――――――――――――――――――
２．ファイル一覧

bin                    ［フォルダ］
　－－－ Data.accdb       データベースファイル
　－－－ sample.ini      データベース接続設定ファイル
AssemblyInfo.vb
basUtility.vb			ユーティリティクラス
basVar.bas		変数定義ファイル
clsDBImport.vb	　　　　データベース接続クラス
clsImportAV.vb		JV-Data登録クラス
clsImportBN.vb			〃
clsImportBR.vb			〃
clsImportCH.vb			〃
clsImportDM.vb			〃
clsImportH1.vb			〃
clsImportH6.vb			〃
clsImportHC.vb			〃
clsImportHN.vb			〃
clsImportHR.vb			〃
clsImportHS.vb			〃
clsImportHY.vb			〃
clsImportJC.vb			〃
clsImportKS.vb			〃
clsImportO1.vb			〃
clsImportO2.vb			〃
clsImportO3.vb			〃
clsImportO4.vb			〃
clsImportO5.vb			〃
clsImportO6.vb			〃
clsImportRA.vb			〃
clsImportRC.vb			〃
clsImportSE.vb			〃
clsImportSK.vb			〃
clsImportTK.vb			〃
clsImportUM.vb			〃
clsImportWE.vb			〃
clsImportWH.vb			〃
clsImportYS.vb			〃
clsImportTC.vb			〃
clsImportCC.vb			〃
clsImportBT.vb			〃
clsImportCS.vb			〃
clsImportJG.vb			〃
clsImportTM.vb			〃
clsImportWF.vb			〃
frmDataImport.resX	JV-Data登録クラスサンプル用フォームXMLリソース
frmDataImport.vb	JV-Data登録クラスサンプル用フォームリソース
JVData_Structure.vb	JV-Data構造体
README.TXT		はじめにお読みください
SampleProject.vbproj	JV-Data登録クラスサンプル用プロジェクトファイル
SampleProject.sln	JV-Data登録クラスサンプル用プロジェクトソリューションファイル

―――――――――――――――――――――――――――――――――――――――
３．動作環境

提供されるクラスおよびサンプルはMicrosoft VisualBasic 2015で作成、確認されています。

―――――――――――――――――――――――――――――――――――――――
４．使い方

アプリケーションに組み込む場合は、プロジェクトに以下のファイルを追加してください。

basUtility.vb			ユーティリティクラス
basVar.bas		変数定義ファイル
clsDBImport.vb	　　　　データベース接続クラス
JVData_Structure.vb	JV-Data構造体
clsImportAV.vb		JV-Data登録クラス
clsImportBN.vb			〃
clsImportBR.vb			〃
clsImportCH.vb			〃
clsImportDM.vb			〃
clsImportH1.vb			〃
clsImportH6.vb			〃
clsImportHC.vb			〃
clsImportHN.vb			〃
clsImportHR.vb			〃
clsImportHS.vb			〃
clsImportHY.vb			〃
clsImportJC.vb			〃
clsImportKS.vb			〃
clsImportO1.vb			〃
clsImportO2.vb			〃
clsImportO3.vb			〃
clsImportO4.vb			〃
clsImportO5.vb			〃
clsImportO6.vb			〃
clsImportRA.vb			〃
clsImportRC.vb			〃
clsImportSE.vb			〃
clsImportSK.vb			〃
clsImportTK.vb			〃
clsImportUM.vb			〃
clsImportWE.vb			〃
clsImportWH.vb			〃
clsImportYS.vb			〃
clsImportTC.vb			〃
clsImportCC.vb			〃
clsImportBT.vb			〃
clsImportCS.vb			〃
clsImportJG.vb			〃
clsImportTM.vb			〃
clsImportWF.vb			〃
bin			[フォルダ]
・・・・ Data.accdb
・・・・ sample.ini

また、登録先のデータベースとして同梱されたData.accdbもしくは、「データベース作成クラス」
で作成したデータベースファイルが必要となります。データベースは、［bin］フォルダの中に
追加して下さい。


―――――――――――――――――――――――――――――――――――――――
５．注意事項

(1)JV-Data登録クラスは、実行フォルダに存在するdata.accdbにデータを登録します。
   当JV-Data登録クラスでは、全レコード種別を対象としております。
   処理時間についてはデータ量に応じた処理時間となりますので、必要に応じて、
   処理対象データを選定して下さい。

   データベースにaccdbを利用しているため登録可能データ量等、accdbの仕様を考慮して
   利用する必要があります。

   目的に応じてソースプログラム、データベースデザイン等を自由に変更し利用して
   頂いて構いません。

(2)INSERT or UPDATEを実行するとaccdbが肥大化します。定期的にデータベースを
   最適化して下さい。(最適化中はデータベースを開かないで下さい。)


--------------------------------------------------------------------------------
JV-Linkおよびサンプルプログラムの使用許諾について
--------------------------------------------------------------------------------

お客様は、以下の条項に同意されない場合、ＪＲＡシステムサービス株式会社は、お
客様に本プログラムのインストール、使用または複製のいずれも許諾できません。

[総則]
本プログラムは、ＪＲＡシステムサービス株式会社の提供するJRA-VANサービスにおい
て使用されるものとし、サービスを有料で利用する場合は「JRA-VAN利用規約」の各条項が
優先して適用されます。

[禁止事項]
ＪＲＡシステムサービス株式会社は、本プログラムの使用にあたり以下の事項を禁止し
ます。
本プログラムのリバースエンジニアリング等プログラムロジックを解析すること本プログラ
ムとサーバ間のインターフェースを解析すること本プログラム以外から本プログラム専用サ
ーバにアクセスすること本プログラムを複製すること本プログラムにより取得した情報を第
三者に知らせること（インターネット上に公開することを含む）

[無保証]
ＪＲＡシステムサービス株式会社は、本プログラムおよびサービスサポートに関して、
いかなる保証も一切いたしません。本プログラムおよびサービスサポートの使用または機能
から生じるすべての危険は、お客様が負担しなければなりません。

[免責事項]
ＪＲＡシステムサービス株式会社は、本プログラムの使用もしくは使用不能、あるいは
サーバからのデータ提供もしくは不提供から生じる損害（逸失利益、機密情報もしくはその
他の情報の喪失、仕事の中断、人身傷害、プライバシーの喪失、信義則または合理的な注意
義務を含めた義務の不履行または過失による、金銭的またはその他の損害を含み、かつこれ
らに限定されません）に関しては、一切責任を負いません。たとえＪＲＡシステムサービス株
式会社がこのような損害の可能性について知らされていた場合でも同様です。
