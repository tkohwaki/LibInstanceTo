## このプロジェクトについて

このプロジェクトはクラスインスタンスをExcelまたはCSVに変換するライブラリです。

下記のネームスペースを持ちます。
- LibInstanceTo  
  ライブラリ基本
- LibInstanceTo.Excel  
  Excel変換用
- LibInstanceTo.CSV  
  CSV変換用

### Excel変換

下記のコンストラクタ,メソッドが使用可能です。  
- InstanceToExcel<T,V>(Stream OutputExcelStream,string DefFile)  
  Excelの雛形のStreamと変換定義ファイル名を指定します。Streamがnullの場合は空のExcelファイルへの出力となります。
- InstanceToExcel<T,V>(string OutputExcelFileName, string DefFile)  
    Excelの雛形のファイル名と変換定義ファイル名を指定します。Streamが空文字列("")の場合は雛形無しのExcelファイルへの出力となります。
- ConvertOne(int Row,T Instance)
  インスタンスを1行分だけ出力します。
- Convert(List<T> Instances)  
  指定された行数分、インスタンスを出力します。
- SaveAs(Stream OutputExcelStream)  
  指定StreamにExcelを出力します。
- SaveAs(string FileName)  
  指定ファイルにExcelを出力します。
