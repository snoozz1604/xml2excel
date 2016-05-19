# xml2excel

## フォルダ構成

        [ _rels ]  
                .rels  
        [ docProps ]  
                app.xml  
                core.xml  
        [ xl ]  
                [ _rels ]  
                        workbook.xml.rels  
                [ theme ]  
                        theme1.xml  
                [ worksheets ]  
                        sheet1.xml  
                        styles.xml  
                        workbook.xml  
        [Content_Types].xml  

## パッケージの構成  
パッケージを構成するパーツの種類は以下の通り

1. **コンテンツタイプ（内容タイプ定義）**  … 1つのみ存在（ファイル名固定 ⇒ [Content_Type].xml）

2. **リレーションシップ（関連付け）**  … 複数存在

3. **ドキュメントプロパティ（アプリケーション情報）**

4. **ドキュメントパーツ**  … アプリケーション固有のドキュメントデータ

1 ~ 3 はアプリケーション共通の仕様  
4 はパッケージごとにword、xl、pptというフォルダ下に格納


## 内容タイプ  
内容タイプ定義 [Content_Types].xml は、パッケージﾞを解凍するとルートに存在するファイル  
どのパーツがどのような機能（属性）を持っているのかを示す  
以下は、空のワークシートが1つだけあるExcelの内容タイプ定義

        <?xml version="1.0" encoding="UTF-8" standalone="true"?>  
        -<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">  
        <Default ContentType="application/vnd.openxmlformats-package.relationships+xml" Extension="rels"/>  
        <Default ContentType="application/xml" Extension="xml"/>  
        <Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" PartName="/xl/workbook.xml"/>  
        <Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" PartName="/xl/worksheets/sheet1.xml"/>  
        <Override ContentType="application/vnd.openxmlformats-officedocument.theme+xml" PartName="/xl/theme/theme1.xml"/>  
        <Override ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml" PartName="/xl/styles.xml"/>  
        <Override ContentType="application/vnd.openxmlformats-package.core-properties+xml" PartName="/docProps/core.xml"/>  
        <Override ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml" PartName="/docProps/app.xml"/>  
        </Types>  

