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
内容タイプ定義 **[Content_Types].xml** は、パッケージﾞを解凍するとルートに存在するファイル  
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


## リレーションシップによる関連付け  
複数あるパーツの構成とパーツへの参照を解決する役割をリレーションシップ＝関連付けという  
リレーションシップファイルは階層ごとに［ _rel ］フォルダ内に作成され、**拡張子はrels**となる  
例）A.xmlというファイル名のパーツのリレーションシップのファイルはA.xml.rels

        <?xml version="1.0" encoding="UTF-8" standalone="true"?>  
        -<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">  
        <Relationship Target="docProps/app.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Id="rId3"/>  
        <Relationship Target="docProps/core.xml" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Id="rId2"/>  
        <Relationship Target="xl/workbook.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Id="rId1"/>  
        </Relationships>

ファイルの内容は上記のように、**Relatinships要素をルート**として、**Relatinship要素が列挙**される  
Relatinship要素は、**Type属性**と**Target属性**を持ち、参照のための**Id値**を持つ
