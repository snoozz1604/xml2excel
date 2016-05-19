# xml2excel

## フォルダ構成
_以下は、空のワークシートが1つだけあるExcelのフォルダ構成の内容_

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
_以下は、空のワークシートが1つだけあるExcelの内容タイプ定義 **[Content_Types].xml** の内容_

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
リレーションシップファイルは階層ごとに**［ _rel ］フォルダ**内に作成され、**拡張子はrels**となる  
例）A.xmlというファイル名のパーツのリレーションシップのファイルはA.xml.rels  
_以下は空のワークシートが1つだけあるExcelの **[ _rels ].relsファイル**の内容_

        <?xml version="1.0" encoding="UTF-8" standalone="true"?>  
        -<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">  
        <Relationship Target="docProps/app.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Id="rId3"/>  
        <Relationship Target="docProps/core.xml" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Id="rId2"/>  
        <Relationship Target="xl/workbook.xml" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Id="rId1"/>  
        </Relationships>

ファイルの内容は上記のように、**Relatinships要素をルート**として、**Relatinship要素が列挙**される  
Relatinship要素は、**Type属性**と**Target属性**を持ち、参照のための**Id値**を持つ


## アプリケーション情報  
docsPropsフォルダにapp.xml（アプリケーション情報）とcore.xml（ドキュメント情報）のファイルを持つ  

### app.xml  
**Properties要素をルート**として、各種要素が列挙され、アプリケーション情報の内容が記述される
_以下は空のワークシートが1つだけあるExcelの **[ docProps ]app.xmlファイル**の内容_

        <?xml version="1.0" encoding="UTF-8" standalone="true"?>  
        -<Properties xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes" xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">  
                <Application>Microsoft Excel</Application>  --アプリケーション名
                <DocSecurity>0</DocSecurity>  --ドキュメントセキュリティ
                <ScaleCrop>false</ScaleCrop>   --サムネイル表示モード
                -<HeadingPairs>  --見出し
                        -<vt:vector baseType="variant" size="2">
                                -<vt:variant>
                                        <vt:lpstr>ワークシート</vt:lpstr>
                                </vt:variant>
                                -<vt:variant>
                                        <vt:i4>1</vt:i4>
                                </vt:variant>
                        </vt:vector>
                </HeadingPairs>
                -<TitlesOfParts>  --タイトル
                        -<vt:vector baseType="lpstr" size="1">
                                <vt:lpstr>Sheet1</vt:lpstr>
                        </vt:vector>
                </TitlesOfParts>
                <Company>Microsoft</Company>  --会社名
                <LinksUpToDate>false</LinksUpToDate>  --最新状態へのリンク
                <SharedDoc>false</SharedDoc>  --共有ドキュメント
                <HyperlinksChanged>false</HyperlinksChanged>  --変更されたハイパーリンク
                <AppVersion>14.0300</AppVersion>  --アプリケーションバージョン
        </Properties>

### core.xml  
**coreProperties要素をルート**として、各種要素が列挙され、ドキュメント情報の内容が記述される
_以下は空のワークシートが1つだけあるExcelの **[ docProps ]app.xmlファイル**の内容_

        <?xml version="1.0" encoding="UTF-8" standalone="true"?>
        -<cp:coreProperties xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties">
                <dc:creator>作成者</dc:creator>
                <cp:lastModifiedBy>前回保存者</cp:lastModifiedBy>
                <dcterms:created xsi:type="dcterms:W3CDTF">2016-05-19T05:33:06Z</dcterms:created>
                <dcterms:modified xsi:type="dcterms:W3CDTF">2016-05-19T05:33:32Z</dcterms:modified>
        </cp:coreProperties>


> 引用・参考文献）Office Open XML Formats入門〔マイコミ〕
