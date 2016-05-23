# Worksheet  
ワークシートは xl/worksheetsフォルダに sheetN.xml のファイル名で格納  
ワークシートの情報は、3つの主なセクションに組織化されていて以下の順に格納  

1. トップレベルシートプロパティ（dimension, sheetViews, sheetFromatPr, cols ...etc）  
2. セルテーブル（sheetData）  
3. 補助シート属性


xl/worksheets/sheetN.xml のサンプル（タイトルと書式設定済みのテーブルのデータ）

## 1. トップレベルのシートプロパティ

    <?xml version="1.0" encoding="UTF-8" standalone="true"?>
    -<worksheet xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" 
    mc:Ignorable="x14ac" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" 
    xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" 
    xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <dimension ref="A1:D7"/>  -- ワークシート寸法

        -<sheetViews>  -- シート表示
            -<sheetView workbookViewId="0" tabSelected="1">
                <selection sqref="D1" activeCell="D1"/>
            </sheetView>
        </sheetViews>
        
        <sheetFormatPr x14ac:dyDescent="0.15" defaultRowHeight="13.5"/>  -- シートフォーマット属性

        -<cols>  -- カラム情報
            <col customWidth="1" bestFit="1" width="15" max="1" min="1"/>
            <col customWidth="1" bestFit="1" width="23.875" max="3" min="3"/>
            <col customWidth="1" bestFit="1" width="57.5" max="4" min="4"/>
        </cols>


## 2. セルテーブル  
▽セルのc要素のt属性＝データタイプ  
b：Booliean／e：Error／inlineStr：Inline String／n：Number（デフォルト）／s：Shared String（テキストデータ）／str：String

        -<sheetData>  -- データの親要素

        -<row r="1" x14ac:dyDescent="0.15" ht="17.25" spans="1:4">  -- 行（r：インデックス、spans：1～4列までデータ有）
            -<c r="A1" t="s" s="2">  -- セル（r：リファレンス(位置情報)、t：データタイプ⇒値により内容の要素が変わる、
                                              s：Style.xml内のインデックス番号？）
                <v>0</v>  -- 値（valueを格納 ⇔ f：数式(formula)を格納）
            </c>
        </row>

        -<row r="3" x14ac:dyDescent="0.15" spans="1:4">
            -<c r="A3" t="s" s="3">
                <v>1</v>
            </c>
            -<c r="B3" t="s" s="3">
                <v>2</v>
            </c>
            -<c r="C3" t="s" s="3">
                <v>3</v>
            </c>
            -<c r="D3" t="s" s="3">
                <v>4</v>
            </c>
        </row>

        -<row r="4" x14ac:dyDescent="0.15" spans="1:4">
            -<c r="A4" s="1">
                <v>42097.762499999997</v>
            </c>
            -<c r="B4">
                <v>9757</v>
            </c>
            -<c r="C4" t="s">
                <v>5</v>
            </c>
            -<c r="D4" t="s">
                <v>6</v>
            </c>
        </row>

        -<row r="5" x14ac:dyDescent="0.15" spans="1:4">
            -<c r="A5" s="1">
                <v>42096.625</v>
            </c>
            -<c r="B5">
                <v>7827</v>
            </c>
            -<c r="C5" t="s">
                <v>8</v>
            </c>
            -<c r="D5" t="s">
                <v>7</v>
            </c>
        </row>

        -<row r="6" x14ac:dyDescent="0.15" spans="1:4">
            -<c r="A6" s="1">
                <v>42096.625</v>
            </c>
            -<c r="B6">
                <v>3083</v>
            </c>
            -<c r="C6" t="s">
                <v>9</v>
            </c>
            -<c r="D6" t="s">
                <v>10</v>
            </c>
        </row>

        -<row r="7" x14ac:dyDescent="0.15" spans="1:4">
            -<c r="A7" s="1">
                <v>42096.416666666664</v>
            </c>
            -<c r="B7">
                <v>5938</v>
            </c>
            -<c r="C7" t="s">
                <v>12</v>
            </c>
            -<c r="D7" t="s">
                <v>11</v>
            </c>
        </row>
        
        </sheetData>


## 3. 補助シート属性  
他にシートの状態を定義する子要素を持つ。以下はサンプルファイルの属性で、他にも子要素は多数あり

        <phoneticPr fontId="1"/>  -- 音声情報
        <pageMargins footer="0.3" header="0.3" bottom="0.75" top="0.75" right="0.7" left="0.7"/>  -- ページマージン
        <pageSetup r:id="rId1" orientation="portrait" paperSize="9"/>  -- ページ設定

        -<tableParts count="1">  -- シート上のテーブル情報
            <tablePart r:id="rId2"/>
        </tableParts>
    </worksheet>

