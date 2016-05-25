# Shared String Table  
ワークシート内の文字テータ（テキスト）は、Shared String Tableと呼ばれる文字列データ用のパーツに格納。  
ワークブックと同階層のxlフォルダ内にsharedStrings.xml で存在。  

sheetN.xmlファイル内）c要素の属性t='s'のとき、v要素の値がShared String Tableのインデックスを表す。  
Shared String Table（sharedStrings.xml）の内容は、sst要素をルートとしてShared Itemのsi要素が列挙される↓

	<?xml version="1.0" encoding="UTF-8" standalone="true"?>
	-<sst uniqueCount="13" count="13" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
		-<si>
			<t>適時開示情報</t>
			-<rPh eb="2" sb="0">
				<t>テキジ</t>
			</rPh>
			-<rPh eb="4" sb="2">
				<t>カイジ</t>
			</rPh>
			-<rPh eb="6" sb="4">
				<t>ジョウホウ</t>
			</rPh>
			<phoneticPr fontId="1"/>
		</si>
		-<si>
			<t>公開日時</t>
			-<rPh eb="2" sb="0">
				<t>コウカイ</t>
			</rPh>
			-<rPh eb="4" sb="2">
				<t>ニチジ</t>
			</rPh>
			<phoneticPr fontId="1"/>
		</si>
		-<si>
			<t>コード</t>
			<phoneticPr fontId="1"/>
		</si>
		：
	</sst>

# Calculation Chain  
SpreadsheetMLでは、ワークシートに連鎖している計算式がある場合にCaluculation Chainというパーツを作成。  
このパーツはワークブックと同じxlフォルダにcalcChain.xmlのファイル名で格納。  
calcChain要素をルートにc要素を列挙。r属性はsheetDataと同じくセルの位置情報。  
i属性はシートID。シートIDが省略された時は直前のセルのシートIDと同じ値。  
s属性はこのセルの計算が他のセルに依存しているかを表す。デフォルトはfalse(0)。  
true(1)は依存しており、false(0)は依存していないので単独でセル計算を行う。

# comment  
ワークシートのコメントの内容はコメント用のパーツに保存。xlフォルダにcommentN.xmlとシート単位に作成。  
ファイルの番号はワークシート番号とは一致しない。  
**コメントパーツの構造**  

    comments
      authors       author  
      commentList   comment     text    t  
                                        r   rPr  
                                            t


# Chart  



