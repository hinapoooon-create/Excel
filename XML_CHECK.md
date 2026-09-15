# シート4の読み込みエラー修正

2026-09-15。対象は `従業員マスター.xlsm` と `330bf0cd-_______.xlsm`（同一内容）。

SHA-256：`5ee5c0b2c7d334434a3dbe097048e608c4eba6ef78ddc17a1e1ddc31d76594d9`

## 原因と修正

`xl/worksheets/sheet4.xml` の末尾で、ボタンの参照を持つ `legacyDrawing` が、テーブルの参照を持つ `tableParts` より後に置かれていた。これはExcelのWorksheet要素の順序に違反する。

```xml
<!-- 修正前 -->
<tableParts count="1">...</tableParts>
<legacyDrawing r:id="rIdButtons" />

<!-- 修正後 -->
<legacyDrawing r:id="rIdButtons" />
<tableParts count="1">...</tableParts>
```

順序は [ECMA-376 Part 4（第5版）](https://ecma-international.org/publications-and-standards/standards/ecma-376/)に付属する `OfficeOpenXML-XMLSchema-Transitional.zip` の `sml.xsd`、`CT_Worksheet` の定義に基づく。[MicrosoftのWorksheetクラスの参照情報](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.spreadsheet.worksheet?view=openxml-3.0.1)も参照。

前回の点検はXMLが文書として解析できることを確認していたが、公式XSDに対する検証をしておらず、この要素順序の不備を見落としていた。

## 修正前後の検証

同じ公式XSDと同じ検証コードで、配布済みファイルと修正版を比較した。

| 対象 | 合格 | 不合格 |
| --- | ---: | ---: |
| 修正前（コミット `f268604`） | 37パーツ | シート4の1パーツ |
| 修正後 | 38パーツ | 0パーツ |

検証対象は全19シート、全16テーブル、workbook.xml、styles.xml、sharedStrings.xml。

修正前の検証エラーは `legacyDrawing: This element is not expected`。修正後は同じエラーが解消した。検証コードは [`tools/validate_ooxml.py`](tools/validate_ooxml.py) に保存した。公式スキーマは改変せず使用した。

ZIPのCRCチェックも通過。コミット `f268604` との差分は `xl/worksheets/sheet4.xml` のみで、その変更も上記2要素の入れ替えのみ。他のパーツはバイト単位で一致している。

VBA本体 `xl/vbaProject.bin` のSHA-256は修正前後とも `c820e3cc749c0c7efde77c5f72aeb774eff53e1c22af7b83351ffc7fc7170194`。

これはXML規格に対する検証結果であり、Windows Excel実機で開く操作やVBA・Power Queryの実行を確認した結果ではない。
