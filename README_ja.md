# VBAForm2Tkinter - VBA UserForm to Tkinter Converter
🌎[English](https://github.com/GUI-Conversion-Tools/VBAForm2Tkinter/blob/main/README.md)<br><br>
このプログラムは、MIcrosoft Office VBAにて作成したユーザーフォームをPythonのTkinter用に変換可能なプログラムです<br>

## 変換例
<img width="681" height="1275" alt="Image" src="https://github.com/user-attachments/assets/e4d37e92-4418-4f69-842e-25275bf596d6" /><br>
<img width="704" height="695" alt="Image" src="https://github.com/user-attachments/assets/45678575-a162-4e18-a516-ba522727d1f9" /><br><br>

## 動作要件
- 対応OS: Windows XP以上
- 必要ソフトウェア: Microsoft Excel/Word/PowerPoint/Outlook 2000以降
- 推奨環境: Microsoft Excel 2016以降

## 動作確認済環境
- Windows XP(SP3) 
- Windows 10/11
- Excel 2000(32bit)
- Excel 2010(32bit)
- Excel 2016(32bit)
- Excel 2019(64bit)
- Word/PowerPoint/Outlook 2000 (32bit)
- Word/PowerPoint/Outlook 2019 (64bit)

## 反映する項目
- 変数名(オブジェクト名)
- コントロールのおおよそのレイアウトとサイズ
- コントロールの色(文字色)
- コントロールの色(背景色)
- テキスト表示(`Label`, `CommandButton`, `CheckBox`, `ToggleButton`, `OptionButton`, `MultiPage`)
- フォント(フォント種類、サイズ、太字、斜体)
- 枠線(`UserForm`, `Frame`, `TextBox`, `Label`, `ListBox`, `Image`)
- マウスカーソル
- テキスト表示の左寄せ・中央・右寄せ(`Label`, `TextBox` [.MultiLine=False], `ComboBox`, `CheckBox`, `ToggleButton`, `OptionButton`, `ListView`)
- `TextBox`, `ComboBox`のデフォルト値
- `ComboBox`, `ListBox`, `ListView`, `TreeView`に設定したアイテム
- `OptionButton`, `CheckBox`, `ToggleButton`の選択状態
- `BackStyle`に設定した透明表示設定 (`Label`, `TextBox`, `CommandButton`, `CheckBox`, `ToggleButton`, `OptionButton`, `Image`, `ComboBox`)
- コントロールに埋め込まれた画像 (`Image`)
- `.TabOrientation`プロパティ (`MultiPage`)
- `.Locked`プロパティ (`TextBox`, `ListBox`, `ComboBox`)
- `.PasswordChar`プロパティ (`TextBox` [.MultiLine=False])
- `.ScrollBars` プロパティ (`TextBox` [.MultiLine=True])
- `.WordWrap` プロパティ (`TextBox` [.MultiLine=True])
- `.Style`プロパティ (`ComboBox`, `MultiPage`)
- `.MultiSelect`プロパティ (`ListBox`, `ListView`)
- `.PictureAlignment`プロパティ (`Image`)
- `.Scroll` プロパティ (`TreeView`)
- `.Expanded` プロパティ (`TreeView.Nodes`)

※ `.BackStyle = fmBackStyleTransparent`の場合、Tkinterではウィジェットの背景色の透過がサポートされていないため以下のように変換されます

-   親コントロールが`.BackColor`プロパティを持つ場合、その色を`.BackColor`に設定します
-   親コントロールが`Page`の場合、`.BackColor`プロパティを持たないため`Page`の視覚的な背景色と一致するシステムカラーの`&H8000000F&`を`.BackColor`に設定します

## 対応しているコントロールの種類
| VBA Formのクラス | Tkinterのクラス|
| ------ | ------ |
| `Label` | `tk.Label` |
| `CommandButton` | `tk.Button` |
| `Frame` (Captionなし) | `tk.Frame` |
| `Frame` (Captionあり) | `tk.LabelFrame` |
| `TextBox` (`MultiLine=False`) | `tk.Entry` |
| `TextBox` (`MultiLine=True`) | `tk.Text` |
| `SpinButton` | `tk.Spinbox` |
| `ListBox` | `tk.Listbox` |
| `CheckBox` | `tk.Checkbutton` |
| `ToggleButton` | `tk.Checkbutton`(`indicatoron=0`) |
| `OptionButton` | `tk.Radiobutton` |
| `Image` | `tk.Canvas` |
| `ScrollBar` | `ttk.Scale` |
| `ComboBox` | `ttk.Combobox` |
| `MultiPage` | `ttk.Notebook` |
| `ListView`(`.View=lvwReport`) | `ttk.Treeview`(`show="headings"`) |
| `TreeView` | `ttk.Treeview`(`show="tree"`) |

※`SpinButton`は仕様が異なるため、配置方法によっては外観が異なります<br>
※`ScrollBar`についてはVBAのフォームには上下調整用のボタンがありますがTkinterの`Scale`にはありません<br>
<br>
上記以外のコントロールがフォーム上にある場合、変換に失敗するので該当のコントロールを削除したうえで再度変換を行ってください<br>

## 使い方
使用前に、変換したいユーザーフォームが作成されたExcelブックを用意する必要があります<br>
また、VBE上でイミディエイトウィンドウが表示されていない場合は表示の設定を行ってください<br><br>
<img width="843" height="768" alt="Image" src="https://github.com/user-attachments/assets/676cd54c-d610-4c25-bd9a-9e064e38dc5e" /><br><br>
1.[ここ](https://github.com/GUI-Conversion-Tools/VBAForm2Tkinter/releases)から最新版のファイルをダウンロードし解凍してください、中の`VBAForm2Tkinter.bas`を使用します<br>
2. Excelの開発→Visual BasicからVBEを開いてください<br>
3. プロジェクトを右クリックし、「ファイルのインポート」より`VBAForm2Tkinter.bas`をインポートします<br>
4. イミディエイトウィンドウに`Call ConvertForm2Tkinter(UserForm1)`と入力しEnterキーを押下します<br>
```vb
Call ConvertForm2Tkinter(UserForm1)
```
※`UserForm1`の部分は変換したいユーザーフォームのオブジェクト名に変えてください<br>
5. 正常に変換が完了した場合、メッセージが表示され`output.py`が作成されます<br>
6. GUIの外観を確認したら、pyファイルを編集し`.mainloop()`の上に`コントロール名.configure()`でボタン押下時の関数の設定などをしてください<br>

## 出力先フォルダ

`VBAForm2Tkinter_output`フォルダがワークブックと同じディレクトリに生成され、すべての生成されたファイルはこのフォルダ内に格納されます

### ExcelまたはWord

ExcelまたはWordの場合出力先フォルダはマクロを含んだファイルと同じディレクトリに作成されます

-   **Excel**: ワークブックが配置されたディレクトリ (`ThisWorkbook.Path`)
-   **Word**: ドキュメントが配置されたディレクトリ (`MacroContainer.Path`)

```
ワークブックが配置されたフォルダ/
├─ MyWorkbook.xlsm
└─ VBAForm2Tkinter_output/
   ├─ output.py
   ├─ image_base64.json
   └─ エクスポートされた画像ファイル...
```

### 他のOfficeの場合
他のOffice (PowerPoint, Outlookなど)で実行する場合または未保存のExcelブック/Wordドキュメント下で実行する場合、出力先フォルダは**ドキュメントフォルダ**に作成されます.

```
C:\Users\%USERNAME%\Documents\
└─ VBAForm2Tkinter_output/
   ├─ output.py
   ├─ image_base64.json
   └─ エクスポートされた画像ファイル...
```

ドキュメントフォルダの取得に失敗した場合は、Cドライブ直下に出力先フォルダが作成されます

## 引数

`ConvertForm2Tkinter`には以下の引数を設定できます:

|**引数**|**型**|**説明**                         |
|----------------|-------------------------------|-----------------------------|
|`frms` |`Variant`|**必須**<br>変換対象の`UserForm`オブジェクトまたは`UserForm`オブジェクトの配列を指定 |
|`useCls`  |`Boolean` |**省略可能 (デフォルト: `False`)** <br>`True`にした場合生成したPythonコードにおいて各フォームをクラス化する &nbsp;&nbsp;`frms`が配列の場合は自動的に`True`に設定される|
|`noMainLoop`  |`Boolean`|**省略可能 (デフォルト: `False`)** <br>`True`にした場合生成したPythonコードに`.mainloop()`を含めなくする &nbsp;&nbsp;`useCls`が`True`の場合はインスタンスの作成(例:`obj_UserForm1 = UserForm1()`)もスキップする|
|`uniqueStyleName`  |`Boolean`|**省略可能 (デフォルト: `True`)**<br>`True`(デフォルト)にした場合、UUIDベースのユニークな識別子をttkの各スタイル名に付与する、これは同じ種類のフォームやウィジェットが複数変換され、同じPython環境内で実行された場合のスタイル名衝突を防ぐ役割を持つ |
|`imageMode`  |`String` |**省略可能 (デフォルト: `"file"`)**<br>変換時の画像ファイルの扱いを設定する、以下の値を設定可能:<br>• `"file"` (デフォルト): 画像は出力先フォルダ内に個別の画像ファイルとして保存され、生成されたコードはそれらの画像を参照する<br>• `"disabled"`: 画像の処理そのものを無効化し、生成されたコード内でも画像の参照設定を行わない<br>• `"reference-only"`: `"file"`と同様、画像を参照するコードを生成するが画像の出力自体はスキップする 既に画像ファイルが存在する場合に有用<br>• `"base64"`: 画像をBase64文字列としてコード内に直接埋め込み単一ファイルに収める<br>• `"base64-dict"`: 画像のBase64文字列をコード内の`dict`に格納する<br>• `"base64-json"`: 画像のBase64文字列を外部の`image_base64.json`内に格納する、生成されたコードはそのJSONファイルを参照する<br>• `"base64-json-reference"`: `"base64-json"`と同様、`image_base64.json`を参照するコードを生成するがJSONファイルの生成自体はスキップする JSONファイルが既に存在する場合に有用|

`ConvertForm2Tkinter`は単一のユーザーフォームまたは配列内の複数のユーザーフォームを変換することが可能です

```vb
' 実行例: 単一のフォームを変換
Call ConvertForm2Tkinter(UserForm1)

' 実行例: 単一のフォームを変換 (クラス化を行う)
Call ConvertForm2Tkinter(UserForm1, useCls:=True)

' 実行例: 複数のフォームを変換 (自動的にクラス化される)
Call ConvertForm2Tkinter(Array(UserForm1, UserForm2))

' 実行例: 単一のフォームを変換 (画像をBase64文字列としてコード内に直接埋め込む)
Call ConvertForm2Tkinter(UserForm1, imageMode:="base64")
```

## 子要素を設定できないコントロールの並び順について
Tkinterでは例として`Label`に`Label`を重ねた場合は設置した順番が後のものが優先して前面に表示されます<br>
ただしVBAのユーザーフォームにおいては前面/背面を変更することができるためこの限りではありません<br>
このプログラムは各コントロールを階層順にソート後、同じ階層のものについては元々の設置順に従いウィジェットを配置します<br>
現状コントロールのZオーダー(前面/背面情報)を取得できる手段がないため反映させることができずVBAでの表示と異なってしまう場合があります<br>
その場合は、Pythonのコードを編集し、前面に表示したいものを後に設置するか、VBA側でコントロールを配置し直すことで順番を後にしてください<br>
なお、新規でGUIを作成する場合は重ねるよりも`Frame`などの明確な親子関係を設定可能なコントロールを使用することを推奨します<br>
