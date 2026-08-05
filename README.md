# VBAForm2Tkinter - VBA UserForm to Tkinter Converter
:jp:[日本語の説明はこちら](https://github.com/GUI-Conversion-Tools/VBAForm2Tkinter/blob/main/README_ja.md)<br><br>
This program converts userforms created in Microsoft Office VBA into Python Tkinter code.<br>

## Example
<img width="681" height="1275" alt="Image" src="https://github.com/user-attachments/assets/1ff1765a-9c84-4b5d-b5b1-29abf2ba40f5" /><br>
<img width="704" height="695" alt="Image" src="https://github.com/user-attachments/assets/ca514378-3017-443f-a3d8-bbd1ed4ceeb6" /><br>

## System Requirements
- Supported OS: Windows XP or later
- Required Software: Microsoft Excel/Word/PowerPoint/Outlook 2000 or later
- Recommended Environment: Microsoft Excel 2016 or later
> Note:  
> Python 3.4 or later is required to run the generated code.

## Verified Operating Environments
- Windows XP(SP3)
- Windows 10/11
- Excel 2000(32bit)
- Excel 2010(32bit)
- Excel 2016(32bit)
- Excel 2019(64bit)
- Word/PowerPoint/Outlook 2000 (32bit)
- Word/PowerPoint/Outlook 2019 (64bit)

## Converted Elements
- Variable names (object names)
- Approximate layout and size of controls
- Control colors (foreground)
- Control colors (background)
- Text display (`Label`, `CommandButton`, `CheckBox`, `ToggleButton`, `OptionButton`, `MultiPage`)
- Font (typeface, size, bold, italic)
- Borders (`UserForm`, `Frame`, `TextBox`, `Label`, `ListBox`, `Image`)
- Mouse cursor
- Text alignment: left, center, right (`Label`, `TextBox` [.MultiLine=False], `ComboBox`, `CheckBox`, `ToggleButton`, `OptionButton`, `ListView`)
- Default values of `TextBox`, `ComboBox`
- Items set in `ComboBox`, `ListBox`, `ListView`, `TreeView`
- Selection state of `OptionButton`, `CheckBox` and `ToggleButton`
- Transparent background setting specified in `.BackStyle`(`Label`, `TextBox`, `CommandButton`, `CheckBox`, `ToggleButton`, `OptionButton`, `Image`, `ComboBox`)
- Images Embedded in Controls (`Image`)
- `.TabOrientation` property (`MultiPage`)
- `.Locked` property (`TextBox`, `ListBox`, `ComboBox`)
- `.PasswordChar` property (`TextBox` [.MultiLine=False])
- `.ScrollBars` property (`TextBox` [.MultiLine=True])
- `.WordWrap` property (`TextBox` [.MultiLine=True])
- `.Style` property (`ComboBox`, `MultiPage`)
- `.MultiSelect` property (`ListBox`, `ListView`)
- `.PictureAlignment` property (`Image`)
- `.Scroll` property (`TreeView`)
- `.Expanded` property (`TreeView.Nodes`)

> Note:  
> - When `.BackStyle = fmBackStyleTransparent`, true transparency is not supported in Tkinter. Instead, the background color is substituted as follows:
>   -   If the parent control has a `.BackColor`, that color is used.
>   -   If the parent is a `Page` (which does not expose `.BackColor`), a system default color (`&H8000000F&`) is used as a fallback, which matches the visual background color of the `Page`.
> 
> - Since Tkinter does not support displaying resized images natively, when exporting images with the `.PictureSizeMode` property set to `fmPictureSizeModeStretch` or `fmPictureSizeModeZoom`, the embedded images are exported after being resized to match the control's dimensions. This preserves the appearance of the original VBA UserForm as closely as possible.

## Supported Controls
| VBA Form Class | Tkinter Class|
| ------ | ------ |
| `Label` | `tk.Label` |
| `CommandButton` | `tk.Button` |
| `Frame` (without Caption) | `tk.Frame` |
| `Frame` (with any Caption) | `tk.LabelFrame` |
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

> Note:
>- `SpinButton` behaves differently in VBA and Tkinter, so appearance may vary depending on placement.
>- `ScrollBar` in VBA has up/down adjustment buttons, but Tkinter’s `Scale` does not.

If unsupported controls exist on the form, the conversion will fail. If that case, please remove those controls and run the conversion again.



## Usage
Before using, prepare the Excel workbook containing the user form you want to convert.
Also, ensure that the Immediate Window is visible in the VBE (Visual Basic Editor).<br><br>
<img width="843" height="768" alt="Image" src="https://github.com/user-attachments/assets/676cd54c-d610-4c25-bd9a-9e064e38dc5e" /><br><br>
1. Download the latest file from [here](https://github.com/GUI-Conversion-Tools/VBAForm2Tkinter/releases) and extract it. Use the `VBAForm2Tkinter.bas` file inside.<br>
2. In Excel, go to Developer -> Visual Basic to open VBE.<br>
3. Right-click your project and import the provided `.bas` file using Import File.<br>
4. In the Immediate Window, enter: `Call ConvertForm2Tkinter(UserForm1)`<br>
```vb
Call ConvertForm2Tkinter(UserForm1)
```
   > Note: Replace `UserForm1` with the object name of the form you want to convert.

5.  If conversion succeeds, a message will appear, and an `output.py` file will be created.<br>
6.  After checking the GUI appearance, edit the `.py` file and, above `.mainloop()`, configure event handlers for controls (e.g., `button.configure(command=...)`).<br>

## Output Directory

A dedicated `VBAForm2Tkinter_output` folder is automatically created in the workbook directory, and all generated files are saved there:

### Excel and Word

When running from Excel or Word, the output folder is created in the same directory as the macro-enabled document.

-   **Excel**: Uses the workbook's directory (`ThisWorkbook.Path`)
-   **Word**: Uses the document's directory (`MacroContainer.Path`)

```
WorkbookFolder/
├─ MyWorkbook.xlsm
└─ VBAForm2Tkinter_output/
   ├─ output.py
   ├─ image_base64.json
   └─ exported images...
```

### Other Office Applications
When running from other Office applications (such as PowerPoint, Outlook, etc.), or when the current Excel workbook or Word document has not yet been saved, the output folder is created in the user's **Documents** folder instead.

```
C:\Users\%USERNAME%\Documents\
└─ VBAForm2Tkinter_output/
   ├─ output.py
   ├─ image_base64.json
   └─ exported images...
```

If the Documents folder cannot be resolved, the output folder will be created in the root of the **C:** drive as a final fallback.

## Parameters

`ConvertForm2Tkinter` accepts the following parameters:

|**Parameter**|**Type**|**Description**                         |
|----------------|-------------------------------|-----------------------------|
|`frms` |`Variant`|**Required.**<br>Accepts a single `UserForm` object or an `Array` of `UserForm` objects to be converted.            |
|`useCls`  |`Boolean` |**Optional (Default: `False`).**<br>If set to `True`, the generated Python code will wrap each form in a Python class structure. This is automatically set to `True` if `frms` is an array.|
|`noMainLoop`  |`Boolean`|**Optional (Default: `False`).**<br>If set to `True`, the `.mainloop()` call will be omitted from the end of the generated Python script. When `useCls` is also `True`, this will additionally skip the code that creates the object instances (e.g., `obj_UserForm1 = UserForm1()`).|
|`uniqueStyleName`  |`Boolean`|**Optional (Default: `True`).**<br>If set to `True` (default), a unique suffix (UUID-based) will be appended to each ttk style name. This prevents styling conflicts when multiple forms or widgets of the same type are converted and run in the same Python environment.|
|`imageMode`  |`String` |**Optional (Default: `"file"`).**<br>Determines how image files used in the UserForm are handled during conversion. You can choose one of the following options:<br>• `"file"` (Default): Images are saved as separate external files in the output directory, and the generated code references these files.<br>• `"disabled"`: Image processing is disabled, and no image-related code is generated.<br>• `"reference-only"`: Similar to `"file"`, generates code that references image files, but does not export the image files. Useful when the image files already exist.<br>• `"base64"`: Images are embedded directly into the generated code as Base64-encoded strings, keeping everything in a single file.<br>• `"base64-dict"`: Images are embedded as Base64 strings within a `dict` inside the generated code.<br>• `"base64-json"`: Images are stored in an external `image_base64.json` file as Base64 strings, and the generated code references the JSON file.<br>• `"base64-json-reference"`: Similar to `"base64-json"`, generates code that references `image_base64.json`, but does not export the JSON file. Useful when the JSON file already exists.|

You can execute the conversion by calling the `ConvertForm2Tkinter` with a single UserForm object or an array of multiple UserForms.

```vb
' Example: Converting a single form
Call ConvertForm2Tkinter(UserForm1)

' Example: Converting a single form (Class-based style)
Call ConvertForm2Tkinter(UserForm1, useCls:=True)

' Example: Converting multiple forms (Automatically uses Class-based style)
Call ConvertForm2Tkinter(Array(UserForm1, UserForm2))

' Example: Converting a single form (With image streams embedded directly as Base64 text)
Call ConvertForm2Tkinter(UserForm1, imageMode:="base64")
```

## Control Order (for Controls Without Child Elements)
In Tkinter, if you place one `Label` on top of another, the later widget appears in front.<br>
However, in VBA, you can change front/back order, so the behavior differs.<br>
The program first sorts controls by hierarchy level; however, it preserves the original creation order within the same hierarchy.<br>
Since VBA’s z-order (front/back) cannot currently be retrieved, some displays may not match VBA.<br>

To adjust:<br>
&nbsp;&nbsp;&nbsp;&nbsp;Edit the Python code so the widget you want in front is placed later, or Reorder the controls in VBA before conversion.<br>
&nbsp;&nbsp;&nbsp;&nbsp;For new GUIs, instead of overlapping controls, it is recommended to use containers like `Frame`, which allow clear parent-child relationships.
