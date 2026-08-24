Attribute VB_Name = "VBAForm2Tkinter"

' VBAForm2Tkinter v1.5.2
' https://github.com/GUI-Conversion-Tools/VBAForm2Tkinter
' Copyright (c) 2025-2026 ZeeZeX
' This software is released under the MIT License.
' https://opensource.org/licenses/MIT

Option Explicit


#If VBA7 = 0 Then
    ' Define a placeholder LongPtr type for VBA6 or earlier.
    ' This allows code that uses LongPtr to compile in older VBA versions,
    ' even though those versions do not natively support the LongPtr type.
    Private Enum LongPtr
        [_]
    End Enum
#End If

#If VBA7 Then
    ' 64bit Office / VBA7 or later
    Private Declare PtrSafe Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
    
    Private Declare PtrSafe Function GdiplusStartup Lib "gdiplus" (ByRef token As LongPtr, ByRef inputbuf As GDIPlusStartupInput, Optional ByVal outputbuf As LongPtr = 0) As Long
    Private Declare PtrSafe Sub GdiplusShutdown Lib "gdiplus" (ByVal token As LongPtr)
    Private Declare PtrSafe Function GdipLoadImageFromFile Lib "gdiplus" (ByVal filename As LongPtr, ByRef image As LongPtr) As Long
    Private Declare PtrSafe Function GdipDisposeImage Lib "gdiplus" (ByVal image As LongPtr) As Long
    Private Declare PtrSafe Function GdipGetImageWidth Lib "gdiplus" (ByVal image As LongPtr, ByRef width As Long) As Long
    Private Declare PtrSafe Function GdipGetImageHeight Lib "gdiplus" (ByVal image As LongPtr, ByRef height As Long) As Long
    Private Declare PtrSafe Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal width As Long, ByVal height As Long, ByVal stride As Long, ByVal format As Long, ByVal scan0 As LongPtr, ByRef bitmap As LongPtr) As Long
    Private Declare PtrSafe Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal image As LongPtr, ByRef graphics As LongPtr) As Long
    Private Declare PtrSafe Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As LongPtr) As Long
    Private Declare PtrSafe Function GdipSetInterpolationMode Lib "gdiplus" (ByVal graphics As LongPtr, ByVal Mode As Long) As Long
    Private Declare PtrSafe Function GdipDrawImageRectI Lib "gdiplus" (ByVal graphics As LongPtr, ByVal image As LongPtr, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long) As Long
    Private Declare PtrSafe Function GdipSaveImageToFile Lib "gdiplus" (ByVal image As LongPtr, ByVal filename As LongPtr, ByRef clsidEncoder As GUID, ByVal encoderParams As LongPtr) As Long
    Private Declare PtrSafe Function CLSIDFromString Lib "ole32" (ByVal lpsz As LongPtr, ByRef pclsid As GUID) As Long
    
    Private Declare PtrSafe Function CryptBinaryToStringW Lib "crypt32" (ByVal pbBinary As LongPtr, ByVal cbBinary As Long, ByVal dwFlags As Long, ByVal pszString As LongPtr, ByRef pcchString As Long) As Long
#Else
    ' 32bit Office
    Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
    
    Private Declare Function GdiplusStartup Lib "gdiplus" (ByRef token As Long, ByRef inputbuf As GDIPlusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
    Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal token As Long)
    Private Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal filename As Long, ByRef image As Long) As Long
    Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal image As Long) As Long
    Private Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal image As Long, ByRef width As Long) As Long
    Private Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal image As Long, ByRef height As Long) As Long
    Private Declare Function GdipCreateBitmapFromScan0 Lib "gdiplus" (ByVal width As Long, ByVal height As Long, ByVal stride As Long, ByVal format As Long, ByVal scan0 As Long, ByRef bitmap As Long) As Long
    Private Declare Function GdipGetImageGraphicsContext Lib "gdiplus" (ByVal image As Long, ByRef graphics As Long) As Long
    Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal graphics As Long) As Long
    Private Declare Function GdipSetInterpolationMode Lib "gdiplus" (ByVal graphics As Long, ByVal Mode As Long) As Long
    Private Declare Function GdipDrawImageRectI Lib "gdiplus" (ByVal graphics As Long, ByVal image As Long, ByVal x As Long, ByVal y As Long, ByVal width As Long, ByVal height As Long) As Long
    Private Declare Function GdipSaveImageToFile Lib "gdiplus" (ByVal image As Long, ByVal filename As Long, ByRef clsidEncoder As GUID, ByVal encoderParams As Long) As Long
    Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpsz As Long, ByRef pclsid As GUID) As Long
    
    Private Declare Function CryptBinaryToStringW Lib "crypt32" (ByVal pbBinary As Long, ByVal cbBinary As Long, ByVal dwFlags As Long, ByVal pszString As Long, ByRef pcchString As Long) As Long
#End If

Private Type GDIPlusStartupInput: GdiPlusVersion As Long: DebugEventCallback As LongPtr: SuppressBackgroundThread As Long: SuppressExternalCodecs As Long: End Type
Private Type GUID: Data1 As Long: Data2 As Integer: Data3 As Integer: Data4(0 To 7) As Byte: End Type

Private Const FORM_WINDOW_NAME As String = "window"
Private Const OUTPUT_FOLDER_NAME As String = "VBAForm2Tkinter_output"

Public Sub TestRunConversion2Tk()
    Call ConvertForm2Tkinter(UserForm1)
End Sub

Public Sub TestRunConversion2Tk_2()
    Call ConvertForm2Tkinter(Array(UserForm1, UserForm2))
End Sub

Public Sub ConvertForm2Tkinter(ByVal frms As Variant, Optional ByVal useCls As Boolean = False, Optional ByVal noMainLoop As Boolean = False, Optional ByVal uniqueStyleName As Boolean = True, Optional ByVal imageMode As String = "file")
    
    ' frms: Variant
    '   Accepts a single UserForm object or an Array of UserForm objects to be converted.
    ' useCls: Boolean
    '   If set to True, the generated Python code will wrap each form in a Python class structure.
    '   This is automatically set to True if frms is an array.
    ' noMainLoop: Boolean
    '   If set to True, the .mainloop() call will be omitted from the end of the generated Python script.
    '   When useCls is also True, this will additionally skip the code that creates the object instances (e.g., obj_UserForm1 = UserForm1()).
    ' uniqueStyleName: Boolean
    '   If set to True (default), a unique suffix (UUID-based) will be appended to each ttk style name.
    '   This prevents styling conflicts when multiple forms or widgets of the same type are converted and run in the same Python environment.
    ' imageMode: String
    '   Determines how image files used in the UserForm are handled during conversion. You can choose one of the following options:
    '     "file" (Default): Images are saved as separate external files in the output directory, and the generated code references these files.
    '     "disabled": Image processing is disabled, and no image-related code is generated.
    '     "reference-only": Similar to "file", generates code that references image files, but does not export the image files. Useful when the image files already exist.
    '     "base64": Images are embedded directly into the generated code as Base64-encoded strings, keeping everything in a single file.
    '     "base64-dict": Images are embedded as Base64 strings within a dict inside the generated code.
    '     "base64-json": Images are stored in an external image_base64.json file as Base64 strings, and the generated code references the JSON file.
    '     "base64-json-reference": Similar to "base64-json", generates code that references image_base64.json, but does not export the JSON file. Useful when the JSON file already exists.
    
    Dim code As String
    Dim filePath As String
    Dim saveDir As String
    code = GenerateTkinterCode(frms, useCls, noMainLoop, uniqueStyleName, imageMode)
    If code <> "" Then
        saveDir = GetSaveDirPath()
        Call CreateFolderIfDoesNotExist(saveDir)
        filePath = saveDir & "\output.py"
        Call SaveUtf8TextNoBom(filePath, code)
        MsgBox "Saved: " & filePath
    Else
        MsgBox "Conversion failed."
    End If
    
End Sub


Public Function GenerateTkinterCode(ByVal frms As Variant, Optional ByVal useCls As Boolean = False, Optional ByVal noMainLoop As Boolean = False, Optional ByVal uniqueStyleName As Boolean = True, Optional ByVal imageMode As String = "file") As String
    Dim root As Variant
    Dim indent As String
    Dim prefix As String
    Dim clsNumber As Long
    Dim formName As String
    Dim controlVarName As String
    Dim parentVarName As String
    Dim childVarName As String
    Dim itemsListName As String
    Dim tkStyleBaseName As String
    Dim instanceName As String
    Dim toplevelInstanceName As String
    Dim unavailableNames() As Variant
    Dim ctrl As MSForms.Control
    Dim ctrls As Collection
    Dim item As Variant
    Dim result As String
    Dim codeList As New Collection
    Const q As String = """"
    Dim fontStyle As String
    Dim fontOpts As String
    Dim styleName As String
    Dim pixelWidth As Long
    Dim pixelHeight As Long
    Dim pixelTop As Long
    Dim pixelLeft As Long
    Dim i As Long
    Dim orientation As String
    Dim cursorType As String
    Dim caption As String
    Dim colorCode As String
    Dim ttkStyleRef As String
    Dim tabPosition As String
    Dim canvasCoordX As String
    Dim canvasCoordY As String
    Dim canvasAnchor As String
    Dim listviewHeaderNames As String
    Dim ttkFontSetting As String
    Dim rowHeight As Double
    Dim rowPixelHeight As Long
    Dim treeviewNodes As Collection
    Dim node As Object
    Dim nodeDictName As String
    Dim nodeVarName As String
    Dim nodeParentVarName As String
    Dim enableScrollBar As Boolean
    Dim hasPicture As Boolean
    Dim tempPath As String
    Dim picturePath As String
    Dim pictureName As String
    Dim pictureWidth As Long
    Dim pictureHeight As Long
    Dim resizePicture As Boolean
    Dim preservePictureAspectRatio As Boolean
    Dim saveDir As String
    Dim uniqueStringToReplace As String
    Dim base64PictureDictStr As String
    Dim base64PictureDictStrs As New Collection
    Dim base64Str As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim supportedImageModeValues() As Variant
    
    imageMode = LCase$(imageMode)
    supportedImageModeValues = Array("file", "disabled", "reference-only", "base64", "base64-dict", "base64-json", "base64-json-reference")
    If Not ContainsValue(supportedImageModeValues, imageMode) Then
        MsgBox "[imageMode] Invalid value: " & q & imageMode & q & vbLf & "Supported values are " & q & Join(supportedImageModeValues, q & ", " & q) & q
        GenerateTkinterCode = ""
        Exit Function
    End If
    
    saveDir = GetSaveDirPath()
    Call CreateFolderIfDoesNotExist(saveDir)
    
    
    If IsArray(frms) Then
        useCls = True
    Else
        frms = VBA.Array(frms)
    End If
    
    If useCls Then
        indent = "        "
        prefix = "self."
    Else
        indent = ""
        prefix = ""
    End If
    
    codeList.Add "import tkinter as tk" & vbLf
    codeList.Add "from tkinter import ttk" & vbLf
    codeList.Add "from tkinter import font" & vbLf
    codeList.Add "import pathlib" & vbLf
    If InStr(imageMode, "json") > 0 Then codeList.Add "import json" & vbLf
    codeList.Add vbLf
    codeList.Add "BASE_DIR = pathlib.Path(__file__).resolve().parent" & vbLf
    
    uniqueStringToReplace = GenerateUUIDv4() & GenerateUUIDv4() & GenerateUUIDv4() & GenerateUUIDv4()
    codeList.Add "" & vbLf
    codeList.Add uniqueStringToReplace
    
    
    If ContainsValue(Array("base64-json", "base64-json-reference"), imageMode) Then
        ' Compatibility note: In Python 3.5 and earlier, pathlib.Path objects cannot be passed directly to open(). Convert them to strings using str() before passing them.
        codeList.Add "with open(str(BASE_DIR / r""image_base64.json""), ""r"", encoding=""utf-8"") as f:" & vbLf
        codeList.Add "    image_base64_dict = json.load(f)" & vbLf & vbLf
    End If
    
    
    For Each root In frms
        unavailableNames = VBA.Array("", "tk", "ttk", "font", "style", "int", "item", "image_base64_dict", "base_dir")
        
        For i = LBound(unavailableNames) To UBound(unavailableNames)
            unavailableNames(i) = LCase$(unavailableNames(i))
        Next
        
        If ContainsValue(unavailableNames, LCase$(root.Name)) Then
            MsgBox GenerateUnavailableNameMessage(root)
            result = ""
            GenerateTkinterCode = result
            Exit Function
        End If
        unavailableNames(0) = LCase$(FORM_WINDOW_NAME)
        
        ' Convert UserForm's size to pixel size
        pixelWidth = UserFormSizeToPixel(root.InsideWidth)
        pixelHeight = UserFormSizeToPixel(root.InsideHeight)
        
        formName = GenerateCtrlVarName(root, prefix, useCls)
        
        If useCls Then
            codeList.Add "class " & root.Name & ":" & vbLf & "    " & "def __init__(self, parent=None):" & vbLf
            codeList.Add indent & "if parent is None:" & vbLf & _
            indent & "    " & formName & " = " & "tk.Tk()" & vbLf & _
            indent & "else:" & vbLf & _
            indent & "    " & formName & " = " & "tk.Toplevel(parent)" & vbLf
        Else
            codeList.Add indent & formName & " = " & "tk.Tk()" & vbLf
        End If
        
        
        caption = root.caption
        caption = Convert2PythonFormatText(caption)
        codeList.Add indent & formName & ".title(" & q & caption & q & ")" & vbLf
        codeList.Add indent & formName & ".geometry(" & q & pixelWidth & "x" & pixelHeight & q & ")" & vbLf
        codeList.Add indent & formName & ".resizable(False, False)" & vbLf
        codeList.Add indent & formName & ".configure(bg=" & q & FormColorToHex(root.BackColor) & q & ")" & vbLf
        codeList.Add indent & formName & GetBorderSetting(root) & vbLf
        
        cursorType = GetControlCursorType(root)
        If cursorType <> "" Then
            codeList.Add indent & formName & ".configure(cursor=" & q & cursorType & q & ")" & vbLf
        End If
        
        codeList.Add vbLf
        codeList.Add indent & prefix & "style = ttk.Style()" & vbLf
        codeList.Add indent & prefix & "style.theme_use('default')" & vbLf
        codeList.Add vbLf
        Set ctrls = SortFormControlsByDepth(root.Controls)
        For Each ctrl In ctrls
            controlVarName = GenerateCtrlVarName(ctrl, prefix, useCls)
            parentVarName = GenerateCtrlVarName(ctrl.Parent, prefix, useCls)
            itemsListName = controlVarName & "_items_value"
            listviewHeaderNames = controlVarName & "_listview_headers"
            enableScrollBar = False
            hasPicture = False
            pictureWidth = -1
            pictureHeight = -1
            resizePicture = False
            preservePictureAspectRatio = False
            base64Str = ""
            
            If useCls Then
                pictureName = "img_" & root.Name & "_" & ctrl.Name
            Else
                pictureName = "img_" & ctrl.Name
            End If
            
            
            ' Generate unique style name to prevent naming conflicts.
            If uniqueStyleName Then
                tkStyleBaseName = GenerateCtrlVarName(ctrl, "", False) & "." & Left$(GenerateUUIDv4, 8) & ".style"
            Else
                tkStyleBaseName = GenerateCtrlVarName(ctrl, "", False) & ".style"
            End If
            
            ttkStyleRef = GenerateTtkStyleRef(controlVarName)
            
            If ContainsValue(unavailableNames, LCase$(ctrl.Name)) Then
                MsgBox GenerateUnavailableNameMessage(ctrl)
                result = ""
                GenerateTkinterCode = result
                Exit Function
            End If
            
            If GetTkWidgetName(ctrl) <> "" Then
                
                pixelLeft = UserFormSizeToPixel(ctrl.Left)
                pixelTop = UserFormSizeToPixel(ctrl.Top)
                pixelWidth = UserFormSizeToPixel(ctrl.width)
                pixelHeight = UserFormSizeToPixel(ctrl.height)
                
                
                codeList.Add indent & controlVarName & " = " & GetTkWidgetName(ctrl) & "(" & parentVarName & ")" & vbLf
                codeList.Add indent & controlVarName & ".place(x=" & pixelLeft & ", y=" & pixelTop & ", width=" & pixelWidth & ", height=" & pixelHeight & ")" & vbLf
                
                If GetTkWidgetName(ctrl) = "tk.LabelFrame" Or ContainsValue(Array("Label", "CommandButton", "TextBox", "SpinButton", "ListBox", "CheckBox", "ToggleButton", "OptionButton"), TypeName(ctrl)) Then
                    ' Set ForeColor
                    codeList.Add indent & controlVarName & ".configure(fg=" & q & FormColorToHex(ctrl.ForeColor) & q & ")" & vbLf
                End If
                
                If ContainsValue(Array("Label", "CommandButton", "Frame", "TextBox", "SpinButton", "ListBox", "CheckBox", "ToggleButton", "OptionButton", "Image"), TypeName(ctrl)) Then
                    ' Set BackColor
                    colorCode = FormColorToHex(ctrl.BackColor)
                    If ContainsValue(Array("Label", "TextBox", "CommandButton", "CheckBox", "ToggleButton", "OptionButton", "Image"), TypeName(ctrl)) Then
                        ' If the BackStyle is set to Transparent, apply the BackColor of the parent control
                        If ctrl.BackStyle = fmBackStyleTransparent Then
                            If TypeName(ctrl.Parent) <> "Page" Then
                                colorCode = FormColorToHex(ctrl.Parent.BackColor)
                            Else
                                ' Because the Page control does not have a BackColor property, set the color to &H8000000F&, which matches the background color of the Page
                                colorCode = FormColorToHex(&H8000000F)
                            End If
                        End If
                    End If
                    codeList.Add indent & controlVarName & ".configure(bg=" & q & colorCode & q & ")" & vbLf
                    
                    If ContainsValue(Array("CommandButton", "CheckBox", "ToggleButton", "OptionButton"), TypeName(ctrl)) Then
                        ' Set the colors when the button is pressed
                        codeList.Add indent & controlVarName & ".configure(activeforeground=" & q & FormColorToHex(ctrl.ForeColor) & q & ")" & vbLf
                        codeList.Add indent & controlVarName & ".configure(activebackground=" & q & colorCode & q & ")" & vbLf
                    End If
                    
                    If TypeName(ctrl) = "ToggleButton" Then
                        codeList.Add indent & controlVarName & ".configure(indicatoron=0)" & vbLf
                        codeList.Add indent & controlVarName & ".configure(selectcolor=" & q & colorCode & q & ")" & vbLf
                    End If
                    
                End If
                
                If GetTkWidgetName(ctrl) = "tk.LabelFrame" Or ContainsValue(Array("Label", "CommandButton", "CheckBox", "ToggleButton", "OptionButton"), TypeName(ctrl)) Then
                    caption = ctrl.caption
                    caption = Convert2PythonFormatText(caption)
                    codeList.Add indent & controlVarName & ".configure(text=" & q & caption & q & ")" & vbLf
                End If
                
                
                If TypeName(ctrl) = "TextBox" Then
                    If GetTkWidgetName(ctrl) = "tk.Entry" Then
                        codeList.Add indent & controlVarName & ".insert(0, " & q & Convert2PythonFormatText(ctrl.text) & q & ")" & vbLf
                        If ctrl.PasswordChar <> "" Then
                            codeList.Add indent & controlVarName & ".configure(show=" & q & Left$(ctrl.PasswordChar, 1) & q & ")" & vbLf
                        End If
                        
                        If ctrl.Locked Then
                             codeList.Add indent & controlVarName & ".configure(state=" & q & "readonly" & q & ")" & vbLf
                        End If
                        
                    ElseIf GetTkWidgetName(ctrl) = "tk.Text" Then
                        codeList.Add indent & controlVarName & ".insert(" & q & "1.0" & q & ", " & q & Convert2PythonFormatText(ctrl.text) & q & ")" & vbLf
                        
                        If ctrl.Locked Then
                             codeList.Add indent & controlVarName & ".configure(state=" & q & "disabled" & q & ")" & vbLf
                        End If
                        
                        If Not ctrl.WordWrap Then
                            codeList.Add indent & controlVarName & ".configure(wrap=" & q & "none" & q & ")" & vbLf
                        End If
                        
                        Select Case ctrl.ScrollBars
                            Case fmScrollBarsHorizontal
                                codeList.Add SetHScrollBarToWidget(ctrl, indent, prefix, useCls, False)
                            Case fmScrollBarsVertical
                                codeList.Add SetVScrollBarToWidget(ctrl, indent, prefix, useCls, False)
                            Case fmScrollBarsBoth
                                codeList.Add SetHScrollBarToWidget(ctrl, indent, prefix, useCls, False)
                                codeList.Add SetVScrollBarToWidget(ctrl, indent, prefix, useCls, False)
                        End Select
                        
                    End If
                End If
                
                If TypeName(ctrl) = "ComboBox" Then
                    styleName = tkStyleBaseName & ".TCombobox"
                    codeList.Add indent & GenerateTtkStyleDefinitionCode(controlVarName, styleName, useCls) & vbLf
                    codeList.Add indent & prefix & "style.configure(" & ttkStyleRef & ", foreground=" & q & FormColorToHex(ctrl.ForeColor) & q & ")" & vbLf
                    colorCode = FormColorToHex(ctrl.BackColor)
                    
                    If ctrl.BackStyle = fmBackStyleTransparent Then
                        If TypeName(ctrl.Parent) <> "Page" Then
                            colorCode = FormColorToHex(ctrl.Parent.BackColor)
                        Else
                            colorCode = FormColorToHex(&H8000000F)
                        End If
                    End If
                    
                    codeList.Add indent & prefix & "style.configure(" & ttkStyleRef & ", fieldbackground=" & q & colorCode & q & ")" & vbLf
                    
                    codeList.Add indent & itemsListName & " = " & GetListBoxValue(ctrl, indent) & vbLf
                    codeList.Add indent & controlVarName & ".configure(value=" & itemsListName & ")" & vbLf
                    codeList.Add indent & controlVarName & ".set(" & q & Convert2PythonFormatText(ctrl.text) & q & ")" & vbLf
                    
                    If ctrl.Style = fmStyleDropDownList Then
                         codeList.Add indent & controlVarName & ".configure(state=" & q & "readonly" & q & ")" & vbLf
                         codeList.Add indent & prefix & "style.map(" & ttkStyleRef & ", foreground=[(" & q & "readonly" & q & ", " & q & FormColorToHex(ctrl.ForeColor) & q & ")]" & ")" & vbLf
                         codeList.Add indent & prefix & "style.map(" & ttkStyleRef & ", fieldbackground=[(" & q & "readonly" & q & ", " & q & colorCode & q & ")]" & ")" & vbLf
                         codeList.Add indent & prefix & "style.map(" & ttkStyleRef & ", selectforeground=[(" & q & "readonly" & q & ", " & q & FormColorToHex(ctrl.ForeColor) & q & ")]" & ")" & vbLf
                         codeList.Add indent & prefix & "style.map(" & ttkStyleRef & ", selectbackground=[(" & q & "readonly" & q & ", " & q & colorCode & q & ")]" & ")" & vbLf
                    End If
                    
                    If ctrl.Locked Then
                         codeList.Add indent & controlVarName & ".configure(state=" & q & "disabled" & q & ")" & vbLf
                    End If
                    
                End If
                
                
                If TypeName(ctrl) = "ListBox" Then
                    codeList.Add indent & itemsListName & " = " & GetListBoxValue(ctrl, indent) & vbLf
                    codeList.Add indent & controlVarName & ".insert(tk.END, " & "*" & itemsListName & ")" & vbLf
                    
                    Select Case ctrl.MultiSelect
                        Case fmMultiSelectMulti
                            codeList.Add indent & controlVarName & ".configure(selectmode=" & q & "multiple" & q & ")" & vbLf
                        Case fmMultiSelectExtended
                            codeList.Add indent & controlVarName & ".configure(selectmode=" & q & "extended" & q & ")" & vbLf
                    End Select
                    
                    If ctrl.Locked Then
                         codeList.Add indent & controlVarName & ".configure(state=" & q & "disabled" & q & ")" & vbLf
                         codeList.Add indent & controlVarName & ".configure(disabledforeground=" & q & FormColorToHex(ctrl.ForeColor) & q & ")" & vbLf
                    End If
                    
                    codeList.Add SetVScrollBarToWidget(ctrl, indent, prefix, useCls, True)
                    
                End If
                
                If TypeName(ctrl) = "ScrollBar" Then
                    Select Case ctrl.orientation
                        Case fmOrientationAuto
                            If ctrl.width > ctrl.height Then
                                orientation = "Horizontal"
                            Else
                                orientation = "Vertical"
                            End If
                            
                        Case fmOrientationVertical
                            orientation = "Vertical"
                        Case fmOrientationHorizontal
                            orientation = "Horizontal"
                        Case Else
                            orientation = "Vertical"
                    End Select
                    codeList.Add indent & controlVarName & ".configure(from_=" & ctrl.Min & ", to=" & ctrl.Max & ",orient=" & q & LCase$(orientation) & q & ")" & vbLf
                    styleName = tkStyleBaseName & "." & orientation & ".TScale"
                    codeList.Add indent & GenerateTtkStyleDefinitionCode(controlVarName, styleName, useCls) & vbLf
                    codeList.Add indent & prefix & "style.configure(" & ttkStyleRef & ", background=" & q & FormColorToHex(ctrl.BackColor) & q & ")" & vbLf
                    
                End If
                
                If IsListView(ctrl) Then
                    styleName = tkStyleBaseName & ".Treeview"
                    rowHeight = GetTextSizeFromCtrlFontSetting(ctrl, "TEST")(1)
                    rowPixelHeight = UserFormSizeToPixel(rowHeight)
                    ttkFontSetting = GenerateTtkFontSetting(ctrl)
                    codeList.Add indent & GenerateTtkStyleDefinitionCode(controlVarName, styleName, useCls) & vbLf
                    codeList.Add indent & prefix & "style.configure(" & ttkStyleRef & " + "".Heading""" & ", " & ttkFontSetting & ")" & vbLf
                    codeList.Add indent & prefix & "style.configure(" & ttkStyleRef & ", " & ttkFontSetting & ")" & vbLf
                    codeList.Add indent & prefix & "style.configure(" & ttkStyleRef & ", rowheight=" & rowPixelHeight & ")" & vbLf
                    codeList.Add indent & prefix & "style.configure(" & ttkStyleRef & ", foreground=" & q & FormColorToHex(ctrl.ForeColor) & q & ")" & vbLf
                    codeList.Add indent & prefix & "style.configure(" & ttkStyleRef & ", background=" & q & FormColorToHex(ctrl.BackColor) & q & ")" & vbLf
                    codeList.Add indent & prefix & "style.configure(" & ttkStyleRef & ", fieldbackground=" & q & FormColorToHex(ctrl.BackColor) & q & ")" & vbLf
                    If ctrl.MultiSelect Then
                        codeList.Add indent & controlVarName & ".configure(selectmode=" & q & "extended" & q & ")" & vbLf
                    Else
                        codeList.Add indent & controlVarName & ".configure(selectmode=" & q & "browse" & q & ")" & vbLf
                    End If
                    codeList.Add indent & listviewHeaderNames & " = " & GenerateListViewHeaders(ctrl) & vbLf
                    codeList.Add indent & controlVarName & ".configure(columns=" & listviewHeaderNames & ",show=" & q & "headings" & q & ")" & vbLf
                    codeList.Add DefineListViewColumns(ctrl, indent, prefix, useCls) & vbLf
                    codeList.Add indent & itemsListName & " = " & GetListViewItems(ctrl, indent) & vbLf
                    codeList.Add indent & "for item in " & itemsListName & ":" & vbLf
                    codeList.Add indent & "    " & controlVarName & ".insert("""", tk.END, values=item)" & vbLf
                    codeList.Add SetHScrollBarToWidget(ctrl, indent, prefix, useCls, False)
                    codeList.Add SetVScrollBarToWidget(ctrl, indent, prefix, useCls, False)
                End If
                
                If IsTreeView(ctrl) Then
                    styleName = tkStyleBaseName & ".Treeview"
                    rowHeight = GetTextSizeFromCtrlFontSetting(ctrl, "TEST")(1)
                    rowPixelHeight = UserFormSizeToPixel(rowHeight)
                    nodeDictName = controlVarName & "_node_dict"
                    ttkFontSetting = GenerateTtkFontSetting(ctrl)
                    codeList.Add indent & GenerateTtkStyleDefinitionCode(controlVarName, styleName, useCls) & vbLf
                    codeList.Add indent & prefix & "style.configure(" & ttkStyleRef & ", " & ttkFontSetting & ")" & vbLf
                    codeList.Add indent & prefix & "style.configure(" & ttkStyleRef & ", rowheight=" & rowPixelHeight & ")" & vbLf
                    codeList.Add indent & controlVarName & ".configure(selectmode=" & q & "browse" & q & ")" & vbLf
                    Set treeviewNodes = GetAllTreeViewNodesBfs(ctrl)
                    codeList.Add indent & controlVarName & ".configure(show=" & q & "tree" & q & ")" & vbLf
                    codeList.Add indent & nodeDictName & " = {}" & vbLf
                    For Each item In treeviewNodes
                        Set node = item(0)
                        nodeVarName = nodeDictName & "[" & q & Convert2PythonFormatText(node.Key) & q & "]"
                        If node.Parent Is Nothing Then
                            nodeParentVarName = q & q
                        Else
                            nodeParentVarName = nodeDictName & "[" & q & Convert2PythonFormatText(node.Parent.Key) & q & "]"
                        End If
                        codeList.Add indent & nodeVarName & " = " & controlVarName & ".insert(" & nodeParentVarName & ", tk.END, text=" & q & Convert2PythonFormatText(node.text) & q & ", open=" & node.Expanded & ")" & vbLf
                    Next item
                    If HasScrollProperty(ctrl) Then
                        enableScrollBar = ctrl.Scroll
                    Else
                        enableScrollBar = True
                    End If
                    If enableScrollBar Then
                        codeList.Add SetHScrollBarToWidget(ctrl, indent, prefix, useCls, False)
                        codeList.Add SetVScrollBarToWidget(ctrl, indent, prefix, useCls, False)
                    End If
                End If
                
                ' Set each Caption and font in MultiPage, font size is rounded
                If TypeName(ctrl) = "MultiPage" Then
                    styleName = tkStyleBaseName & ".TNotebook"
                    codeList.Add indent & GenerateTtkStyleDefinitionCode(controlVarName, styleName, useCls) & vbLf
                    
                    ttkFontSetting = GenerateTtkFontSetting(ctrl)
                    
                    codeList.Add indent & prefix & "style.configure(" & ttkStyleRef & ", background=" & q & FormColorToHex(ctrl.BackColor) & q & ")" & vbLf
                    
                    Select Case ctrl.Style
                        Case fmTabStyleTabs
                             codeList.Add indent & prefix & "style.configure(" & ttkStyleRef & ", borderwidth=2, relief=" & q & "raised" & q & ")" & vbLf
                        Case fmTabStyleButtons
                             codeList.Add indent & prefix & "style.configure(" & ttkStyleRef & ", borderwidth=0, relief=" & q & "flat" & q & ")" & vbLf
                        Case fmTabStyleNone
                             codeList.Add indent & prefix & "style.configure(" & ttkStyleRef & ", borderwidth=0, relief=" & q & "flat" & q & ")" & vbLf
                             codeList.Add indent & prefix & "style.layout(" & ttkStyleRef & " + "".Tab"", [])" & vbLf
                    End Select
                   
                    Select Case ctrl.TabOrientation
                        Case fmTabOrientationTop
                            tabPosition = "nw"
                        Case fmTabOrientationBottom
                            tabPosition = "sw"
                        Case fmTabOrientationLeft
                            tabPosition = "wn"
                        Case fmTabOrientationRight
                            tabPosition = "en"
                        Case Else
                            tabPosition = "n"
                    End Select
                    
                    codeList.Add indent & prefix & "style.configure(" & ttkStyleRef & ", tabposition=" & q & tabPosition & q & ")" & vbLf
                    
                    codeList.Add indent & prefix & "style.configure(" & ttkStyleRef & " + "".Tab""" & ", " & ttkFontSetting & ")" & vbLf
                    codeList.Add indent & prefix & "style.configure(" & ttkStyleRef & " + "".Tab""" & ", foreground=" & q & FormColorToHex(ctrl.ForeColor) & q & ")" & vbLf
                    
                    
                    For Each item In ctrl.Pages
                        childVarName = GenerateCtrlVarName(item, prefix, useCls)
                        caption = item.caption
                        caption = Convert2PythonFormatText(caption)
                        codeList.Add indent & childVarName & " = tk.Frame(" & controlVarName & ", bg=" & q & FormColorToHex(&H8000000F) & q & ")" & vbLf
                        codeList.Add indent & controlVarName & ".add(" & childVarName & ", text=" & q & caption & q & ")" & vbLf
                    Next
                    
                    
                End If
                
                
                ' Font size is rounded because Tkinter does not support floats in font settings
                If GetTkWidgetName(ctrl) = "tk.LabelFrame" Or ContainsValue(Array("Label", "CommandButton", "TextBox", "ListBox", "CheckBox", "ToggleButton", "OptionButton", "ComboBox"), TypeName(ctrl)) Then
                    fontStyle = ""
                    fontOpts = ""
                    
                    If ctrl.Font.Bold Then fontStyle = fontStyle & ", weight=" & q & "bold" & q
                    If ctrl.Font.Italic Then fontStyle = fontStyle & ", slant=" & q & "italic" & q
                    If ctrl.Font.Underline Then fontOpts = fontOpts & ", underline=1"
                    If ctrl.Font.Strikethrough Then fontOpts = fontOpts & ", overstrike=1"
                    
                    codeList.Add indent & controlVarName & ".configure(font=font.Font(family=" & q & ctrl.Font.Name & q & ", size=" & Round(ctrl.Font.Size) & fontStyle & fontOpts & "))" & vbLf
                End If
                
                
                If ContainsValue(Array("Frame", "TextBox", "Label", "ListBox", "Image"), TypeName(ctrl)) Then
                    ' Tkinter's Combobox does not support customizing border colors or relief
                    codeList.Add indent & controlVarName & GetBorderSetting(ctrl) & vbLf
                End If
                
                If GetTkWidgetName(ctrl) <> "tk.Text" And ContainsValue(Array("Label", "TextBox", "ComboBox", "CheckBox", "ToggleButton", "OptionButton"), TypeName(ctrl)) Then
                    codeList.Add GetTextAlignSetting(ctrl, indent, prefix, useCls) & vbLf
                End If
                
                ' Set mouse cursor
                If TypeName(ctrl) <> "MultiPage" Then
                    cursorType = GetControlCursorType(ctrl)
                    If cursorType <> "" Then
                        codeList.Add indent & controlVarName & ".configure(cursor=" & q & cursorType & q & ")" & vbLf
                    End If
                End If
                
                
                
                If TypeName(ctrl) = "Image" Then
                    
                    Select Case ctrl.PictureAlignment
                        Case fmPictureAlignmentTopLeft
                            canvasCoordX = "0"
                            canvasCoordY = "0"
                            canvasAnchor = "nw"
                        Case fmPictureAlignmentTopRight
                            canvasCoordX = "int(" & controlVarName & ".place_info()[""width""])"
                            canvasCoordY = "0"
                            canvasAnchor = "ne"
                        Case fmPictureAlignmentCenter
                            canvasCoordX = "int(" & controlVarName & ".place_info()[""width""])//2"
                            canvasCoordY = "int(" & controlVarName & ".place_info()[""height""])//2"
                            canvasAnchor = "center"
                        Case fmPictureAlignmentBottomLeft
                            canvasCoordX = "0"
                            canvasCoordY = "int(" & controlVarName & ".place_info()[""height""])"
                            canvasAnchor = "sw"
                        Case fmPictureAlignmentBottomRight
                            canvasCoordX = "int(" & controlVarName & ".place_info()[""width""])"
                            canvasCoordY = "int(" & controlVarName & ".place_info()[""height""])"
                            canvasAnchor = "se"
                        Case Else
                            canvasCoordX = "0"
                            canvasCoordY = "0"
                            canvasAnchor = "se"
                    End Select
                    
                    Select Case ctrl.PictureSizeMode
                        Case fmPictureSizeModeClip
                            pictureWidth = -1
                            pictureHeight = -1
                            resizePicture = False
                            preservePictureAspectRatio = False
                        Case fmPictureSizeModeStretch
                            pictureWidth = UserFormSizeToPixel(ctrl.width)
                            pictureHeight = UserFormSizeToPixel(ctrl.height)
                            resizePicture = True
                            preservePictureAspectRatio = False
                        Case fmPictureSizeModeZoom
                            pictureWidth = UserFormSizeToPixel(ctrl.width)
                            pictureHeight = UserFormSizeToPixel(ctrl.height)
                            resizePicture = True
                            preservePictureAspectRatio = True
                        Case Else
                            pictureWidth = -1
                            pictureHeight = -1
                            resizePicture = False
                            preservePictureAspectRatio = False
                    End Select
                    
                    If ctrl.Picture Is Nothing Or imageMode = "disabled" Then
                        hasPicture = False
                    Else
                        hasPicture = True
                    End If
                    
                    If resizePicture Then
                        ' Set the minimum image size to 2 to prevent conversion errors.
                        If 2 > pictureWidth And pictureWidth <> -1 Then pictureWidth = 2
                        If 2 > pictureHeight And pictureHeight <> -1 Then pictureHeight = 2
                    End If
                    
                    If hasPicture Then
                        tempPath = fso.BuildPath(fso.GetSpecialFolder(2).Path, fso.GetTempName())
                        
                        If ContainsValue(Array("base64", "base64-dict", "base64-json"), imageMode) Then
                            picturePath = tempPath & "tmp2.png"
                        Else
                            picturePath = saveDir & "\" & pictureName & ".png"
                        End If
                        
                        tempPath = tempPath & ".bmp"
                        
                        If ContainsValue(Array("file", "base64", "base64-dict", "base64-json"), imageMode) Then
                            Call SavePicture(ctrl.Picture, tempPath)
                            If resizePicture Then
                                ' When image resizing is needed, we temporarily export the resized image as a BMP without an alpha channel before converting it to PNG.
                                ' Since images in VBA UserForms are natively Bitmaps with no transparency support, dropping the alpha channel causes no side effects.
                                ' This step is necessary because running both resizing and PNG conversion simultaneously in ConvertImageFormat can introduce unintended transparent pixels.
                                ' (e.g., Occurred when upscaling a 1x1 solid-color image to a 100x100 PNG.)
                                Call ConvertImageFormat(tempPath, tempPath, resize:=resizePicture, dstWidth:=pictureWidth, dstHeight:=pictureHeight, preserveAspectRatio:=preservePictureAspectRatio)
                            End If
                            Call ConvertImageFormat(tempPath, picturePath, resize:=False)
                            If fso.FileExists(tempPath) Then Call fso.DeleteFile(tempPath)
                        End If
                        
                        
                        If ContainsValue(Array("base64", "base64-dict", "base64-json"), imageMode) Then
                            base64Str = FileToBase64(picturePath)
                            base64PictureDictStrs.Add "    " & q & pictureName & q & ": " & q & base64Str & q
                            If fso.FileExists(picturePath) Then Call fso.DeleteFile(picturePath)
                        End If
                        
                        If ContainsValue(Array("base64-dict", "base64-json", "base64-json-reference"), imageMode) Then
                            codeList.Add indent & controlVarName & "_photo = tk.PhotoImage(data=image_base64_dict[" & q & pictureName & q & "])" & vbLf
                        ElseIf imageMode = "base64" Then
                            codeList.Add indent & controlVarName & "_photo = tk.PhotoImage(data=" & q & base64Str & q & ")" & vbLf
                        ElseIf ContainsValue(Array("file", "reference-only"), imageMode) Then
                            codeList.Add indent & controlVarName & "_photo = tk.PhotoImage(file=BASE_DIR / r" & q & fso.GetFileName(picturePath) & q & ")" & vbLf
                        End If
                        
                        codeList.Add indent & controlVarName & ".create_image(" & canvasCoordX & ", " & canvasCoordY & ", image=" & controlVarName & "_photo" & ", anchor=" & q & canvasAnchor & q & ")" & vbLf
                    Else
                        codeList.Add indent & "#" & controlVarName & "_photo = tk.PhotoImage(file=r" & q & "" & q & ")" & vbLf
                        codeList.Add indent & "#" & controlVarName & ".create_image(" & canvasCoordX & ", " & canvasCoordY & ", image=" & controlVarName & "_photo" & ", anchor=" & q & canvasAnchor & q & ")" & vbLf
                    End If
                    

                End If
                
                
                codeList.Add vbLf
                
            Else
                MsgBox GenerateUnsupportedControlMessage(ctrl)
                result = ""
                GenerateTkinterCode = result
                Exit Function
            End If
        Next ctrl
        codeList.Add SetTkRadiobuttonValues(ctrls, indent, prefix, useCls) & vbLf
        codeList.Add SetTkCheckbuttonValues(ctrls, indent, prefix, useCls) & vbLf
        If Not useCls And Not noMainLoop Then
            codeList.Add formName & ".mainloop()" & vbLf
        End If
    Next root
    
    If useCls And Not noMainLoop Then
        clsNumber = 0
        For Each root In frms
            clsNumber = clsNumber + 1
            instanceName = "obj_" & root.Name
            If clsNumber <= 1 Then
                codeList.Add instanceName & " = " & root.Name & "()" & vbLf
                toplevelInstanceName = instanceName
            Else
                codeList.Add instanceName & " = " & root.Name & "(" & toplevelInstanceName & "." & FORM_WINDOW_NAME & ")" & vbLf
            End If
        Next
        codeList.Add toplevelInstanceName & "." & FORM_WINDOW_NAME & ".mainloop()"
    End If
    
    result = JoinCollection(codeList, "")
    
    If ContainsValue(Array("base64-dict", "base64-json", "base64-json-reference"), imageMode) Then
        base64PictureDictStr = "{" & vbLf & JoinCollection(base64PictureDictStrs, "," & vbLf) & vbLf & "}" & vbLf
        If imageMode = "base64-dict" Then
            result = VBA.Replace$(result, uniqueStringToReplace, "image_base64_dict = " & base64PictureDictStr)
        Else
            result = VBA.Replace$(result, uniqueStringToReplace, "")
        End If
        
        
        If imageMode = "base64-json" Then
            Call SaveUtf8TextNoBom(saveDir & "\" & "image_base64.json", base64PictureDictStr)
        End If
    
    Else
        result = VBA.Replace$(result, uniqueStringToReplace, "")
    End If
    
    
    GenerateTkinterCode = result
End Function

Private Function GetSaveDirPath() As String
    Dim result As String
    Dim wsh As Object
    result = ""
    On Error Resume Next
    result = CallByName(Application, "ThisWorkbook", VbGet).Path ' Excel VBA
    result = CallByName(Application, "MacroContainer", VbGet).Path ' Word VBA
    On Error GoTo 0
    
    If result = "" Then ' Other Office / Unsaved Workbook
        On Error Resume Next
        Set wsh = CreateObject("WScript.Shell")
        result = wsh.SpecialFolders("MyDocuments")
        On Error GoTo 0
    End If
    
    If result = "" Then ' If fail to retrieve the path to the Documents folder
        result = "C:"
    End If
    
    result = result & "\" & OUTPUT_FOLDER_NAME
    GetSaveDirPath = result
End Function

Private Sub CreateFolderIfDoesNotExist(ByVal folderPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Not fso.FolderExists(folderPath) Then
        Call fso.CreateFolder(folderPath)
    End If
    
End Sub

Private Function GetUserFormObjectFromCtrl(ByVal ctrl As Object) As Object
    ' Get the ancestor (UserForm) of the control.
    
    Dim root As Object
    
    If ctrl Is Nothing Then
        Err.Raise 13
    End If
    
    Set root = ctrl
    ' Loop to get root(UserForm) object
    On Error GoTo Finally:
    Do While True
        Set root = root.Parent
    Loop
    On Error GoTo 0

Finally:
    Set GetUserFormObjectFromCtrl = root
End Function

Private Function GenerateCtrlVarName(ByVal ctrl As Object, ByVal prefix As String, ByVal useCls As Boolean) As String
    ' Generates a valid, unique identifier for a control in the target language.
    Dim controlVarName As String
    If IsRootForm(ctrl) And useCls Then
        controlVarName = prefix & FORM_WINDOW_NAME
    Else
        If TypeName(ctrl) = "Page" Then
        ' VBA allows duplicate names for Page objects if they belong to different MultiPage controls.
        ' To ensure unique variable names in the target language (which typically uses a flat
        ' namespace), namespace the Page by prepending its parent MultiPage's name.
        ' Example: "Page1" inside "MultiPage1" becomes "MultiPage1_Page1"
            controlVarName = prefix & ctrl.Parent.Name & "_" & ctrl.Name
        Else
            controlVarName = prefix & ctrl.Name
        End If
    End If
    GenerateCtrlVarName = controlVarName
End Function

Private Function IsRootForm(ByVal ctrl As Object) As Boolean
    ' Determines whether the specified control is the root UserForm.
    '
    ' This function returns True only when:
    '   - The control is of type MSForms.UserForm, and
    '   - The control exists at the top level (i.e., its hierarchy depth is 0).
    '
    ' Note:
    '   Even if the control is of type MSForms.UserForm, this function will return False
    '   if the control is not the root window (for example, if it is nested or owned
    '   within another container or context).
    Dim result As Boolean
    If GetFormControlDepth(ctrl) = 0 And TypeOf ctrl Is MSForms.UserForm Then
        result = True
    Else
        result = False
    End If
    IsRootForm = result
End Function

Private Function GetBorderSetting(ByVal ctrl As Object) As String
    Dim result As String
    Const q As String = """"
    Dim borderWidth As Long
    Dim relief As String
    Dim highlightBorderWidth As Long
    Dim hexBorderColor As String
    hexBorderColor = FormColorToHex(ctrl.BorderColor)
    relief = "tk.FLAT"
    highlightBorderWidth = 0

    borderWidth = 2
    Select Case ctrl.BorderStyle
        Case fmBorderStyleSingle
            ' SpecialEffect is fmSpecialEffectFlat if BorderStyle is fmBorderStyleSingle
            borderWidth = 0
            highlightBorderWidth = 1
            relief = "tk.FLAT"
        Case fmBorderStyleNone
            Select Case ctrl.SpecialEffect
                Case fmSpecialEffectFlat
                    borderWidth = 0
                    relief = "tk.FLAT"
                Case fmSpecialEffectRaised
                    relief = "tk.RAISED"
                Case fmSpecialEffectSunken
                    relief = "tk.SUNKEN"
                Case fmSpecialEffectEtched
                    relief = "tk.GROOVE"
                Case fmSpecialEffectBump
                    relief = "tk.RIDGE"
            End Select
    End Select

    result = ".configure(relief=" & relief & ", bd=" & borderWidth & ", highlightthickness=" & highlightBorderWidth & ", highlightbackground=" & q & hexBorderColor & q & ", highlightcolor=" & q & hexBorderColor & q & ")"
    GetBorderSetting = result
End Function


Private Function GetTextAlignSetting(ByVal ctrl As Object, ByVal indent As String, ByVal prefix As String, ByVal useCls As Boolean) As String
    Dim result As String
    Const q As String = """"
    Dim anchor As String
    Dim justify As String
    Dim controlVarName As String
    controlVarName = GenerateCtrlVarName(ctrl, prefix, useCls)
    result = ""
    If ContainsValue(Array("CheckBox", "OptionButton", "ToggleButton"), TypeName(ctrl)) Then
        Select Case ctrl.TextAlign
            Case fmTextAlignLeft
                anchor = "w"
                justify = "left"
            Case fmTextAlignCenter
                anchor = "center"
                justify = "center"
            Case fmTextAlignRight
                anchor = "e"
                justify = "right"
            Case Else
                anchor = "center"
                justify = "center"
        End Select
    Else
        Select Case ctrl.TextAlign
            Case fmTextAlignLeft
                anchor = "nw"
                justify = "left"
            Case fmTextAlignCenter
                anchor = "n"
                justify = "center"
            Case fmTextAlignRight
                anchor = "ne"
                justify = "right"
            Case Else
                anchor = "n"
                justify = "center"
        End Select
    End If
    If Not ContainsValue(Array("TextBox", "ComboBox"), TypeName(ctrl)) Then
        result = indent & controlVarName & ".configure(anchor=" & q & anchor & q & ")" & vbLf
    End If
    result = result & indent & controlVarName & ".configure(justify=" & q & justify & q & ")"
    GetTextAlignSetting = result
End Function

Private Function GenerateTtkStyleRef(ByVal tkVarName As String) As String
    Dim result As String
    result = tkVarName & ".cget(""style"")"
    GenerateTtkStyleRef = result
End Function

Private Function GenerateTtkStyleDefinitionCode(ByVal tkVarName As String, ByVal styleName As String, ByVal useCls As Boolean) As String
    ' If useCls is True, "self.__class__.__name__" is prepended to the style name to prevent naming conflicts.
    ' Example:
    ' GenerateTtkStyleDefinitionCode("self.ComboBox1", "ComboBox1.style.TComboBox", True)
    ' -> self.ComboBox1.configure(style=self.__class__.__name__ + "." + "ComboBox1.style.TComboBox")
    ' GenerateTtkStyleDefinitionCode("ComboBox1", "ComboBox1.style.TComboBox", False)
    ' -> ComboBox1.configure(style="ComboBox1.style.TComboBox")
    Dim codeList As New Collection
    Dim code As String
    Const q As String = """"
    codeList.Add tkVarName & ".configure(style="
    
    If useCls Then
        codeList.Add "self.__class__.__name__ + "
        codeList.Add q & "." & styleName & q & ")"
    Else
        codeList.Add q & styleName & q & ")"
    End If
    code = JoinCollection(codeList, "")
    GenerateTtkStyleDefinitionCode = code
End Function

Private Function GenerateTtkFontSetting(ByVal ctrl As Object) As String
    ' Returns the font setting for a ttk widget, such as:
    ' font=("Arial Narrow", 10, "")
    Dim ttkFontSetting As String
    Dim ttkFontStyleSetting As String
    Dim ttkFontStyles As Collection
    Const q As String = """"
    Set ttkFontStyles = New Collection
    ttkFontStyleSetting = ""
    
    If ctrl.Font.Bold Then ttkFontStyles.Add "bold"
    If ctrl.Font.Italic Then ttkFontStyles.Add "italic"
    If ctrl.Font.Underline Then ttkFontStyles.Add "underline"
    If ctrl.Font.Strikethrough Then ttkFontStyles.Add "overstrike"
    
    ttkFontStyleSetting = q & JoinCollection(ttkFontStyles, " ") & q
    ttkFontSetting = "font=(" & Join(Array(q & ctrl.Font.Name & q, Round(ctrl.Font.Size), ttkFontStyleSetting), ", ") & ")"
    GenerateTtkFontSetting = ttkFontSetting
End Function

Private Function GetTkWidgetName(ByVal ctrl As Object) As String
    Dim result As String
    Select Case TypeName(ctrl)
        Case "Label"
            result = "tk.Label"
        Case "CommandButton"
            result = "tk.Button"
        Case "Frame"
            If ctrl.caption = "" Then
                result = "tk.Frame"
            Else
                result = "tk.LabelFrame"
            End If
        Case "TextBox"
            If ctrl.MultiLine Then
                result = "tk.Text"
            Else
                result = "tk.Entry"
            End If
        Case "SpinButton"
            result = "tk.Spinbox"
        Case "ListBox"
            result = "tk.Listbox"
        Case "CheckBox"
            result = "tk.Checkbutton"
        Case "ToggleButton"
            result = "tk.Checkbutton"
        Case "OptionButton"
            result = "tk.Radiobutton"
        Case "Image"
            result = "tk.Canvas"
        Case "ScrollBar"
            result = "ttk.Scale"
        Case "ComboBox"
            result = "ttk.Combobox"
        Case "MultiPage"
            result = "ttk.Notebook"
        Case Else
            result = ""
            
            If IsListView(ctrl) Then
                result = "ttk.Treeview"
            End If
            
            If IsTreeView(ctrl) Then
                result = "ttk.Treeview"
            End If
            
    End Select
    GetTkWidgetName = result
End Function


Private Function IsListView(ByVal ctrl As Object) As Boolean
    ' Since the class name of the ListView may vary depending on the version, so use InStr to check it.
    ' e.g ListView/ListView2/ListView3/ListView4
    If InStr(TypeName(ctrl), "ListView") = 1 Then
        IsListView = True
    Else
        IsListView = False
    End If
End Function

Private Function IsTreeView(ByVal ctrl As Object) As Boolean
    ' Since the class name of the TreeView may vary depending on the version, so use InStr to check it.
    ' e.g TreeView/TreeView2/TreeView3/TreeView4
    If InStr(TypeName(ctrl), "TreeView") = 1 Then
        IsTreeView = True
    Else
        IsTreeView = False
    End If
End Function

Private Function GetControlCursorType(ByVal ctrl As Object) As String
    Dim cursorType As String
    Select Case ctrl.MousePointer
        Case fmMousePointerDefault
            cursorType = ""               ' Default cursor
        Case fmMousePointerArrow
            cursorType = "arrow"          ' Arrow(normal)
        Case fmMousePointerCross
            cursorType = "cross"          ' Cross
        Case fmMousePointerIBeam
            cursorType = "xterm"          ' For inputting text
        Case fmMousePointerSizeNESW
            cursorType = "size_ne_sw"     ' Arrow(NESW)
        Case fmMousePointerSizeNS
            cursorType = "size_ns"        ' Arrow(NS)
        Case fmMousePointerSizeNWSE
            cursorType = "size_nw_se"     ' Arrow(NWSE)
        Case fmMousePointerSizeWE
            cursorType = "size_we"        ' Arrow(WE)
        Case fmMousePointerUpArrow
            cursorType = "center_ptr"     ' Arrow(up)
        Case fmMousePointerHourGlass
            cursorType = "watch"          ' Busy(hourglass)
        Case fmMousePointerNoDrop
            cursorType = "no"             ' "Not allowed" synbol
        Case fmMousePointerAppStarting
            cursorType = "watch"          ' Busy(hourglass) (Subsutitute it because Tkinter does not support same cursor.)
        Case fmMousePointerHelp
            cursorType = "question_arrow" ' Question arrow
        Case fmMousePointerSizeAll
            cursorType = "fleur"          ' Four headed Arrow
        Case Else
            cursorType = ""               ' Others are default cursor.
    End Select
    GetControlCursorType = cursorType
End Function

Private Function SetTkRadiobuttonValues(ByVal ctrls As Variant, ByVal indent As String, ByVal prefix As String, ByVal useCls As Boolean) As String
    Dim codeList As New Collection
    Dim parentList As New Collection
    Const q As String = """"
    Dim varName As String
    Dim ctrl As Variant
    Dim result As String
    Dim parentVarName As String
    Dim controlVarName As String
    Dim radioButtonStrValue As String
    For Each ctrl In ctrls
        controlVarName = GenerateCtrlVarName(ctrl, prefix, useCls)
        parentVarName = GenerateCtrlVarName(ctrl.Parent, prefix, useCls)
        radioButtonStrValue = "StrVar_" & GenerateCtrlVarName(ctrl, "", False)
        If TypeName(ctrl) = "OptionButton" Then
            varName = parentVarName & "_radiobutton_value"
            If Not CollContainsKey(parentList, parentVarName) Then
                ' Use the Collection key to check and avoid redeclaring a variable that has already been declared
                parentList.Add "", parentVarName
                codeList.Add indent & varName & " = tk.StringVar()" & vbLf
                codeList.Add indent & varName & ".set(None)" & vbLf ' Deselect the radio button
            End If
            codeList.Add indent & controlVarName & ".configure(variable=" & varName & ", value=" & q & radioButtonStrValue & q & ")" & vbLf
            If ctrl.value = True Then
                codeList.Add indent & varName & ".set(" & q & radioButtonStrValue & q & ")" & vbLf
            End If
            
        End If
    Next
    result = JoinCollection(codeList, "")
    SetTkRadiobuttonValues = result
End Function

Private Function SetTkCheckbuttonValues(ByVal ctrls As Variant, ByVal indent As String, ByVal prefix As String, ByVal useCls As Boolean) As String
    Dim codeList As New Collection
    Dim varName As String
    Dim ctrl As Variant
    Dim value As Boolean
    Dim result As String
    Dim controlVarName As String
    For Each ctrl In ctrls
        controlVarName = GenerateCtrlVarName(ctrl, prefix, useCls)
        If TypeName(ctrl) = "CheckBox" Or TypeName(ctrl) = "ToggleButton" Then
            varName = controlVarName & "_checkbutton_value"
            codeList.Add indent & varName & " = tk.BooleanVar()" & vbLf
            codeList.Add indent & controlVarName & ".configure(variable=" & varName & ")" & vbLf
            If ctrl.value = True Then
                value = True
            Else
                value = False
            End If
            codeList.Add indent & varName & ".set(" & value & ")" & vbLf
            
        End If
    Next
    result = JoinCollection(codeList, "")
    SetTkCheckbuttonValues = result
End Function

Private Function SetVScrollBarToWidget(ByVal ctrl As Object, ByVal indent As String, ByVal prefix As String, ByVal useCls As Boolean, ByVal setToInside As Boolean) As String
    ' This function generates and returns the Tkinter Python code required to attach a vertical scrollbar to a specific widget.
    ' setToInside:
    '   A boolean flag that determines the scrollbar's placement. If set to True, the scrollbar is placed within the widget's boundaries (overlay mode). If False, it is placed outside the widget's edge.
    Dim codeList As New Collection
    Dim result As String
    Dim scrollBarVarName As String
    Dim scrollBarCoordX As String
    Dim scrollBarCoordY As String
    Dim scrollBarCoordWidth As String
    Dim scrollBarCoordHeight As String
    Dim controlVarName As String
    Dim parentVarName As String
    controlVarName = GenerateCtrlVarName(ctrl, prefix, useCls)
    parentVarName = GenerateCtrlVarName(ctrl.Parent, prefix, useCls)
    scrollBarVarName = controlVarName & "_" & "vscrollbar"
    scrollBarCoordX = "int(" & controlVarName & ".place_info()[""x""])" & "+" & "int(" & controlVarName & ".place_info()[""width""])"
    scrollBarCoordY = "int(" & controlVarName & ".place_info()[""y""])"
    scrollBarCoordWidth = "20"
    scrollBarCoordHeight = "int(" & controlVarName & ".place_info()[""height""])"
    
    If setToInside Then
        scrollBarCoordX = scrollBarCoordX & "-" & scrollBarCoordWidth
        scrollBarCoordY = scrollBarCoordY & "+1"
        scrollBarCoordWidth = scrollBarCoordWidth & "-1"
        scrollBarCoordHeight = scrollBarCoordHeight & "-2"
    End If
    
    codeList.Add indent & scrollBarVarName & " = tk.Scrollbar(" & parentVarName & ", orient=""vertical"")" & vbLf
    codeList.Add indent & scrollBarVarName & ".place(x=" & scrollBarCoordX & ",y=" & scrollBarCoordY & ", width=" & scrollBarCoordWidth & ", height=" & scrollBarCoordHeight & ")" & vbLf
    codeList.Add indent & controlVarName & ".configure(yscrollcommand=" & scrollBarVarName & ".set)" & vbLf
    codeList.Add indent & scrollBarVarName & ".configure(command=" & controlVarName & ".yview)" & vbLf
    result = JoinCollection(codeList, "")
    SetVScrollBarToWidget = result
End Function

Private Function SetHScrollBarToWidget(ByVal ctrl As Object, ByVal indent As String, ByVal prefix As String, ByVal useCls As Boolean, ByVal setToInside As Boolean) As String
    ' This function generates and returns the Tkinter Python code required to attach a horizontal scrollbar to a specific widget
    ' setToInside:
    '   A boolean flag that determines the scrollbar's placement. If set to True, the scrollbar is placed within the widget's boundaries (overlay mode). If False, it is placed outside the widget's edge.
    Dim codeList As New Collection
    Dim result As String
    Dim scrollBarVarName As String
    Dim scrollBarCoordX As String
    Dim scrollBarCoordY As String
    Dim scrollBarCoordWidth As String
    Dim scrollBarCoordHeight As String
    Dim controlVarName As String
    Dim parentVarName As String
    controlVarName = GenerateCtrlVarName(ctrl, prefix, useCls)
    parentVarName = GenerateCtrlVarName(ctrl.Parent, prefix, useCls)
    scrollBarVarName = controlVarName & "_" & "hscrollbar"
    scrollBarCoordX = "int(" & controlVarName & ".place_info()[""x""])"
    scrollBarCoordY = "int(" & controlVarName & ".place_info()[""y""])" & "+" & "int(" & controlVarName & ".place_info()[""height""])"
    scrollBarCoordWidth = "int(" & controlVarName & ".place_info()[""width""])"
    scrollBarCoordHeight = "20"
    
    If setToInside Then
        scrollBarCoordX = scrollBarCoordX & "+1"
        scrollBarCoordY = scrollBarCoordY & "-" & scrollBarCoordHeight
        scrollBarCoordWidth = scrollBarCoordWidth & "-2"
        scrollBarCoordHeight = scrollBarCoordHeight & "-1"
    End If
    
    codeList.Add indent & scrollBarVarName & " = tk.Scrollbar(" & parentVarName & ", orient=""horizontal"")" & vbLf
    codeList.Add indent & scrollBarVarName & ".place(x=" & scrollBarCoordX & ",y=" & scrollBarCoordY & ", width=" & scrollBarCoordWidth & ", height=" & scrollBarCoordHeight & ")" & vbLf
    codeList.Add indent & controlVarName & ".configure(xscrollcommand=" & scrollBarVarName & ".set)" & vbLf
    codeList.Add indent & scrollBarVarName & ".configure(command=" & controlVarName & ".xview)" & vbLf
    result = JoinCollection(codeList, "")
    SetHScrollBarToWidget = result
End Function

Private Function GetListBoxValue(ByVal ctrl As Object, ByVal indent As String) As String
    ' Retrieve the items of a ListBox or ComboBox as a string in the format ["1", "2", "3"].
    Const q As String = """"
    Dim codeList As New Collection
    Dim item As Variant
    Dim i As Long: i = 0
    Dim result As String
    Dim listIndent As String: listIndent = "    " & indent
    Const maxItemsPerLine As Long = 3
    If ctrl.ListCount > 0 Then
        If ctrl.ListCount > maxItemsPerLine Then codeList.Add vbLf & listIndent
        For Each item In ctrl.List
            i = i + 1
            codeList.Add q & Convert2PythonFormatText(item) & q
            If Not i = ctrl.ListCount Then
                codeList.Add ", "
                If i Mod maxItemsPerLine = 0 And ctrl.ListCount > maxItemsPerLine Then codeList.Add vbLf & listIndent
            Else
                If ctrl.ListCount > maxItemsPerLine Then codeList.Add vbLf
                Exit For
            End If
        Next item
    End If
    
    result = JoinCollection(codeList, "")
    
    If ctrl.ListCount > maxItemsPerLine Then
        result = "[" & result & indent & "]"
    Else
        result = "[" & result & "]"
    End If
    
    GetListBoxValue = result
End Function

Private Function DefineListViewColumns(ByVal ctrl As Object, ByVal indent As String, ByVal prefix As String, ByVal useCls As Boolean) As String
    ' Generate code for the ListView headers.
    ' Example:
    ' ListView1.heading("col1", text="Header1", anchor="w")
    ' ListView1.heading("col2", text="Header2", anchor="w")
    ' ListView1.heading("col3", text="Header3", anchor="w")
    ' ListView1.column("col1", width=133, anchor="w")
    ' ListView1.column("col2", width=133, anchor="w")
    ' ListView1.column("col3", width=133, anchor="w")
    Dim codeList As New Collection
    Dim objHeaders As Object
    Set objHeaders = ctrl.ColumnHeaders
    Dim controlVarName As String
    Dim i As Long
    Dim item As Variant
    Dim result As String
    Dim colName As String
    Dim headerText As String
    Dim colWidth As Long
    Dim anchor As String
    Const q As String = """"
    controlVarName = GenerateCtrlVarName(ctrl, prefix, useCls)
    i = 0
    For Each item In objHeaders
        i = i + 1
        colName = "col" & i
        headerText = Convert2PythonFormatText(item.text)
        anchor = GetTtkTreeviewAnchor(item)
        codeList.Add indent & controlVarName & ".heading(" & q & colName & q & ", text=" & q & headerText & q & ", anchor=" & q & anchor & q & ")" & vbLf
    Next item
    
    i = 0
    For Each item In objHeaders
        i = i + 1
        colName = "col" & i
        anchor = GetTtkTreeviewAnchor(item)
        colWidth = UserFormSizeToPixel(item.width)
        codeList.Add indent & controlVarName & ".column(" & q & colName & q & ", width=" & colWidth & ", anchor=" & q & anchor & q & ", stretch=False)" & vbLf
    Next item
    
    result = JoinCollection(codeList, "")
    
    DefineListViewColumns = result
End Function

Private Function GetTtkTreeviewAnchor(ByVal objLvHeader As Object) As String
    ' Header Alignment(ListView) -> anchor(ttk.Treeview)
    Const cnsLvwColumnLeft As Long = 0
    Const cnsLvwColumnRight As Long = 1
    Const cnsLvwColumnCenter As Long = 2
    Dim result As String
    Select Case objLvHeader.Alignment
        Case cnsLvwColumnLeft
            result = "w"
        Case cnsLvwColumnRight
            result = "e"
        Case cnsLvwColumnCenter
            result = "center"
        Case Else
            result = "w"
    End Select
    GetTtkTreeviewAnchor = result
End Function

Private Function GenerateListViewHeaders(ByVal ctrl As Object) As String
    ' Generate the header names of a ListView as a string in the format ["col1", "col2", "col3"]
    Dim objHeaders As Object
    Set objHeaders = ctrl.ColumnHeaders
    Dim i As Long
    Dim item As Variant
    Dim colName As String
    Dim result As String
    Const q As String = """"
    Dim coll As New Collection
    For Each item In objHeaders
        i = i + 1
        colName = "col" & i
        coll.Add q & colName & q
    Next item
    result = "[" & JoinCollection(coll, ", ") & "]"
    GenerateListViewHeaders = result
End Function

Private Function GetListViewItems(ByVal ctrl As Object, ByVal indent As String) As String
    ' Retrieve the items of a ListView as a string in the format:
    ' [
    '     ["Item1-1", "Item1-2", "Item1-3"],
    '     ["Item2-1", "Item2-2", "Item2-3"],
    '     ["Item3-1", "Item3-2", "Item3-3"]
    ' ]
    Dim item As Object
    Dim i As Long
    Dim coll As Collection
    Dim resultColl As New Collection
    Dim result As String
    Const arrayLiteralStart As String = "["
    Const arrayLiteralEnd As String = "]"
    Const q As String = """"
    For Each item In ctrl.ListItems
        Set coll = New Collection
        coll.Add q & Convert2PythonFormatText(item.text) & q
        
        ' Because older versions of the ListView control do not support For Each for SubItems, use index-based access.
        For i = 1 To ctrl.ColumnHeaders.Count - 1
            coll.Add q & Convert2PythonFormatText(item.SubItems(i)) & q
        Next
        
        resultColl.Add arrayLiteralStart & JoinCollection(coll, ", ") & arrayLiteralEnd
        
    Next
    
    If resultColl.Count > 0 Then
        result = arrayLiteralStart & vbLf & indent & "    " & JoinCollection(resultColl, ", " & vbLf & indent & "    ") & vbLf & indent & arrayLiteralEnd
    Else
        result = arrayLiteralStart & arrayLiteralEnd
    End If
    GetListViewItems = result
End Function

Private Function GetTextSizeFromCtrlFontSetting(ByVal ctrl As Object, ByVal targetText As String) As Variant()
    '------------------------------------------------------------------------------
    ' Returns the rendered text size (Width, Height) for a given text string
    ' using the same font settings as the specified control.
    ' Size is measured in points, not pixels.
    '
    ' Parameters:
    '   ctrl        - The reference control whose font settings will be used.
    '   targetText  - The text to measure. If empty, "i" is used to ensure a measurable size is returned.
    '                 (The letter "i" is one of the ASCII characters with the narrowest rendering width.)
    '
    ' Returns:
    '   Variant() Array
    '       (0) = Text width
    '       (1) = Text height
    '
    ' Notes:
    '   - A temporary hidden Label control is dynamically created on the parent
    '     UserForm to calculate the actual rendered text dimensions.
    '   - AutoSize is enabled so the Label automatically resizes to fit the text.
    '   - The temporary control is removed immediately after measurement.
    '
    ' Compatibility Note:
    '   In Excel 2013 and earlier, it was confirmed that enabling .AutoSize does not
    '   correctly update the .Width and .Height properties, which remain 0.
    '   To avoid returning invalid measurements, this function falls back to
    '   an estimated size calculation based on the font size and text length.
    '   This fallback is less accurate than actual rendered text measurement.
    '
    '------------------------------------------------------------------------------
    Dim rootForm As Object
    Dim tempLabel As Object
    Dim tempName As String
    Dim textWidthSize As Double
    Dim textHeightSize As Double
    ' Prevent zero-size measurement for empty strings.
    If targetText = "" Then targetText = "i"
    ' Generate a unique temporary control name.
    tempName = "TempLabel_" & VBA.Replace$(GenerateUUIDv4(), "-", "_")
    ' Get the parent UserForm from the specified control.
    Set rootForm = GetUserFormObjectFromCtrl(ctrl)
    ' Create a temporary Label control for text measurement.
    Set tempLabel = rootForm.Controls.Add("Forms.Label.1", tempName, True)
    ' Initialize control properties.
    tempLabel.height = 0
    tempLabel.width = 0
    tempLabel.caption = ""
    tempLabel.AutoSize = True
    tempLabel.WordWrap = False
    ' Optional debug background color.
    tempLabel.BackColor = &H80C0FF
    ' Copy font settings from the source control.
    tempLabel.Font.Name = ctrl.Font.Name
    tempLabel.Font.Size = ctrl.Font.Size
    tempLabel.Font.Bold = ctrl.Font.Bold
    tempLabel.Font.Italic = ctrl.Font.Italic
    tempLabel.Font.Underline = ctrl.Font.Underline
    tempLabel.Font.Strikethrough = ctrl.Font.Strikethrough
    ' Apply target text so AutoSize calculates the rendered dimensions.
    tempLabel.caption = targetText
    ' Read calculated size.
    textWidthSize = tempLabel.width
    textHeightSize = tempLabel.height
    
    ' In Excel 2013 and earlier, it was confirmed that the result of .AutoSize
    ' is not reflected in .Width/.Height and remains 0.
    ' As a fallback, the font size is used instead,
    ' although the measurement accuracy is reduced.
    If textWidthSize = 0 Then textWidthSize = ctrl.Font.Size * Len(targetText)
    If textHeightSize = 0 Then textHeightSize = ctrl.Font.Size
    
    ' The .Controls.Remove method does not accept a String argument; the argument must be of type Variant (String).
    ' example: tempLabel.Name or CVar(tempName)
    Call rootForm.Controls.Remove(tempLabel.Name)
    ' Release object reference.
    Set tempLabel = Nothing
    ' Return width and height as an array.
    GetTextSizeFromCtrlFontSetting = VBA.Array(textWidthSize, textHeightSize)
End Function

Private Function GetAllTreeViewNodesBfs(ByVal treeviewCtrl As Object) As Collection
    ' This function performs a Breadth-First Search (BFS) on a TreeView control
    ' and returns a collection of nodes along with their hierarchy path.
    ' example:
    ' [[node, "1"], [node, "1-1"], [node, "1-2"], [node, "1-3"], [node, "1-4"], [node, "1-5"], [node, "1-6"], [node, "1-1-1"]]
    Dim queue As Collection
    Dim item As Variant
    Dim node As Object
    Dim child As Object
    Dim hierarchy As String
    Dim childIndex As Long
    Dim resultColl As Collection
    Set resultColl = New Collection
    Set queue = New Collection
    
    Dim nd As Object
    Dim rootIndex As Long
    rootIndex = 1
    
    ' Step 1: Add all root nodes (nodes without parents) to the queue
    ' Each root gets a hierarchy label like "1", "2", etc.
    For Each nd In treeviewCtrl.nodes
        If nd.Parent Is Nothing Then
            queue.Add VBA.Array(nd, CStr(rootIndex))
            rootIndex = rootIndex + 1
        End If
    Next nd
    
    ' Step 2: Perform BFS traversal
    Do While queue.Count > 0
        item = queue(1)
        queue.Remove 1
        
        Set node = item(0)
        hierarchy = item(1)
        ' Add current node and its hierarchy to the result collection
        resultColl.Add VBA.Array(node, hierarchy)
        ' Step 3: Enqueue all children of the current node
        Set child = node.child ' Get first child
        childIndex = 1
        
        Do While Not child Is Nothing
            ' Append child index to hierarchy (e.g., "1-2", "1-2-1")
            queue.Add VBA.Array(child, hierarchy & "-" & childIndex)
            childIndex = childIndex + 1
            Set child = child.Next
        Loop
    Loop
    ' Return the collection of (node, hierarchy) pairs
    Set GetAllTreeViewNodesBfs = resultColl
End Function

Private Function HasScrollProperty(ByVal ctrl As Object) As Boolean
    ' Since the Scroll property does not exist in older versions of TreeView, use this function to check for the property beforehand.
    Dim temp As Variant
    On Error GoTo Exception
    temp = VBA.Array(ctrl.Scroll)
    HasScrollProperty = True
    On Error GoTo 0
    Exit Function
Exception:
    HasScrollProperty = False
End Function

Private Function Convert2PythonFormatText(ByVal text As String) As String
    ' Escape special characters in the string
    text = VBA.Replace$(text, "\", "\\")
    text = VBA.Replace$(text, """", "\" & """")
    text = VBA.Replace$(text, "'", "\" & "'")
    ' Convert VBA line breaks to Python format
    ' vbCrLf should be replaced first
    text = VBA.Replace$(text, vbCrLf, vbLf)
    text = VBA.Replace$(text, vbCr, vbLf)
    text = VBA.Replace$(text, vbLf, "\n")
    Convert2PythonFormatText = text
End Function

Private Function FormColorToHex(ByVal clr As Long) As String
    ' Example:
    ' 16777215 -> "#FFFFFF"
    ' 0 -> "#000000"
    ' &H000000FF& (255) -> "#FF0000"
    ' &H00B4769E& (11826846) -> "#9E76B4"
    ' &H8000000F& (-2147483633) -> "#F0F0F0"(Windows XP[Luna Theme]/10/11), "#D4D0C8"(Windows 2000/XP[Classic Theme])
    Dim r As Long, g As Long, b As Long
    ' Convert a system color to its decimal color code when the parameter is a system color
    If 0 > clr Or clr >= 2147483648# Then
        clr = GetSysColor(clr And &HFF)
    End If
    ' Retrieve each component of the RGB color.
    r = clr And &HFF            ' Extract low-order 8 bits
    g = (clr \ &H100) And &HFF  ' Extract bits 8-15
    b = (clr \ &H10000) And &HFF ' Extract bits 16-23
    
    ' Convert the decimal RGB values to a #RRGGBB hex string and return it
    FormColorToHex = "#" & _
                     Right$("0" & Hex(r), 2) & _
                     Right$("0" & Hex(g), 2) & _
                     Right$("0" & Hex(b), 2)
End Function

Private Sub ConvertImageFormat(ByVal srcPath As String, ByVal dstPath As String, _
    Optional ByVal resize As Boolean = False, _
    Optional ByVal dstWidth As Long = -1, _
    Optional ByVal dstHeight As Long = -1, _
    Optional ByVal preserveAspectRatio As Boolean = False)
    
    '----------------------------------------------------------------------------------------------------
    ' Converts an image file to a specified output format and optionally resizes the image before saving.
    ' Uses GDI+ API without relying on WIA components (for compatibility with older versions of Windows).
    '
    ' Parameters:
    '   srcPath              - Full path of the source image file.
    '   dstPath              - Full path of the destination image file. Output format is
    '                          determined by the file extension (.png, .jpg, .jpeg, .bmp, .gif).
    '   resize               - If True, resizes the image before conversion.
    '                          Default: False.
    '   dstWidth             - Target width for resizing in pixels.
    '                          If -1, uses the original image width.
    '                          Default: -1.
    '   dstHeight            - Target height for resizing in pixels.
    '                          If -1, uses the original image height.
    '                          Default: -1.
    '   preserveAspectRatio  - If True, maintains the original aspect ratio during resizing.
    '                          Default: False.
    '
    '
    ' Supported Formats:
    '   PNG (.png)
    '   JPEG (.jpg, .jpeg)
    '   BMP (.bmp)
    '   GIF (.gif)
    '
    ' Errors:
    '   Raises Error if the destination file extension is not supported.
    '
    '----------------------------------------------------------------------------------------------------
    
    Const PixelFormat32bppARGB As Long = &H26200A
    Const InterpolationModeHighQualityBicubic As Long = 7
    
    ' Encoder CLSID Constants
    Const CLSID_BMP As String = "{557CF400-1A04-11D3-9A73-0000F81EF32E}"
    Const CLSID_JPG As String = "{557CF401-1A04-11D3-9A73-0000F81EF32E}"
    Const CLSID_GIF As String = "{557CF402-1A04-11D3-9A73-0000F81EF32E}"
    Const CLSID_PNG As String = "{557CF406-1A04-11D3-9A73-0000F81EF32E}"
    
    Dim fso As Object
    Dim ext As String
    Dim encoderCLSID As String
    
    ' Determine output format by file extension
    Set fso = CreateObject("Scripting.FileSystemObject")
    ext = LCase$(fso.GetExtensionName(dstPath))
    
    Select Case ext
        Case "png"
            encoderCLSID = CLSID_PNG
        Case "jpg", "jpeg"
            encoderCLSID = CLSID_JPG
        Case "bmp"
            encoderCLSID = CLSID_BMP
        Case "gif"
            encoderCLSID = CLSID_GIF
        Case Else
            Err.Raise Number:=513, Description:="[ConvertImageFormat] [dstPath] Invalid Extension (." & ext & ")"
    End Select
    
    ' Initialize GDI+
    Dim gsi As GDIPlusStartupInput
    Dim token As LongPtr
    gsi.GdiPlusVersion = 1
    If GdiplusStartup(token, gsi) <> 0 Then
        Err.Raise Number:=514, Description:="[ConvertImageFormat] Failed to initialize GDI+"
    End If
    
    On Error GoTo CleanUp
    
    ' Load source image
    Dim hSrcImage As LongPtr
    If GdipLoadImageFromFile(StrPtr(srcPath), hSrcImage) <> 0 Then
        Err.Raise Number:=515, Description:="[ConvertImageFormat] Failed to load source image: " & srcPath
    End If
    
    ' Get original dimensions
    Dim origWidth As Long, origHeight As Long
    GdipGetImageWidth hSrcImage, origWidth
    GdipGetImageHeight hSrcImage, origHeight
    
    ' Determine target dimensions
    Dim targetW As Long, targetH As Long
    targetW = IIf(dstWidth = -1, origWidth, dstWidth)
    targetH = IIf(dstHeight = -1, origHeight, dstHeight)
    
    ' Calculate aspect-ratio adjusted dimensions if required
    Dim finalW As Long, finalH As Long
    If resize And preserveAspectRatio Then
        Dim scaleW As Double, scaleH As Double, scaleFactor As Double
        scaleW = CDbl(targetW) / CDbl(origWidth)
        scaleH = CDbl(targetH) / CDbl(origHeight)
        
        ' Calculate ratio based on min(scaleW, scaleH) when preserving aspect ratio
        If scaleW < scaleH Then
            scaleFactor = scaleW
        Else
            scaleFactor = scaleH
        End If
        
        finalW = CLng(origWidth * scaleFactor)
        finalH = CLng(origHeight * scaleFactor)
        If finalW < 1 Then finalW = 1
        If finalH < 1 Then finalH = 1
    ElseIf resize Then
        finalW = targetW
        finalH = targetH
    Else
        finalW = origWidth
        finalH = origHeight
    End If
    
    ' Create new bitmap and draw resized image if necessary or convert directly
    Dim hDstBitmap As LongPtr
    Dim hGraphics As LongPtr
    
    If resize Or (finalW <> origWidth Or finalH <> origHeight) Then
        GdipCreateBitmapFromScan0 finalW, finalH, 0, PixelFormat32bppARGB, 0, hDstBitmap
        GdipGetImageGraphicsContext hDstBitmap, hGraphics
        GdipSetInterpolationMode hGraphics, InterpolationModeHighQualityBicubic
        GdipDrawImageRectI hGraphics, hSrcImage, 0, 0, finalW, finalH
    Else
        hDstBitmap = hSrcImage
    End If
    
    ' Prepare Encoder CLSID
    Dim tCLSID As GUID
    CLSIDFromString StrPtr(encoderCLSID), tCLSID
    
    ' Save image using temporary file pattern
    Dim tmpPath As String
    tmpPath = dstPath & ".tmp"
    
    If fso.FileExists(tmpPath) Then fso.DeleteFile tmpPath
    
    Dim status As Long
    status = GdipSaveImageToFile(hDstBitmap, StrPtr(tmpPath), tCLSID, 0)
    
    ' Clean up GDI+ image resources prior to moving file
    If hGraphics <> 0 Then GdipDeleteGraphics hGraphics
    If hDstBitmap <> 0 And hDstBitmap <> hSrcImage Then GdipDisposeImage hDstBitmap
    If hSrcImage <> 0 Then GdipDisposeImage hSrcImage
    
    hGraphics = 0
    hDstBitmap = 0
    hSrcImage = 0
    
    If status <> 0 Then
        If fso.FileExists(tmpPath) Then fso.DeleteFile tmpPath
        Err.Raise Number:=516, Description:="[ConvertImageFormat] Failed to save output image."
    End If
    
    ' Overwrite target file
    If fso.FileExists(dstPath) Then fso.DeleteFile dstPath
    fso.MoveFile tmpPath, dstPath

CleanUp:
    ' Ensure GDI+ resources are freed
    If hGraphics <> 0 Then GdipDeleteGraphics hGraphics
    If hDstBitmap <> 0 And hDstBitmap <> hSrcImage Then GdipDisposeImage hDstBitmap
    If hSrcImage <> 0 Then GdipDisposeImage hSrcImage
    If token <> 0 Then GdiplusShutdown token
    
    If Err.Number <> 0 Then
        Dim errNum As Long, errDesc As String
        errNum = Err.Number
        errDesc = Err.Description
        Err.Raise errNum, Description:=errDesc
    End If
End Sub


Private Function FileToBase64(ByVal filePath As String) As String
    ' Converts a file to a Base64-encoded string.
    Dim stream As Object
    Dim bytes() As Byte
    Dim emptyBytes() As Byte
    emptyBytes = VBA.vbNullString ' Empty Byte Array
    
    ' Load file as binary
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 ' binary
    stream.Open
    stream.LoadFromFile filePath
    
    If stream.Size > 0 Then
        bytes = stream.Read
    Else
        bytes = emptyBytes
    End If
    
    stream.Close
    Set stream = Nothing
    
    ' Convert binary to Base64
    FileToBase64 = BytesToBase64(bytes)

End Function

Private Function BytesToBase64(ByRef bytes() As Byte) As String
    ' Converts a byte array to a Base64 string.
    ' Implemented using the Windows API for high performance.
    ' The implementation supports byte arrays up to 1,610,612,733 bytes by design,
    ' but in practice, memory limitations may be reached with smaller arrays.
    Dim cch As Long
    Dim requiredCharCount As Long
    Dim buffer As String
    Dim dwFlags As Long
    Dim charCountOffset As Long
    Dim dataSize As Double
    Const STRING_FUNCTION_MAXIMUM_LIMIT As Long = 1073741823
    Const STRING_VARIABLE_MAXIMUM_LIMIT As Long = 2147483645
    Const MAXIMUM_DATA_SIZE As Long = 1610612733 ' The largest value that ensures the Base64-encoded size does not exceed the String type's maximum length of 2,147,483,645 characters.
    Const CRYPT_STRING_BASE64 As Long = &H1
    Const CRYPT_STRING_NOCRLF As Long = &H40000000

    dwFlags = CRYPT_STRING_BASE64 Or CRYPT_STRING_NOCRLF
    dataSize = 0
    On Error Resume Next
    dataSize = UBound(bytes) - LBound(bytes) + 1
    On Error GoTo 0
    
    If dataSize = 0 Then
        BytesToBase64 = ""
        Exit Function
    End If
    
    If dataSize > MAXIMUM_DATA_SIZE Then
        Err.Raise Number:=9000, Description:="Error: Data size (" & dataSize & ") exceeds the maximum limit (" & MAXIMUM_DATA_SIZE & ")."
    End If
    
    ' Get required output length (includes terminating null)
    If CryptBinaryToStringW( _
            VarPtr(bytes(LBound(bytes))), _
            UBound(bytes) - LBound(bytes) + 1, _
            dwFlags, _
            0, _
            cch) = 0 Then
        Err.Raise vbObjectError + 1, , "CryptBinaryToStringW failed."
    End If
    
    If 0 >= cch Then
        Err.Raise Number:=9500, Description:="Error: Overflow occurred because the data size is too large"
    End If
    
    ' CryptBinaryToStringW returns the required character count including
    ' the terminating null character. VBA strings (BSTR) store their length
    ' separately, so the null terminator is not part of the string length.
    ' Allocate one fewer character to avoid including the terminating null
    ' in the resulting VBA string.
    requiredCharCount = cch - 1
    
    If requiredCharCount > STRING_VARIABLE_MAXIMUM_LIMIT Then
        Err.Raise Number:=10000, Description:="Error: Data size (" & requiredCharCount & ") exceeds the maximum limit (" & STRING_VARIABLE_MAXIMUM_LIMIT & ")."
    End If
    
    ' Allocate output buffer
    'The String$ function is limited to 1,073,741,823 characters per call,
    'although a String variable can hold up to 2,147,483,645 characters.
    'Allocate larger buffers by concatenating multiple String$ calls.
    If requiredCharCount > STRING_FUNCTION_MAXIMUM_LIMIT Then
        charCountOffset = requiredCharCount - STRING_FUNCTION_MAXIMUM_LIMIT
        buffer = String$(requiredCharCount - charCountOffset, vbNullChar) & String$(charCountOffset, vbNullChar)
    Else
        buffer = String$(requiredCharCount, vbNullChar)
    End If
    

    ' Convert
    If CryptBinaryToStringW( _
            VarPtr(bytes(LBound(bytes))), _
            UBound(bytes) - LBound(bytes) + 1, _
            dwFlags, _
            StrPtr(buffer), _
            cch) = 0 Then
        Err.Raise vbObjectError + 2, , "CryptBinaryToStringW failed."
    End If
    
    ' Windows XP does not support CRYPT_STRING_NOCRLF with CryptBinaryToStringW, so the generated Base64 string may contain line breaks.
    ' Therefore, if the string contains CRLF, replace them with an empty string to remove them.
    If InStr(buffer, VBA.vbCrLf) > 0 Then
        buffer = VBA.Replace$(buffer, VBA.vbCrLf, "")
    End If
    
    BytesToBase64 = buffer
End Function

Private Function ContainsValue(ByVal itemList As Variant, ByVal value As Variant) As Boolean
    ' Check if a specific value exists in Array/Collection/Dictionary
    ' itemList - Array/Collection/Dictionary to search
    ' value - value to check
    ' Performs strict type comparison for non-numeric values
    ' Nested arrays are not supported. Objects are compared by reference
    ' Dependency: IsStrictlyEqual(helper function)
    Dim item As Variant
    Dim temp As Variant
    If LCase$(TypeName(itemList)) = "dictionary" Then
        itemList = itemList.items
    End If
    If IsArray(itemList) Then
        On Error GoTo Finally
        ' Uninitialized Array -> False
        temp = LBound(itemList)
        On Error GoTo 0
    End If
    For Each item In itemList
    
        If IsStrictlyEqual(item, value) Then
            ContainsValue = True
            Exit Function
        End If
    Next
Finally:
    ContainsValue = False
    
End Function

Private Function IsStrictlyEqual(ByVal value1 As Variant, ByVal value2 As Variant) As Boolean
    ' Performs a strict equality comparison including data types.
    ' Numeric types (Integer, Long, Double, etc.) are treated as compatible.
    ' Boolean and Date types are NOT treated as numeric.
    Dim t1 As VbVarType, t2 As VbVarType
    t1 = VarType(value1)
    t2 = VarType(value2)
    
    ' Returns True if objects point to the same reference.
    ' Objects are evaluated first to prevent false matches (e.g., Empty vs empty Cells).
    ' (Also applies to variables holding both objects and other data types)
    If IsObject(value1) Or IsObject(value2) Then
        If IsObject(value1) And IsObject(value2) Then
            IsStrictlyEqual = (value1 Is value2)
        End If
        Exit Function
    End If
    
    ' Null / Empty
    If IsNull(value1) Or IsNull(value2) Then
        IsStrictlyEqual = (IsNull(value1) And IsNull(value2))
        Exit Function
    ElseIf IsEmpty(value1) Or IsEmpty(value2) Then
        IsStrictlyEqual = (IsEmpty(value1) And IsEmpty(value2))
        Exit Function
    End If
    
    
    ' Arrays are not supported (Extend if necessary).
    If IsArray(value1) Or IsArray(value2) Then
        IsStrictlyEqual = False
        Exit Function
    End If
    
    ' Error values
    If t1 = vbError Or t2 = vbError Then
        IsStrictlyEqual = (t1 = t2 And value1 = value2)
        Exit Function
    End If
    
    ' String, Date, Boolean
    If (t1 = vbString Or t2 = vbString) Or (t1 = vbDate Or t2 = vbDate) Or (t1 = vbBoolean Or t2 = vbBoolean) Then
        IsStrictlyEqual = (t1 = t2 And value1 = value2)
        Exit Function
    End If
    
    ' Other data types (e.g., Numeric)
    On Error Resume Next
    IsStrictlyEqual = (value1 = value2)
    Exit Function
    On Error GoTo 0
    IsStrictlyEqual = False
End Function

Private Function UserFormSizeToPixel(ByVal ufSize As Double) As Long
    ' Function to convert the size of a UserForm or control to pixels
    ' Excel VBA UserForm dimensions are internally handled as
    ' DPI-independent logical points based on a fixed 96 DPI system.
    ' Therefore, point-to-pixel conversion can be calculated as:
    '     pixel = point * (96 / 72)
    ' and works consistently regardless of the monitor DPI setting.
    Dim pixelSize As Long
    pixelSize = Round(ufSize * (96 / 72))
    UserFormSizeToPixel = pixelSize
End Function

Private Function GenerateUUIDv4() As String
    Dim i As Long
    Dim b(15) As Byte
    Dim s As String
    Dim hexStr As String
    
    ' Initialize random number generator
    Randomize
    
    ' Generate 16 bytes of random values
    For i = 0 To 15
        b(i) = Int(Rnd() * 256)
    Next i
    
    ' Set version (4) (set bits 7-4 to 0100)
    b(6) = (b(6) And &HF) Or &H40
    
    ' Set variant (10xx)
    b(8) = (b(8) And &H3F) Or &H80
    
    ' Convert the 16 bytes to a string (with hyphen format)
    hexStr = ""
    For i = 0 To 15
        hexStr = hexStr & Right$("0" & Hex(b(i)), 2)
        Select Case i
            Case 3, 5, 7, 9
                hexStr = hexStr & "-"
        End Select
    Next i
    
    GenerateUUIDv4 = LCase$(hexStr)
End Function

Private Sub SaveUtf8TextNoBom(ByVal filePath As String, ByVal textData As String)
    ' Save the specified string as UTF-8 without BOM
    Dim textStream As Object
    Dim binaryStream As Object
    Dim bytes() As Byte
    Dim emptyBytes() As Byte
    emptyBytes = VBA.vbNullString ' Empty Byte Array
    
    ' Normalize line endings
    textData = VBA.Replace$(textData, vbCrLf, vbLf)
    textData = VBA.Replace$(textData, vbCr, vbLf)
    textData = VBA.Replace$(textData, vbLf, vbNewLine)
    
    ' Convert to UTF-8 and remove BOM
    Set textStream = CreateObject("ADODB.Stream")
    With textStream
        .Type = 2 ' Text mode
        .Charset = "utf-8"
        .Open
        .WriteText textData
        .position = 0
        .Type = 1 ' Switch to binary mode
        
        If .Size > 0 Then
            bytes = .Read
        Else
            bytes = emptyBytes
        End If
        
        .Close
    End With
    Set textStream = Nothing
    
    ' Remove BOM if present
    If UBound(bytes) >= 2 Then
        If bytes(0) = &HEF And bytes(1) = &HBB And bytes(2) = &HBF Then
            bytes = MidB(bytes, 4) ' Remove BOM (EF BB BF)
        End If
    End If
    
    ' Save file in binary mode
    Set binaryStream = CreateObject("ADODB.Stream")
    With binaryStream
        .Type = 1
        .Open
        .Write bytes
        .SaveToFile filePath, 2
        .Close
    End With
    Set binaryStream = Nothing
End Sub


Private Function GenerateUnsupportedControlMessage(ByVal ctrl As Object) As String
    Const q As String = """"
    GenerateUnsupportedControlMessage = "Control type " & q & TypeName(ctrl) & q & " is not supported."
End Function

Private Function GenerateUnavailableNameMessage(ByVal ctrl As Object) As String
    Const q As String = """"
    GenerateUnavailableNameMessage = "Object Name " & q & ctrl.Name & q & " is not available." & vbLf & "Please use a different name instead."
End Function

Private Function GetFormControlDepth(ByVal ctrl As Object) As Long
    ' Get the hierarchy depth of the control
    Dim depth As Long
    Dim temp As Variant
    depth = 0
    Set temp = ctrl
    Do While True
        If depth Mod 10 = 0 Then DoEvents
        On Error GoTo Finally
        Set temp = temp.Parent
        depth = depth + 1
        On Error GoTo 0
    Loop
Finally:
    
    If Err.Number <> 438 Then
        Err.Raise Number:=Err.Number
    End If
    
    GetFormControlDepth = depth
    
End Function

Private Function SortFormControlsByDepth(ByVal frmControls As Variant) As Collection
    ' Sort the list of UserForm controls in ascending order of hierarchy depth
    Dim tempColl As Collection
    Set tempColl = New Collection
    Dim sortedColl As Collection
    Set sortedColl = New Collection
    Dim ctrl As Variant
    Dim tempArray() As Variant
    Dim depth As Long
    Dim item As Variant
    For Each ctrl In frmControls
        depth = GetFormControlDepth(ctrl)
        tempColl.Add VBA.Array(depth, ctrl)
    Next ctrl
    If tempColl.Count > 0 Then
        tempArray = CollectionToArray(tempColl)
        Call InsertionSortJaggedArray(tempArray, reverse:=False)
        For Each item In tempArray
            sortedColl.Add item(1)
        Next item
    End If
    Set SortFormControlsByDepth = sortedColl
End Function


Private Function CollectionToArray(ByVal coll As Collection, Optional ByVal isStartIdx1 As Boolean = False) As Variant()
    ' Convert a Collection to an array
    ' If isStartIdx1 is True, create an array starting from index 1 (to match Collection numbering)
    Dim arr() As Variant
    Dim item As Variant
    Dim idx As Long
    If coll.Count > 0 Then
        If isStartIdx1 Then
            ReDim arr(1 To coll.Count)
        Else
            ReDim arr(0 To coll.Count - 1)
        End If
        idx = LBound(arr)
        For Each item In coll
            ' Use "Set" when assigning objects.
            If IsObject(item) Then
                Set arr(idx) = item
            Else
                arr(idx) = item
            End If
            idx = idx + 1
        Next
    Else
        arr = VBA.Array()
    End If
    CollectionToArray = arr
End Function


Private Sub InsertionSortJaggedArray(ByRef arr As Variant, _
    Optional ByVal reverse As Boolean = False, _
    Optional ByVal strSort As Boolean = False, _
    Optional ByVal ignoreCase As Boolean = True)
    
    ' Sorts a jagged array using the Insertion Sort algorithm based on the first element of each nested array.
    '   e.g., [[1, "A"], [3, "B"], [2, "C"]] -> [[1, "A"], [2, "C"], [3, "B"]]
    '   Does not affect the relative order of items with the same numeric value
    '   e.g., [[3, "C"], [3, "A"], [1, "A"], [3, "B"]] -> [[1, "A"], [3, "C"], [3, "A"], [3, "B"]]
    ' reverse: Set to True for descending order.
    '   e.g., [[1, "A"], [3, "B"], [2, "C"]] -> [[3, "B"], [2, "C"], [1, "A"]]
    ' strSort: Set to True for string-based comparison, False for numeric comparison.
    ' ignoreCase: Valid only when strSort is True. Set to True to perform case-insensitive comparison.
    ' Dependency: DynamicCompare
    If Not IsArray(arr) Then Err.Raise Number:=13
    Dim minIndex As Long
    Dim maxIndex As Long
    Dim idxToRef1 As Long
    Dim idxToRef2 As Long
    Dim op As String
    
    If reverse Then
        op = "<"
    Else
        op = ">"
    End If
    
    minIndex = LBound(arr)
    maxIndex = UBound(arr)
    Dim i As Long, j As Long
    Dim swap As Variant
    For i = minIndex + 1 To maxIndex
        swap = arr(i)
        For j = i - 1 To minIndex Step -1
            idxToRef1 = LBound(arr(j))
            idxToRef2 = LBound(swap)
            If DynamicCompare(arr(j)(idxToRef1), swap(idxToRef2), op, strSort, ignoreCase) Then
                arr(j + 1) = arr(j)
            Else
                Exit For
            End If
        Next
        arr(j + 1) = swap
    Next
End Sub


Private Function DynamicCompare(ByVal a As Variant, ByVal b As Variant, ByVal op As String, _
    Optional ByVal shouldStrComp As Boolean = False, Optional ByVal ignoreCase As Boolean = True) As Boolean
    ' Performs dynamic comparison using a string representation of an operator.
    ' a, b: Values to compare.
    ' op: Comparison operator as a string (">", ">=", "<", "<=", "=", "<>").
    ' shouldStrComp: Set to True for string comparison mode, False for numeric/default comparison.
    ' ignoreCase: Valid only when shouldStrComp is True. Set to True to ignore case sensitivity.
    Dim result As Boolean
    Dim compareMode As VbCompareMethod
    
    If shouldStrComp Then
        If ignoreCase Then
            compareMode = vbTextCompare
        Else
            compareMode = vbBinaryCompare
        End If
        
        Select Case op
            Case ">"
                result = StrComp(a, b, compareMode) > 0
            Case ">="
                result = StrComp(a, b, compareMode) >= 0
            Case "<"
                result = StrComp(a, b, compareMode) < 0
            Case "<="
                result = StrComp(a, b, compareMode) <= 0
            Case "="
                result = StrComp(a, b, compareMode) = 0
            Case "<>"
                result = StrComp(a, b, compareMode) <> 0
            Case Else
                Err.Raise vbObjectError, , "Unknown operator: " & op
        End Select
    Else
        Select Case op
            Case ">"
                result = (a > b)
            Case ">="
                result = (a >= b)
            Case "<"
                result = (a < b)
            Case "<="
                result = (a <= b)
            Case "="
                result = (a = b)
            Case "<>"
                result = (a <> b)
            Case Else
                Err.Raise vbObjectError, , "Unknown operator: " & op
        End Select
    End If
    DynamicCompare = result
End Function

Private Function CollContainsKey(ByVal coll As Collection, ByVal strKey As String) As Boolean
    ' Check if a specific key exists in the Collection
    CollContainsKey = False
    If coll Is Nothing Then Exit Function
    If coll.Count = 0 Then Exit Function
     
    On Error GoTo Exception
    Call coll.item(strKey)
    On Error GoTo 0
    CollContainsKey = True
    
    Exit Function
Exception:
    CollContainsKey = False
    Exit Function
End Function

Private Function JoinCollection(ByVal coll As Collection, Optional ByVal delimiter As String = "") As String
    Dim arr() As Variant
    Dim result As String
    arr = CollectionToArray(coll)
    result = Join(arr, delimiter)
    JoinCollection = result
End Function
