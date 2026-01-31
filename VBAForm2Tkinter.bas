Attribute VB_Name = "VBAForm2Tkinter"

' VBAForm2Tkinter v1.2.2
' https://github.com/GUI-Conversion-Tools/VBAForm2Tkinter
' Copyright (c) 2025-2026 ZeeZeX
' This software is released under the MIT License.
' https://opensource.org/licenses/MIT

Option Explicit


#If VBA7 Then
    ' 64bit Office / VBA7 or later
    Private Declare PtrSafe Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function FindWindowW Lib "user32" (ByVal lpClassName As LongPtr, ByVal lpWindowName As LongPtr) As LongPtr
    Private Declare PtrSafe Function GetClientRect Lib "user32" (ByVal hwnd As LongPtr, lpRect As RECT) As Long
    Private Declare PtrSafe Function GetWindowRect Lib "user32" (ByVal hwnd As LongPtr, lpRect As RECT) As Long
    Private Type RECT: Left As Long: Top As Long: Right As Long: Bottom As Long: End Type
    Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
    Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
    Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
#Else
    ' 32bit Office
    Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
    Private Declare Function FindWindowW Lib "user32" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long
    Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
    Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
    Private Type RECT: Left As Long: Top As Long: Right As Long: Bottom As Long: End Type
    Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
    Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
    Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
#End If



Sub TestRunConversion2Tk()
    Call ConvertForm2Tkinter(UserForm1)
End Sub


Sub ConvertForm2Tkinter(ByVal frm As Object)
    Dim code As String
    Dim filePath As String
    Dim saveDir As String
    code = VBAForm2Tkinter(frm)
    If code <> "" Then
        If ThisWorkbook.Path = "" Then
            saveDir = "C:"
        Else
            saveDir = ThisWorkbook.Path
        End If
        filePath = saveDir & "\output.py"
        Call SaveUTF8Text_NoBOM(filePath, code)
        MsgBox "Saved: " & filePath
    Else
        MsgBox "Conversion failed."
    End If
    
End Sub


Function VBAForm2Tkinter(ByVal root As Object) As String
    Dim ctrl As MSForms.Control
    Dim ctrls As Collection
    Dim item As Variant
    Dim r As String
    Const q As String = """"
    Dim fontStyle As String
    Dim fontOpts As String
    Dim widgetType As String
    Dim styleName As String
    Dim sizeFactorsAndOffsets() As Variant
    Dim sizeFactorX As Double
    Dim sizeFactorY As Double
    Dim pixelWidth As Long
    Dim pixelHeight As Long
    Dim pixelTop As Long
    Dim pixelLeft As Long
    Dim i As Long
    Dim orientation As String
    Dim cursorType As String
    Dim caption As String
    Dim dpis() As Variant
    Dim scaleFactorX As Double
    Dim scaleFactorY As Double
    Dim colorCode As String
    
    r = ""
    
    dpis = GetPrimaryMonitorDPI
    scaleFactorX = dpis(0) / 96
    scaleFactorY = dpis(1) / 96
    
    ' Get factor for size conversion
    sizeFactorsAndOffsets = GetUserFormScaleFactorsAndOffsets(root)
    sizeFactorX = sizeFactorsAndOffsets(0)
    sizeFactorY = sizeFactorsAndOffsets(1)
    ' Convert UserForm's size to pixel size
    pixelWidth = UserFormSizeToPixel(root.Width, sizeFactorX)
    pixelHeight = UserFormSizeToPixel(root.Height, sizeFactorY)
    pixelWidth = pixelWidth - sizeFactorsAndOffsets(2)
    pixelHeight = pixelHeight - sizeFactorsAndOffsets(3)
    ' Divide window size by scaling factor
    pixelWidth = Round(pixelWidth / scaleFactorX)
    pixelHeight = Round(pixelHeight / scaleFactorY)
    
    r = r & "import tkinter as tk" & vbLf
    r = r & "from tkinter import ttk" & vbLf
    r = r & "from tkinter import font" & vbLf
    r = r & vbLf
    r = r & root.Name & " = " & "tk.Tk()" & vbLf
    caption = root.caption
    caption = Convert2PythonFormatText(caption)
    r = r & root.Name & ".title(" & q & caption & q & ")" & vbLf
    r = r & root.Name & ".geometry(" & q & pixelWidth & "x" & pixelHeight & q & ")" & vbLf
    r = r & root.Name & ".resizable(False, False)" & vbLf
    r = r & root.Name & ".configure(bg=" & q & FormColorToHex(root.BackColor) & q & ")" & vbLf
    r = r & GetBorderSetting(root) & vbLf
    
    cursorType = GetControlCursorType(root)
    If cursorType <> "" Then
        r = r & root.Name & ".configure(cursor=" & q & cursorType & q & ")" & vbLf
    Else
        r = r & root.Name & ".configure(cursor=" & "None" & ")" & vbLf
    End If
    
    r = r & vbLf
    r = r & "style = ttk.Style()" & vbLf
    r = r & "style.theme_use('default')" & vbLf
    r = r & vbLf
    Set ctrls = SortFormControlsByDepth(root.Controls)
    For Each ctrl In ctrls
        If GetTkWidgetName(ctrl) <> "" Then
            widgetType = GetTkWidgetName(ctrl)
            
            ' For ttk widget
            If ContainsValue(Array("Combobox", "Notebook", "Scale"), widgetType) Then
                widgetType = "ttk." & widgetType
            Else
                widgetType = "tk." & widgetType
            End If
            
            pixelLeft = UserFormSizeToPixel(ctrl.Left, sizeFactorX)
            pixelTop = UserFormSizeToPixel(ctrl.Top, sizeFactorY)
            pixelWidth = UserFormSizeToPixel(ctrl.Width, sizeFactorX)
            pixelHeight = UserFormSizeToPixel(ctrl.Height, sizeFactorY)
            
            pixelLeft = Round(pixelLeft / scaleFactorX)
            pixelTop = Round(pixelTop / scaleFactorY)
            pixelWidth = Round(pixelWidth / scaleFactorX)
            pixelHeight = Round(pixelHeight / scaleFactorY)
            
            r = r & ctrl.Name & " = " & widgetType & "(" & ctrl.Parent.Name & ")" & vbLf
            r = r & ctrl.Name & ".place(x=" & pixelLeft & ", y=" & pixelTop & ", width=" & pixelWidth & ", height=" & pixelHeight & ")" & vbLf
            
            If GetTkWidgetName(ctrl) = "LabelFrame" Or Not ContainsValue(Array("ComboBox", "Frame", "Image", "ScrollBar", "MultiPage"), TypeName(ctrl)) Then
                ' Set ForeColor
                r = r & ctrl.Name & ".configure(fg=" & q & FormColorToHex(ctrl.ForeColor) & q & ")" & vbLf
            End If
            
            If Not ContainsValue(Array("ComboBox", "MultiPage", "ScrollBar"), TypeName(ctrl)) Then
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
                r = r & ctrl.Name & ".configure(bg=" & q & colorCode & q & ")" & vbLf
                
                If ContainsValue(Array("CommandButton", "CheckBox", "ToggleButton", "OptionButton"), TypeName(ctrl)) Then
                    ' Set the colors when the button is pressed
                    r = r & ctrl.Name & ".configure(activeforeground=" & q & FormColorToHex(ctrl.ForeColor) & q & ")" & vbLf
                    r = r & ctrl.Name & ".configure(activebackground=" & q & colorCode & q & ")" & vbLf
                End If
                
                If TypeName(ctrl) = "ToggleButton" Then
                    r = r & ctrl.Name & ".configure(indicatoron=0)" & vbLf
                    r = r & ctrl.Name & ".configure(selectcolor=" & q & colorCode & q & ")" & vbLf
                End If
                
            End If
            
            If GetTkWidgetName(ctrl) = "LabelFrame" Or ContainsValue(Array("Label", "CommandButton", "CheckBox", "ToggleButton", "OptionButton"), TypeName(ctrl)) Then
                caption = ctrl.caption
                caption = Convert2PythonFormatText(caption)
                r = r & ctrl.Name & ".configure(text=" & q & caption & q & ")" & vbLf
            End If
            
            
            If TypeName(ctrl) = "TextBox" Then
                If GetTkWidgetName(ctrl) = "Entry" Then
                    r = r & ctrl.Name & ".insert(0, " & q & Convert2PythonFormatText(ctrl.text) & q & ")" & vbLf
                ElseIf GetTkWidgetName(ctrl) = "Text" Then
                    r = r & ctrl.Name & ".insert(" & q & "1.0" & q & ", " & q & Convert2PythonFormatText(ctrl.text) & q & ")" & vbLf
                End If
            End If
            
            If TypeName(ctrl) = "ComboBox" Then
                styleName = ctrl.Name & "_style" & ".TCombobox"
                r = r & "style.configure(" & q & styleName & q & ", foreground=" & q & FormColorToHex(ctrl.ForeColor) & q & ")" & vbLf
                colorCode = FormColorToHex(ctrl.BackColor)
                
                If ctrl.BackStyle = fmBackStyleTransparent Then
                    If TypeName(ctrl.Parent) <> "Page" Then
                        colorCode = FormColorToHex(ctrl.Parent.BackColor)
                    Else
                        colorCode = FormColorToHex(&H8000000F)
                    End If
                End If
                
                r = r & "style.configure(" & q & styleName & q & ", fieldbackground=" & q & colorCode & q & ")" & vbLf
                r = r & ctrl.Name & ".configure(style=" & q & styleName & q & ")" & vbLf
                r = r & ctrl.Name & "_items_value = " & GetListBoxValue(ctrl) & vbLf
                r = r & ctrl.Name & ".configure(value=" & ctrl.Name & "_items_value" & ")" & vbLf
                r = r & ctrl.Name & ".set(" & q & Convert2PythonFormatText(ctrl.text) & q & ")" & vbLf
            End If
            
            If TypeName(ctrl) = "ListBox" Then
                r = r & ctrl.Name & "_items_value = " & GetListBoxValue(ctrl) & vbLf
                r = r & ctrl.Name & ".insert(tk.END, " & "*" & ctrl.Name & "_items_value" & ")" & vbLf
            End If
            
            If TypeName(ctrl) = "ScrollBar" Then
                Select Case ctrl.orientation
                    Case -1
                        If ctrl.Width > ctrl.Height Then
                            orientation = "Horizontal"
                        Else
                            orientation = "Vertical"
                        End If
                        
                    Case 0
                        orientation = "Vertical"
                    Case 1
                        orientation = "Horizontal"
                    Case Else
                        orientation = "Vertical"
                End Select
                r = r & ctrl.Name & ".configure(from_=" & ctrl.Min & ", to=" & ctrl.Max & ",orient=" & q & LCase(orientation) & q & ")" & vbLf
                styleName = ctrl.Name & "_style" & "." & orientation & ".TScale"
                r = r & "style.configure(" & q & styleName & q & ", background=" & q & FormColorToHex(ctrl.BackColor) & q & ")" & vbLf
                r = r & ctrl.Name & ".configure(style=" & q & styleName & q & ")" & vbLf
            End If
            
            ' Set each Caption and font in MultiPage, font size is rounded
            If TypeName(ctrl) = "MultiPage" Then
                For Each item In ctrl.Pages
                    caption = item.caption
                    caption = Convert2PythonFormatText(caption)
                    r = r & item.Name & " = tk.Frame(" & ctrl.Name & ")" & vbLf
                    r = r & ctrl.Name & ".add(" & item.Name & ", text=" & q & caption & q & ")" & vbLf
                Next
                
                
                fontStyle = ""
                fontOpts = ""
                
                If ctrl.Font.Bold Then fontStyle = fontStyle & ", weight=" & q & "bold" & q
                If ctrl.Font.Italic Then fontStyle = fontStyle & ", slant=" & q & "italic" & q
                If ctrl.Font.Underline Then fontOpts = fontOpts & ", underline=1"
                If ctrl.Font.Strikethrough Then fontOpts = fontOpts & ", overstrike=1"
                
                
                styleName = ctrl.Name & "_style" & ".Tab"
                r = r & "style.configure(" & q & styleName & q & ", foreground=" & q & FormColorToHex(ctrl.ForeColor) & q & ")" & vbLf
                r = r & "style.configure(" & q & styleName & q & ", background=" & q & FormColorToHex(ctrl.BackColor) & q & ")" & vbLf
                r = r & "style.configure(" & q & styleName & q & ",font=font.Font(family=" & q & ctrl.Font.Name & q & ", size=" & Round(ctrl.Font.Size) & fontStyle & fontOpts & "))" & vbLf
                r = r & ctrl.Name & ".configure(style=" & q & styleName & q & ")" & vbLf
            End If
            
            
            ' Font size is rounded because Tkinter does not support floats in font settings
            If GetTkWidgetName(ctrl) = "LabelFrame" Or Not ContainsValue(Array("Frame", "ScrollBar", "Image", "SpinButton", "MultiPage"), TypeName(ctrl)) Then
                fontStyle = ""
                fontOpts = ""
                
                If ctrl.Font.Bold Then fontStyle = fontStyle & ", weight=" & q & "bold" & q
                If ctrl.Font.Italic Then fontStyle = fontStyle & ", slant=" & q & "italic" & q
                If ctrl.Font.Underline Then fontOpts = fontOpts & ", underline=1"
                If ctrl.Font.Strikethrough Then fontOpts = fontOpts & ", overstrike=1"
                
                r = r & ctrl.Name & ".configure(font=font.Font(family=" & q & ctrl.Font.Name & q & ", size=" & Round(ctrl.Font.Size) & fontStyle & fontOpts & "))" & vbLf
            End If
            
            
            If ContainsValue(Array("Frame", "TextBox", "Label", "ListBox", "Image"), TypeName(ctrl)) Then
                ' Tkinter's Combobox does not support customizing border colors or relief
                r = r & GetBorderSetting(ctrl) & vbLf
            End If
            
            If GetTkWidgetName(ctrl) <> "Text" And ContainsValue(Array("Label", "TextBox", "ComboBox", "CheckBox", "ToggleButton", "OptionButton"), TypeName(ctrl)) Then
                r = r & GetTextAlignSetting(ctrl) & vbLf
            End If
            
            ' Set mouse cursor
            If TypeName(ctrl) <> "MultiPage" Then
                cursorType = GetControlCursorType(ctrl)
                If cursorType <> "" Then
                    r = r & ctrl.Name & ".configure(cursor=" & q & cursorType & q & ")" & vbLf
                Else
                    r = r & ctrl.Name & ".configure(cursor=" & "None" & ")" & vbLf
                End If
            End If
            
            If TypeName(ctrl) = "Image" Then
                r = r & "#" & ctrl.Name & "_photo = tk.PhotoImage(file=r" & q & q & ")" & vbLf
                r = r & "#" & ctrl.Name & ".create_image(0, 0, image=" & ctrl.Name & "_photo" & ", anchor=tk.NW)" & vbLf
            End If
            
            r = r & vbLf
            
        Else
            MsgBox GenerateUnsupportedControlMessage(ctrl)
            r = ""
            VBAForm2Tkinter = r
            Exit Function
        End If
    Next ctrl
    r = r & SetTkRadiobuttonValues(ctrls) & vbLf
    r = r & SetTkCheckbuttonValues(ctrls) & vbLf
    r = r & root.Name & ".mainloop()"
    VBAForm2Tkinter = r
End Function

Private Function GetBorderSetting(ByVal ctrl As Object) As String
    Dim r As String
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
        Case 1
            ' SpecialEffect is 0 if BorderStyle is 1
            borderWidth = 0
            highlightBorderWidth = 1
            relief = "tk.FLAT"
        Case 0
            Select Case ctrl.SpecialEffect
                Case 0
                    borderWidth = 0
                    relief = "tk.FLAT"
                Case 1
                    relief = "tk.RAISED"
                Case 2
                    relief = "tk.SUNKEN"
                Case 3
                    relief = "tk.GROOVE"
                Case 6
                    relief = "tk.RIDGE"
            End Select
    End Select

    r = ctrl.Name & ".configure(relief=" & relief & ", bd=" & borderWidth & ", highlightthickness=" & highlightBorderWidth & ", highlightbackground=" & q & hexBorderColor & q & ", highlightcolor=" & q & hexBorderColor & q & ")"
    GetBorderSetting = r
End Function

Private Function GetTextAlignSetting(ByVal ctrl As Object) As String
   Dim r As String
   Const q As String = """"
   Dim anchor As String
   Dim justify As String
   r = ""
   
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
    If Not ContainsValue(Array("TextBox", "ComboBox"), TypeName(ctrl)) Then
        r = ctrl.Name & ".configure(anchor=" & q & anchor & q & ")" & vbLf
    End If
    r = r & ctrl.Name & ".configure(justify=" & q & justify & q & ")"
    GetTextAlignSetting = r
End Function

Private Function GetTkWidgetName(ByVal ctrl As Object) As String
    Dim r As String
    Select Case TypeName(ctrl)
        Case "Label"
            r = "Label"
        Case "CommandButton"
            r = "Button"
        Case "Frame"
            If ctrl.caption = "" Then
                r = "Frame"
            Else
                r = "LabelFrame"
            End If
        Case "TextBox"
            If ctrl.MultiLine Then
                r = "Text"
            Else
                r = "Entry"
            End If
        Case "SpinButton"
            r = "Spinbox"
        Case "ListBox"
            r = "Listbox"
        Case "CheckBox"
            r = "Checkbutton"
        Case "ToggleButton"
            r = "Checkbutton"
        Case "OptionButton"
            r = "Radiobutton"
        Case "Image"
            r = "Canvas"
        Case "ScrollBar"
            r = "Scale"
        Case "ComboBox"
            r = "Combobox"
        Case "MultiPage"
            r = "Notebook"
        Case Else
            r = ""
    End Select
    GetTkWidgetName = r
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

Private Function SetTkRadiobuttonValues(ByVal ctrls As Variant) As String
    Dim parentList As New Collection
    Const q As String = """"
    Dim varName As String
    Dim ctrl As Variant
    Dim r As String
    r = ""
    For Each ctrl In ctrls
        If TypeName(ctrl) = "OptionButton" Then
            varName = ctrl.Parent.Name & "_radiobutton_value"
            If Not CollContainsKey(parentList, ctrl.Parent.Name) Then
                ' Use the Collection key to check and avoid redeclaring a variable that has already been declared
                parentList.Add "", ctrl.Parent.Name
                r = r & varName & "= tk.StringVar()" & vbLf
                r = r & varName & ".set(None)" & vbLf ' Deselect the radio button
            End If
            r = r & ctrl.Name & ".configure(variable=" & varName & ", value=" & q & ctrl.Name & q & ")" & vbLf
            If ctrl.value = True Then
                r = r & varName & ".set(" & q & ctrl.Name & q & ")" & vbLf
            End If
            
        End If
    Next
    SetTkRadiobuttonValues = r
End Function

Private Function SetTkCheckbuttonValues(ByVal ctrls As Variant) As String
    Dim varName As String
    Dim ctrl As Variant
    Dim value As Boolean
    Dim r As String
    r = ""
    For Each ctrl In ctrls
        If TypeName(ctrl) = "CheckBox" Or TypeName(ctrl) = "ToggleButton" Then
            varName = ctrl.Name & "_checkbutton_value"
            r = r & varName & "= tk.BooleanVar()" & vbLf
            r = r & ctrl.Name & ".configure(variable=" & varName & ")" & vbLf
            If ctrl.value = True Then
                value = True
            Else
                value = False
            End If
            r = r & varName & ".set(" & value & ")" & vbLf
            
        End If
    Next
    SetTkCheckbuttonValues = r
End Function

Private Function GetListBoxValue(ByVal ctrl As Object) As String
    ' Retrieve the items of a ListBox or ComboBox as a string in the format ["1", "2", "3"].
    Const q As String = """"
    Dim item As Variant
    Dim i As Long: i = 0
    Dim r As String
    Const indent As String = "    "
    Const maxItemsPerLine As Long = 3
    r = ""
    If ctrl.ListCount > 0 Then
        If ctrl.ListCount > maxItemsPerLine Then r = r & vbLf & indent
        For Each item In ctrl.List
            i = i + 1
            r = r & q & Convert2PythonFormatText(item) & q
            If Not i = ctrl.ListCount Then
                r = r & ", "
                If i Mod maxItemsPerLine = 0 And ctrl.ListCount > maxItemsPerLine Then r = r & vbLf & indent
            Else
                If ctrl.ListCount > maxItemsPerLine Then r = r & vbLf
                Exit For
            End If
        Next item
    End If
    r = "[" & r & "]"
    GetListBoxValue = r
End Function

Private Function Convert2PythonFormatText(ByVal text As String) As String
    ' Escape special characters in the string
    text = VBA.Replace(text, "\", "\\")
    text = VBA.Replace(text, """", "\" & """")
    text = VBA.Replace(text, "'", "\" & "'")
    ' Convert VBA line breaks to Python format
    ' vbCrLf should be replaced first
    text = VBA.Replace(text, vbCrLf, vbLf)
    text = VBA.Replace(text, vbCr, vbLf)
    text = VBA.Replace(text, vbLf, "\n")
    Convert2PythonFormatText = text
End Function

Private Function FormColorToHex(ByVal clr As Long) As String
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
                     Right("0" & Hex(r), 2) & _
                     Right("0" & Hex(g), 2) & _
                     Right("0" & Hex(b), 2)
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
    If LCase(TypeName(itemList)) = "dictionary" Then
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

Private Function Win32_FindWindowW(ByVal className As String, ByVal windowTitle As String) As LongPtr
    ' Get the window's hwnd
    ' className: The window's class name (exact match). If not specified, provide "", Empty, or vbNullString
    ' windowTitle: The window's title (exact match). If not specified, provide "", Empty, or vbNullString
    ' Example: Get Excel's main window by specifying only the class name
    ' hwnd = Win32_FindWindowW("XLMAIN", Empty)
    Dim hwnd As LongPtr
    If className = "" Then className = vbNullString
    If windowTitle = "" Then windowTitle = vbNullString
    hwnd = FindWindowW(StrPtr(className), StrPtr(windowTitle))
    Win32_FindWindowW = hwnd
End Function

Private Function GetUserFormScaleFactorsAndOffsets(ByVal frm As Object) As Variant()
    ' Function to get the factors and offsets for converting a UserForm's size to pixel units
    ' Obtains the window size in pixels via Windows API and compares it with the UserForm's design size
    Dim clRect As RECT
    Dim winRect As RECT
    Dim pixClWidth As Long, pixClHeight As Long
    Dim pixWinWidth As Long, pixWinHeight As Long
    Dim pixWidthOffset As Long, pixHeightOffset As Long
    Dim scaleX As Double, scaleY As Double
    Dim hwnd As LongPtr
    Dim originalFrmTitle As String
    Dim tempFrmTitle As String
    Dim results(0 To 3) As Variant
    
    ' To avoid getting the handle of a window with the same name, temporarily change the title to a unique name when obtaining hwnd
    ' Restore the original title immediately after obtaining hwnd
    originalFrmTitle = frm.caption
    tempFrmTitle = "TempName_" & GenerateUUIDv4()
    frm.caption = tempFrmTitle
    hwnd = Win32_FindWindowW("", tempFrmTitle)
    frm.caption = originalFrmTitle
    
    If CLng(hwnd) = 0 Then
        Err.Raise Number:=513, Description:="Failed to get HWND."
    End If
    
    ' Get the actual client area size
    GetClientRect hwnd, clRect
    pixClWidth = clRect.Right - clRect.Left
    pixClHeight = clRect.Bottom - clRect.Top
    
    ' Get the difference in X and Y between the actual window size and the client area size
    GetWindowRect hwnd, winRect
    pixWinWidth = winRect.Right - winRect.Left
    pixWinHeight = winRect.Bottom - winRect.Top
    pixWidthOffset = pixWinWidth - pixClWidth
    pixHeightOffset = pixWinHeight - pixClHeight
    
    ' Twips -> pixel conversion factors
    scaleX = pixClWidth / frm.InsideWidth
    scaleY = pixClHeight / frm.InsideHeight
    
    ' If horizontal and vertical scales are almost the same, return the average
    If Abs(scaleX - scaleY) < 0.01 Then
        results(0) = (scaleX + scaleY) / 2
        results(1) = (scaleX + scaleY) / 2
    Else
        ' If there is a difference between horizontal and vertical scales
        results(0) = scaleX
        results(1) = scaleY
    End If
    results(2) = pixWidthOffset
    results(3) = pixHeightOffset
    GetUserFormScaleFactorsAndOffsets = results
End Function

Private Function UserFormSizeToPixel(ByVal ufSize As Double, ByVal factor As Double) As Long
    ' Function to convert the size of a UserForm or control to pixels
    UserFormSizeToPixel = Round(ufSize * factor)
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

Private Sub SaveUTF8Text_NoBOM(ByVal filePath As String, ByVal textData As String)
    ' Save the specified string as UTF-8 without BOM
    Dim stream As Object
    Dim bytes() As Byte
    
    ' Normalize line endings
    textData = VBA.Replace(textData, vbCrLf, vbLf)
    textData = VBA.Replace(textData, vbCr, vbLf)
    textData = VBA.Replace(textData, vbLf, vbNewLine)
    
    ' Convert to UTF-8 and remove BOM
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2 ' Text mode
    stream.Charset = "utf-8"
    stream.Open
    stream.WriteText textData
    stream.position = 0
    stream.Type = 1 ' Switch to binary mode
    bytes = stream.Read
    stream.Close
    Set stream = Nothing
    
    ' Remove BOM if present
    If UBound(bytes) >= 2 Then
        If bytes(0) = &HEF And bytes(1) = &HBB And bytes(2) = &HBF Then
            bytes = MidB(bytes, 4) ' Remove BOM (EF BB BF)
        End If
    End If
    
    ' Save file in binary mode
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write bytes
    stream.SaveToFile filePath, 2
    stream.Close
    Set stream = Nothing
End Sub

Private Function GetPrimaryMonitorDPI() As Variant()
    Dim hdc As LongPtr
    Dim dpiX As Long, dpiY As Long
    Dim results(0 To 1) As Variant
    Const LOGPIXELSX As Long = 88 ' Horizontal DPI
    Const LOGPIXELSY As Long = 90 ' Vertical DPI
    
    ' Get device context for the entire screen
    hdc = GetDC(0)
    
    ' Get horizontal and vertical DPI
    dpiX = GetDeviceCaps(hdc, LOGPIXELSX)
    dpiY = GetDeviceCaps(hdc, LOGPIXELSY)
    
    ' Release the device context
    ReleaseDC 0, hdc
    
    results(0) = dpiX
    results(1) = dpiY
    
    ' Return DPI
    GetPrimaryMonitorDPI = results
End Function

Private Function GenerateUnsupportedControlMessage(ByVal ctrl As Object) As String
    Const q As String = """"
    GenerateUnsupportedControlMessage = "Control type " & q & TypeName(ctrl) & q & " is not supported."
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
        tempColl.Add Array(depth, ctrl)
    Next ctrl
    If tempColl.Count > 0 Then
        tempArray = Collection2Array(tempColl)
        Call InsertionSortJaggedArray(tempArray, reverse:=False)
        For Each item In tempArray
            sortedColl.Add item(1)
        Next item
    End If
    Set SortFormControlsByDepth = sortedColl
End Function


Private Function Collection2Array(ByVal coll As Collection, Optional ByVal isStartIdx1 As Boolean = False) As Variant()
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
        arr = Array()
    End If
    Collection2Array = arr
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


