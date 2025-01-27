Public ignoreColumns() As Integer
Public colMax As Integer
Public makeIntoFractions As Boolean

Sub GenerateIndividualSizeGuides()
    Dim sizeGuide As String
    makeIntoFractions = Worksheets("Generate").fractioncbx.Value

    For Each sht In Worksheets
        sizeGuide = singleSizeGuide(sht.Name)
    Next
    
    MsgBox "Done"
    
    
End Sub

Function singleSizeGuide(sheetName As String)
    If sheetName = "variables" Or sheetName = "Generate" Then
            Exit Function
    End If
    
    Dim titleCell, brand, titleFull As String
    titleCell = Worksheets(sheetName).Range("B1")
    brand = Worksheets("variables").Range("C7")
    titleFull = brand + " " + titleCell
    
    Dim sizeGuideHTML As String
    sizeGuideHTML = ""
    sizeGuideHTML = sizeGuideHTML + "<div class='size-guide-container'>" + vbNewLine + vbNewLine
    
    sizeGuideHTML = sizeGuideHTML + styleSheetsHTML
    sizeGuideHTML = sizeGuideHTML + logoHTML
    sizeGuideHTML = sizeGuideHTML + titleHTML(titleFull)
    sizeGuideHTML = sizeGuideHTML + unitSelectorHTML(sheetName)
    sizeGuideHTML = sizeGuideHTML + cmGuideHTML(Worksheets(sheetName))
    sizeGuideHTML = sizeGuideHTML + inchGuideHTML(Worksheets(sheetName))
    
    sizeGuideHTML = sizeGuideHTML + "</div>"
    
    Call WriteToFile(titleFull + ".html", brand, sizeGuideHTML)
    
    'singleSizeGuide(0) = titleCell + ".html"
    'singleSizeGuide = sizeGuideHTML
End Function

Function check_variables_sheet() As Boolean
    
    Dim Sheet As Worksheet
    
    For Each Sheet In ThisWorkbook.Worksheets
        If Sheet.Name = "variables" Then
            check_variables_sheet = True
        End If
    Next Sheet
    
    check_variables_sheet = False

End Function


Function styleSheetsHTML() As String
    Dim linkStart As String
    linkStart = "<link rel='stylesheet' type='text/css' href='"
    
    Dim linkEnd As String
    linkEnd = "'>"
    
    Dim localCss, RemoteCss As String
    
    localCss = Worksheets("variables").Range("C2")
    RemoteCss = Worksheets("variables").Range("C3")
    
    Dim returnHTML As String
    
    returnHTML = "<!-- Stylesheets -->" + vbNewLine
    returnHTML = returnHTML + linkStart + localCss + linkEnd + vbNewLine
    returnHTML = returnHTML + linkStart + RemoteCss + linkEnd + vbNewLine
    returnHTML = returnHTML + "<!-- Stylesheets -->" + vbNewLine + vbNewLine
    
    styleSheetsHTML = returnHTML
    
End Function

Function logoHTML() As String
    Dim alt, src, dataImage As String
    
    alt = Worksheets("variables").Range("C4")
    src = Worksheets("variables").Range("C5")
    dataImage = Worksheets("variables").Range("C6")
    
    Dim returnHTML As String
    
    returnHTML = "<!-- Logo -->" + vbNewLine
    returnHTML = returnHTML + "<figure>" + vbNewLine
    returnHTML = returnHTML + vbTab + "<img style='float: right; display: inline-block; max-width: 130px; max-height: 70px;' "
    returnHTML = returnHTML + "alt='" + alt + "' "
    returnHTML = returnHTML + "src='" + src + "' "
    returnHTML = returnHTML + "data-image='" + dataImage + "'>" + vbNewLine
    returnHTML = returnHTML + "</figure>" + vbNewLine
    returnHTML = returnHTML + "<!-- Logo -->" + vbNewLine + vbNewLine
    
    logoHTML = returnHTML
    
End Function

Function titleHTML(t As String) As String
    Dim returnHTML As String
    
    
    
    returnHTML = "<!-- Title -->" + vbNewLine
    returnHTML = returnHTML + "<h5><b>Size Guide</b></h5>" + vbNewLine
    returnHTML = returnHTML + "<h6>" + t + "</h6>" + vbNewLine
    returnHTML = returnHTML + "<!-- Title -->" + vbNewLine + vbNewLine
    
    titleHTML = returnHTML
    
End Function

Function unitSelectorHTML(ByVal t As String) As String
    t = Replace(t, " ", "")
    
    Dim returnHTML As String
    
    returnHTML = "<!-- Unit Selector -->" + vbNewLine
    returnHTML = returnHTML + "<input type='radio' name='unitselector" + t + "' id='inputInch' checked/>" + vbNewLine
    returnHTML = returnHTML + "<input type='radio' name='unitselector" + t + "' id='inputCM' />" + vbNewLine
    returnHTML = returnHTML + vbNewLine
    returnHTML = returnHTML + "<label class='unitBtn' for='inputInch'>Inch</label>" + vbNewLine
    returnHTML = returnHTML + "<label class='unitBtn' for='inputCM'>CM</label>" + vbNewLine
    returnHTML = returnHTML + "<!-- Unit Selector -->" + vbNewLine + vbNewLine
    
    unitSelectorHTML = returnHTML
    
End Function

Function cmGuideHTML(Sheet As Worksheet) As String
    Dim returnHTML As String
    Dim cellValue, unitsValue As String
    Dim row, column As Integer
    
    Dim currentSize  As Integer
    
    
    currentSize = 0
    colMax = 0
    
    returnHTML = "<!-- CM Guide -->" + vbNewLine
    returnHTML = returnHTML + "<span class='cm'>" + vbNewLine
    returnHTML = returnHTML + vbTab + "<div class='size-guide'>" + vbNewLine
    
    
    returnHTML = returnHTML + tableHeaderHTML(Sheet, "inch")
    
    returnHTML = returnHTML + tableBodyHTML(Sheet)
    
    returnHTML = returnHTML + vbTab + "</div>" + vbNewLine
    returnHTML = returnHTML + "</span>" + vbNewLine
    returnHTML = returnHTML + "<!-- CM Guide -->" + vbNewLine + vbNewLine
    
    cmGuideHTML = returnHTML
End Function

Function inchGuideHTML(Sheet As Worksheet) As String
    Dim returnHTML As String
    Dim cellValue, unitsValue As String
    Dim row, column As Integer
    
    Dim currentSize  As Integer
    
    
    currentSize = 0
    colMax = 0
    
    returnHTML = "<!-- Inch Guide -->" + vbNewLine
    returnHTML = returnHTML + "<span class='inch'>" + vbNewLine
    returnHTML = returnHTML + vbTab + "<div class='size-guide'>" + vbNewLine
    
    
    returnHTML = returnHTML + tableHeaderHTML(Sheet, "cm")
    
    returnHTML = returnHTML + tableBodyHTML(Sheet)
    
    returnHTML = returnHTML + vbTab + "</div>" + vbNewLine
    returnHTML = returnHTML + "</span>" + vbNewLine
    returnHTML = returnHTML + "<!-- Inch Guide -->" + vbNewLine + vbNewLine
    
    inchGuideHTML = returnHTML
End Function

Function tableHeaderHTML(Sheet As Worksheet, unitIgnore As String, Optional indent As Integer = 2, Optional row As Integer = 4, Optional column As Integer = 1) As String
    Dim outputHTML As String
    Dim indentHTML As String
    
    indentHTML = ""
    For i = 1 To indent
        indentHTML = indentHTML + vbTab
    Next
    
    outputHTML = outputHTML + indentHTML + "<div>" + vbNewLine
    
    cellValue = Sheet.Cells(row, column)
    Do While cellValue <> ""
    
        'check if this head is in inches instead of cm
        unitsValue = Sheet.Cells(row - 1, column)
        If UCase(unitsValue) = UCase(unitIgnore) Then
            
            ReDim Preserve ignoreColumns(currentSize)
            ignoreColumns(currentSize) = column
            currentSize = currentSize + 1
                        
            GoTo NextIteration
        End If
        
        
        outputHTML = outputHTML + indentHTML + vbTab + "<div class='header'>" + cellValue + "</div>" + vbNewLine

NextIteration:
        column = column + 1
        cellValue = Sheet.Cells(row, column)
        colMax = colMax + 1
    Loop
    
    outputHTML = outputHTML + indentHTML + "</div>" + vbNewLine
    
    tableHeaderHTML = outputHTML
End Function

Function tableBodyHTML(Sheet As Worksheet, Optional indent As Integer = 2, Optional row As Integer = 5, Optional column As Integer = 1) As String
    Dim outputHTML As String
    Dim indentHTML As String
    Dim cellValue As String
    
    indentHTML = ""
    For i = 1 To indent
        indentHTML = indentHTML + vbTab
    Next
    
    cellValue = Sheet.Cells(row, column)
    
    outputHTML = ""
    
    Do While cellValue <> ""
    
        outputHTML = outputHTML + indentHTML + "<div>" + vbNewLine
        
        For col = column To colMax
        
            If IsInArray(col, ignoreColumns) Then
                GoTo NextCol
            End If
            
            cellValue = Sheet.Cells(row, col)
            
            If InStr(cellValue, "/") <> 0 Then
                cellValue = makeFractions(cellValue)
                cellValue = Replace(cellValue, " ", "")
            ElseIf VarType(cellValue) = vbString Then
                cellValue = Replace(cellValue, " ", "")
            Else
                cellValue = Round(cellValue, 1)
            End If
                
            
            outputHTML = outputHTML + indentHTML + vbTab + "<div>" + CStr(cellValue) + "</div>" + vbNewLine
            
NextCol:
        Next
        

        row = row + 1
        cellValue = Sheet.Cells(row, column)
        
        outputHTML = outputHTML + indentHTML + "</div>" + vbNewLine
    Loop
    
    tableBodyHTML = outputHTML
End Function

Function makeFractions(fullString As String) As String
    If Not makeIntoFractions Or InStr(fullString, "/") = 0 Then
        Exit Function
    End If
    
    Dim index As Integer
    Dim stringLeft, stringRight As String
    Dim numerator, denominator As String
    Dim fractionString As String
    
    index = InStr(fullString, "/")
    stringLeft = Left(fullString, index - 1)
    stringRight = Right(fullString, Len(fullString) - index)
    
    If InStr(stringLeft, " ") <> 0 Then
        stringLeft = StrReverse(stringLeft)
        index = InStr(stringLeft, " ")
        numerator = Left(stringLeft, index - 1)
        stringLeft = Right(stringLeft, Len(stringLeft) - index)
        stringLeft = StrReverse(stringLeft)
        
    Else
        numerator = stringLeft

    End If
    
    If InStr(stringRight, " ") <> 0 Then
        index = InStr(stringRight, " ")
        denominator = Left(stringRight, index - 1)
        stringRight = Right(stringRight, Len(stringRight) - index)
        
    Else
        denominator = stringRight
    End If
    
    fractionString = stringLeft + "<sup>" + numerator + "</sup>&frasl;<sub>" + denominator + "</sub>"
    
    'MsgBox fractionString
    
    makeFractions = fractionString + makeFractions(stringRight)
End Function

'Function to check if a value is in an array
'taken from online with error chekcing removed
Private Function IsInArray(valToBeFound As Variant, arr As Variant) As Boolean
    Dim element As Variant
    If IsEmpty(arr) Or IsNull(arr) Then
        IsInArray = False
        Exit Function
    End If
    
    For Each element In arr
        If element = valToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next element
    IsInArray = False
End Function

Private Function WriteToFile(filename As String, ByVal folder As String, contents As String)
    ' Get the current directory
    filePath = "D:\Users\Joe\OneDrive - armclaim\Documents\Online Content\Website\Size Guides" + "\" + folder + "\" + filename

    Open filePath For Output As #1
    
    Print #1, contents
    
    Close #1
End Function

