Attribute VB_Name = "Module1"
Sub Reorder_GL_Report()

 

 

    'Check if you are using the raw Ledger transction list.

    If Range("B1").Value <> "Ledger transaction list" Or Range("N7").Value <> "Debit" Then

    MsgBox "Invalid report format!" & vbCrLf & "Please use the Ledger transaction list - modified raw report from AX!", vbCritical

    Exit Sub

    End If

   

    'Check if the Revenue account/Consumption account in the report are for GP rec.

    Range("B4").FormulaR1C1 = "=COUNTIF(R[1]C:R[19]C,""1*"")+COUNTIF(R[1]C:R[19]C,""2*"")+COUNTIF(R[1]C:R[19]C,""6*"")+COUNTIF(R[1]C:R[19]C,""9*"")+COUNTIF(R[1]C:R[19]C,""<40000"")+COUNTIFS(R[1]C:R[19]C,"">59999"",R[1]C:R[19]C,""<70000"")+COUNTIF(R[1]C:R[19]C,"">89999"")"

 

    If Range("B4").Value > 0 Then

    Range("B4").Clear

    MsgBox "Invalid ledger account found!" & vbCrLf & "This Macro is for GP rec only!", vbExclamation

    Exit Sub

    End If

 

   

    

    Cells.WrapText = False

    Cells.MergeCells = False

 

    Range("N7:O7").Copy

    Range("N6").Select

    ActiveSheet.Paste

    Range("K6").Copy

    Range("N6:O6").PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

 

 

    Range("1:5,7:7").Delete Shift:=xlUp

    Range("A:A,D:D,L:L,M:M,P:P,Q:Q").Delete Shift:=xlToLeft

 

    Columns("U:V").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeft

    Range("U1").FormulaR1C1 = "Opportunity Id"

    Range("V1").FormulaR1C1 = "Charge code"

   

     'Place the column headers in the end result order as per the GP rec.

    Dim arrColOrder As Variant, ndx As Integer

    Dim Found As Range, counter As Integer

    arrColOrder = Array("Customer account", "Name", "Sales responsible", "Tax invoice", "Date", "Sales order", "Voucher", "Item number", "Description", "Division", "Debit", "Credit", "Ledger account", "Service Id", "Opportunity Id", "Charge code", "Purchase order")

    counter = 1

    Application.ScreenUpdating = False

    For ndx = LBound(arrColOrder) To UBound(arrColOrder)

        Set Found = Rows("1:1").Find(arrColOrder(ndx), LookIn:=xlValues, lookat:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)

        If Not Found Is Nothing Then

            If Found.Column <> counter Then

                Found.EntireColumn.Cut

                Columns(counter).Insert Shift:=xlToRight

                Application.CutCopyMode = False

            End If

            counter = counter + 1

        End If

    Next ndx

  

    Columns("D:D").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

    Columns("L:O").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

    Columns("R:S").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

 

   

    Range("D1").FormulaR1C1 = "Presales resource"

    Range("L1").FormulaR1C1 = "Manufacturer"

    Range("M1").FormulaR1C1 = "External item number"

    Range("N1").FormulaR1C1 = "Quantity"

    Range("O1").FormulaR1C1 = "Unit price"

    Range("R1").FormulaR1C1 = "Gross profit"

    Range("S1").FormulaR1C1 = "GP%"

   

    

    Columns("P:Q").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

    Range("P1").FormulaR1C1 = "Gross amount"

    Range("Q1").FormulaR1C1 = "Cost value"

   

    Dim Rrow As Long

    Rrow = Cells(Cells.Rows.Count, "H").End(xlUp).Row

   

    Range("P2:P" & Rrow).FormulaR1C1 = "=IF(OR(LEFT(RC[6],1)=""4"",LEFT(RC[6],1)=""7"")=TRUE,RC[3]-RC[2],0)"

    Range("P2:P" & Rrow).Value = Range("P2:P" & Rrow).Value

   

    Range("Q2:Q" & Rrow).FormulaR1C1 = "=IF(OR(LEFT(RC[5],1)=""5"",LEFT(RC[5],1)=""8"")=TRUE,RC[1]-RC[2],0)"

    Range("Q2:Q" & Rrow).Value = Range("Q2:Q" & Rrow).Value

   

    Range("T2:T" & Rrow).FormulaR1C1 = "=RC[-4]-RC[-3]"

    Columns("R:S").Delete Shift:=xlToLeft

   

    Columns("T:U").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

    Range("T1").FormulaR1C1 = "Revenue account"

    Range("T2:T" & Rrow).FormulaR1C1 = "=IF(OR(LEFT(RC[2],1)=""4"",LEFT(RC[2],1)=""7"")=TRUE,RC[2],"""")"

    Range("T2:T" & Rrow).Value = Range("T2:T" & Rrow).Value

    Range("U1").FormulaR1C1 = "Consumption account"

    Range("U2:U" & Rrow).FormulaR1C1 = "=IF(OR(LEFT(RC[1],1)=""5"",LEFT(RC[1],1)=""6"",LEFT(RC[1],1)=""8"")=TRUE,RC[1],"""")"

    Range("U2:U" & Rrow).Value = Range("U2:U" & Rrow).Value

    Columns("V:V").Delete Shift:=xlToLeft

   

 

    Range("N2:O" & Rrow).FormulaR1C1 = "0"

   

    Columns("T:T").TextToColumns Destination:=Range("T1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True

    Columns("U:U").TextToColumns Destination:=Range("U1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True

    Columns("V:V").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

    Columns("X:X").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove

    Range("V1").FormulaR1C1 = "Code"

    Range("W1").FormulaR1C1 = "EG number"

    Range("X1").FormulaR1C1 = "Tax invoice"

   

    Columns("AB:AJ").Delete

   

    Range("S2:S" & Rrow).FormulaR1C1 = "=IF(RC[-1]<=0,0,100)"

   Range("S2:S" & Rrow).Value = Range("S2:S" & Rrow).Value

 

   

    Columns("A:AA").ColumnWidth = 9.3

    Columns("J:J").ColumnWidth = 27.5

    Range("D:D,K:O,S:S,V:Z").ColumnWidth = 5

    Columns("A:AA").AutoFilter

   

    

    Dim rng As Range

    Set rng = Range("A1").CurrentRegion

    'Check if the transactions are from more than one division

    Dim MyRange As Range

    Dim myValue

    Dim allSame As Boolean

    'Set column to check

    Set MyRange = Range("K2:K" & Rrow)

    'Get first value from myRange

    myValue = MyRange(1, 1).Value

    allSame = (WorksheetFunction.CountA(MyRange) = WorksheetFunction.CountIf(MyRange, myValue))

    If allSame = False Then

    'Delete rows that are not from the division chosen

    Dim RemainStr As Variant

    'Keeping looping until getting a correct division

    Do

    'Retrieve an answer from the user

      RemainStr = Application.InputBox("Which Division to remain?", , "ADM/CRS/EBC/EPS/ESA", Type:=2)

    'Check if user click the Cancel button

    If TypeName(RemainStr) = "Boolean" Then GoTo Line1

    Loop Until UCase(RemainStr) = "ADM" Or UCase(RemainStr) = "CRS" Or UCase(RemainStr) = "EBC" Or UCase(RemainStr) = "EPS" Or UCase(RemainStr) = "ESA"

    rng.AutoFilter Field:=11, Criteria1:=Array("<>" & RemainStr), Operator:=xlAnd

    rng.Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete

    ActiveSheet.ShowAllData

    End If

   

    GoTo Line2

   

Line1:

    MsgBox "Please manually choose the division."

Line2:

    'Moving data from Column "Service Id" to Column "External item number"

    Range("W2:W" & Rrow).Copy

    Range("M2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

    'Clearing contents in EG number column that are not EG number

    rng.AutoFilter Field:=23, Criteria1:=Array("<>EG*"), Operator:=xlAnd

    Range("W2:W" & Rrow).SpecialCells(xlCellTypeVisible).ClearContents

    ActiveSheet.ShowAllData

   

    Dim Drow As Long

    Drow = Cells(Cells.Rows.Count, "H").End(xlUp).Row

    'Empty Revenue account

    If Application.WorksheetFunction.CountBlank(Range("T2:T" & Drow)) > 0 Then

    Range("T2:T" & Drow).SpecialCells(xlCellTypeBlanks).FormulaR1C1 = "=IF(RC[1]=51000,41100,IF(RC[1]=51150,41150,IF(RC[1]=51100,41200,IF(RC[1]=51121,41221,IF(RC[1]=51122,41222,IF(RC[1]=51120,41220,IF(RC[1]=51101,41201,IF(RC[1]=51102,41202,IF(RC[1]=51103,41203,IF(RC[1]=51104,41204,IF(RC[1]=51105,41205,IF(RC[1]=51111,41211,IF(RC[1]=51112,41212,IF(RC[1]=51200,41300,IF(RC[1]=51110,41210,IF(RC[1]=87100,71100,IF(RC[1]=51500,71200,"""")))))))))))))))))"

    Range("T2:T" & Drow).Value = Range("T2:T" & Drow).Value

    End If

    'Empty Consumption account

    If Application.WorksheetFunction.CountBlank(Range("U2:U" & Drow)) > 0 Then

    Range("U2:U" & Drow).SpecialCells(xlCellTypeBlanks).FormulaR1C1 = "=IF(RC[-1]=41100,51000,IF(RC[-1]=41150,51150,IF(RC[-1]=41200,51100,IF(RC[-1]=41221,51121,IF(RC[-1]=41222,51122,IF(RC[-1]=41220,51120,IF(RC[-1]=41201,51101,IF(RC[-1]=41202,51102,IF(RC[-1]=41203,51103,IF(RC[-1]=41204,51104,IF(RC[-1]=41205,51105,IF(RC[-1]=41211,51111,IF(RC[-1]=41212,51112,IF(RC[-1]=41300,51200,IF(RC[-1]=41210,51110,IF(RC[-1]=71100,87100,IF(RC[-1]=71200,51500,"""")))))))))))))))))"

    Range("U2:U" & Drow).Value = Range("U2:U" & Drow).Value

    End If

   

    Range("A1:AA" & Drow).Interior.ColorIndex = 0

   

    'Go back to cell A1

    Application.Goto Range("A1"), True

 

 

End Sub
Sub Reorder_GP_Report()


    If Range("A1").Value = "Textbox6" Then


        Columns("A:C").Delete
    
        Columns("A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeft
    
        Columns("C:D").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeft
    
        Columns("F:G").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeft
    
        Columns("J").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeft
    
        Columns("AM").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeft
    
        Rows("1:14").Insert


        Range("B15").Value = "Customer account": Range("E15").Value = "Name": Range("H15").Value = "Type"
    
        Range("I15").Value = "Sales responsible": Range("K15").Value = "Presales resource": Range("L15").Value = "Invoice id"
    
        Range("M15").Value = "Invoice date": Range("N15").Value = "Sales order number": Range("O15").Value = "Quotation"
    
        Range("P15").Value = "Opportunity Id": Range("Q15").Value = "Charge code": Range("R15").Value = "MSA Product code"
    
        Range("S15").Value = "MSA Product description": Range("T15").Value = "MSA Product Category": Range("U15").Value = "MSA Product Subcategory"
    
        Range("V15").Value = "MSA Unit Quantity": Range("W15").Value = "MSA Product Unit Price": Range("X15").Value = "Charge From Date"
    
        Range("Y15").Value = "Charge To Date": Range("Z15").Value = "Division": Range("AA15").Value = "Voucher"
    
        Range("AB15").Value = "Item number": Range("AC15").Value = "Text": Range("AD15").Value = "Manufacturer"
    
        Range("AE15").Value = "External item number": Range("AF15").Value = "Revenue account": Range("AG15").Value = "Consumption account"
    
        Range("AH15").Value = "Quantity": Range("AI15").Value = "Unit price": Range("AJ15").Value = "Gross amount"
    
        Range("AK15").Value = "Cost value": Range("AL15").Value = "Gross profit": Range("AN15").Value = "GP%"


        Range("B2").Value = "Gross Profit Report"

    Else

        If Range("B2").Value <> "Gross Profit Report" Or Range("AN15").Value <> "GP%" Then
    
            MsgBox "Invalid report format!" & vbCrLf & "Please use the Gross profit raw report from AX!", vbCritical
        
            Exit Sub
    
        End If
   
    End If

   
    Rows("1:14").Delete Shift:=xlUp

    Cells.WrapText = False

    Cells.MergeCells = False



    Columns("C:E").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeft

    Range("C1").FormulaR1C1 = "Code"

    Range("D1").FormulaR1C1 = "EG number"

    Range("E1").FormulaR1C1 = "Tax invoice"

   

    

    Dim arrColOrder As Variant, ndx As Integer

    Dim Found As Range, counter As Integer

   

    'Place the column headers in the end result order you want.

    arrColOrder = Array("Customer account", "Name", "Sales responsible", "Presales resource", "Invoice id", "Invoice date", "Sales order number", "Voucher", "Item number", "Text", "Division", "Manufacturer", "External item number", "Quantity", "Unit price", "Gross amount", "Cost value", "Gross profit", "GP%", "Revenue account", "Consumption account", "Code", "EG Number", "Tax Invoice", "Opportunity Id", "Charge code", "Quotation", "Type")

                      

    counter = 1

   

    Application.ScreenUpdating = False

   

    For ndx = LBound(arrColOrder) To UBound(arrColOrder)

   

        Set Found = Rows("1:1").Find(arrColOrder(ndx), LookIn:=xlValues, lookat:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)

       

        If Not Found Is Nothing Then

            If Found.Column <> counter Then

                Found.EntireColumn.Cut

                Columns(counter).Insert Shift:=xlToRight

                Application.CutCopyMode = False

            End If

            counter = counter + 1

        End If

       

    Next ndx

   

    

    Columns("T:T").TextToColumns Destination:=Range("T1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True

    Columns("U:U").TextToColumns Destination:=Range("U1"), DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo:=Array(1, 1), TrailingMinusNumbers:=True

   

    Columns("AC:AJ").Delete

   

    Columns("A:Z").ColumnWidth = 9.3

    Columns("J:J").ColumnWidth = 27.5

    Range("D:D,K:O,S:S,V:W").ColumnWidth = 5

    Columns("A:AA").AutoFilter

   

    Dim Rrow As Long

    Rrow = Cells(Cells.Rows.Count, "H").End(xlUp).Row

   

    If Range("K2").Value = "BCA" Or Range("K2").Value = "GNS" Or Range("K2").Value = "HWT" Then

    Range("A2:AA" & Rrow).Sort Key1:=Range("K2"), Order1:=xlAscending

    GoTo Line2

    End If

   

    Dim rng As Range

    Set rng = Range("A1").CurrentRegion

    'Check if the transactions are from more than one division

    Dim MyRange As Range

    Dim myValue

    Dim allSame As Boolean

    'Set column to check

    Set MyRange = Range("K2:K" & Rrow)

    'Get first value from myRange

    myValue = MyRange(1, 1).Value

    allSame = (WorksheetFunction.CountA(MyRange) = WorksheetFunction.CountIf(MyRange, myValue))

    If allSame = False Then

    Dim RemainStr As Variant

    'Keeping looping until getting a correct division

    Do

    'Retrieve an answer from the user

    RemainStr = Application.InputBox("Which Division to remain?", , "ADM/CRS/EPS/ESA/UAN", Type:=2)

    'Check if user selected cancel button

    If TypeName(RemainStr) = "Boolean" Then GoTo Line1

    Loop Until UCase(RemainStr) = "ADM" Or UCase(RemainStr) = "CRS" Or UCase(RemainStr) = "UAN" Or UCase(RemainStr) = "EPS" Or UCase(RemainStr) = "ESA"

    'Delete rows that are not from the division chosen

    rng.AutoFilter Field:=11, Criteria1:=Array("<>" & RemainStr), Operator:=xlAnd

    rng.Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete

    ActiveSheet.ShowAllData

    End If

   

    GoTo Line2

   

Line1:

    MsgBox "Please manually choose the division."

Line2:

    'Copy EG number from Column "External item number" to Column "EG number"

    Range("M2:M" & Rrow).Copy

    Range("W2").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

    Range("A1:Z1").AutoFilter Field:=23, Criteria1:=Array("<>EG*"), Operator:=xlAnd

    Range("W2:W" & Rrow).SpecialCells(xlCellTypeVisible).ClearContents

    ActiveSheet.ShowAllData

   

    Dim Drow As Long

    Drow = Cells(Cells.Rows.Count, "H").End(xlUp).Row

    'Filling empty Revenue account

    If Application.WorksheetFunction.CountBlank(Range("T2:T" & Drow)) > 0 Then

    Range("T2:T" & Drow).SpecialCells(xlCellTypeBlanks).FormulaR1C1 = "=IF(RC[1]=51000,41100,IF(RC[1]=51150,41150,IF(RC[1]=51100,41200,IF(RC[1]=51121,41221,IF(RC[1]=51122,41222,IF(RC[1]=51120,41220,IF(RC[1]=51101,41201,IF(RC[1]=51102,41202,IF(RC[1]=51103,41203,IF(RC[1]=51104,41204,IF(RC[1]=51105,41205,IF(RC[1]=51111,41211,IF(RC[1]=51112,41212,IF(RC[1]=51200,41300,IF(RC[1]=51110,41210,IF(RC[1]=87100,71100,IF(RC[1]=51500,71200,"""")))))))))))))))))"

    Range("T2:T" & Drow).Value = Range("T2:T" & Drow).Value

    End If

    'Filling empty Consumption account

    If Application.WorksheetFunction.CountBlank(Range("U2:U" & Drow)) > 0 Then

    Range("U2:U" & Drow).SpecialCells(xlCellTypeBlanks).FormulaR1C1 = "=IF(RC[-1]=41100,51000,IF(RC[-1]=41150,51150,IF(RC[-1]=41200,51100,IF(RC[-1]=41221,51121,IF(RC[-1]=41222,51122,IF(RC[-1]=41220,51120,IF(RC[-1]=41201,51101,IF(RC[-1]=41202,51102,IF(RC[-1]=41203,51103,IF(RC[-1]=41204,51104,IF(RC[-1]=41205,51105,IF(RC[-1]=41211,51111,IF(RC[-1]=41212,51112,IF(RC[-1]=41300,51200,IF(RC[-1]=41210,51110,IF(RC[-1]=71100,87100,IF(RC[-1]=71200,51500,"""")))))))))))))))))"

    Range("U2:U" & Drow).Value = Range("U2:U" & Drow).Value

    End If

    'Setting formula for Column "Gross profit"

    Range("R2:R" & Drow).FormulaR1C1 = "=RC[-2]-RC[-1]"

   

    Range("N2:S" & Drow).NumberFormat = "[$-10C09]#,##0.00;(#,##0.00)"

    Range("T2:U" & Drow).HorizontalAlignment = xlRight

   

    

    'For GP Rec - BCA

    If Range("K" & Drow).Value = "BCA" Or Range("K" & Drow).Value = "GNS" Or Range("K" & Drow).Value = "HWT" Or Range("K" & Drow).Value = "SALES" Then

     'Place the column headers in the end result order as per the GP rec.

    arrColOrder = Array("Customer account", "Name", "Type", "Sales responsible", "Presales resource", "Invoice id", "Invoice date", "Sales order number", "Division", "Voucher", "Item number", "Text", "Manufacturer", "External item number", "Revenue account", "Consumption account", "Quantity", "Unit price", "Gross amount", "Cost value", "Gross profit", "GP%", "Charge code", "Opportunity Id", "Quotation")

    counter = 1

    Application.ScreenUpdating = False

    For ndx = LBound(arrColOrder) To UBound(arrColOrder)

        Set Found = Rows("1:1").Find(arrColOrder(ndx), LookIn:=xlValues, lookat:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)

        If Not Found Is Nothing Then

            If Found.Column <> counter Then

                Found.EntireColumn.Cut

                Columns(counter).Insert Shift:=xlToRight

                Application.CutCopyMode = False

            End If

            counter = counter + 1

        End If

    Next ndx

    Range("F1").Value = "Invoice id": Range("G1").Value = "Invoice date"

    Range("H1").Value = "Sales order number": Range("L1").Value = "Text"

    Range("Z1").Value = "AAPT ACC#"

    Columns("AA:AB").Delete

    Range("C:C,X:X").ColumnWidth = 5

    Range("Y:Y").ColumnWidth = 9

   

    'Go back to cell A1

    Application.Goto Range("A1"), True

    Exit Sub

    End If

    Columns("AB").Delete

 

    'Go back to cell A1

    Application.Goto Range("A1"), True

 

End Sub

Sub Inventory_Ageing()

   

    'Check which step in the Macro would like to be executed

    Dim Answer As Integer

    Answer = MsgBox("Would like to go through the whole code execution process?" & vbCrLf & vbCrLf & "Yes - Go through the whole Macro execution process" & vbCrLf & "No - Skip to the PivotTable generation step" & vbCrLf & "Cancel - Stop code execution", vbQuestion + vbYesNoCancel + vbDefaultButton2, "Inventory Ageing Reporting")

    If Answer = vbNo Then

    GoTo Line2

    Else

    If Answer = vbCancel Then

    Exit Sub

    End If

    End If

 

    

    Application.DisplayAlerts = False

    On Error Resume Next

    Worksheets("Working").Delete

    Worksheets("Available Physical").Delete

    Worksheets("PivotTable").Delete

    Worksheets("EG Inventory").Delete

    Worksheets("247 Inventory").Delete

    Worksheets("POESA20633").Delete

    On Error GoTo 0

    Application.DisplayAlerts = True

   

    

    ''''Add a "Working" sheet

    Worksheets.Add After:=Worksheets("Period Lookup")

    ActiveSheet.Name = "Working"

    With Worksheets("Working").Tab

        .ThemeColor = xlThemeColorAccent1

        .TintAndShade = 0

    End With

    Sheets("RAW").Cells.Copy

    Sheets("Working").Cells.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

       

    'To make sure the columns are included as per prevous Ageing report's format

    Dim arrColOrder As Variant, ndx As Integer

    Dim Found As Range, counter As Integer

    arrColOrder = Array("Item number", "Product name", "Search name", "Warehouse", "Batch number", "Location", "Serial number", "Physical date", "Cost price", "Physical inventory", "Physical reserved", "Available physical", "Ordered in total", "On order", "Ordered reserved", "Total available", "Uses warehouse management processes")

    counter = 1

    Application.ScreenUpdating = False

    For ndx = LBound(arrColOrder) To UBound(arrColOrder)

        Set Found = Rows("1:1").Find(arrColOrder(ndx), LookIn:=xlValues, lookat:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)

        If Not Found Is Nothing Then

            If Found.Column <> counter Then

                Found.EntireColumn.Cut

                Columns(counter).Insert Shift:=xlToRight

                Application.CutCopyMode = False

            End If

            counter = counter + 1

        End If

    Next ndx

   

    If IsEmpty(Range("R1").Value) = False Or Range("Q1").Value <> "Uses warehouse management processes" Then

    MsgBox "Please make sure you have included only the required columns as per previous Ageing report's format"

    Exit Sub

    End If

   

    Columns("H:H").NumberFormat = "m/d/yyyy"

   

    ''''Add a "Available Physical" sheet

    Worksheets.Add After:=Worksheets("Working")

    ActiveSheet.Name = "Available Physical"

    With Worksheets("Available Physical").Tab

        .ThemeColor = xlThemeColorAccent1

        .TintAndShade = 0

    End With

    Sheets("Working").Cells.Copy

    Sheets("Available Physical").Cells.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

 

    Columns("A:A").NumberFormat = "@"

    Range("R1").Value = "Available Stock Value": Range("S1").FormulaR1C1 = "=TODAY()": Range("S1").Value = Range("S1").Value

    Range("T1").Value = "Period": Range("U1").Value = "Manufacturer"

   

 

    Dim Wrow As Long

    Wrow = Sheets("Available Physical").Cells(Rows.Count, "A").End(xlUp).Row

    Sheets("Available Physical").Select

    Range("H2:H" & Wrow).NumberFormat = "m/d/yyyy"

    Range("R2:R" & Wrow).FormulaR1C1 = "=RC[-9]*RC[-6]"

    Range("S2:S" & Wrow).FormulaR1C1 = "=R1C19-RC[-11]"

    Range("T2:T" & Wrow).FormulaR1C1 = "=VLOOKUP(RC[-1],'Period Lookup'!C[-19]:C[-18],2,1)"

    Range("U2:U" & Wrow).FormulaR1C1 = "=VLOOKUP(RC[-20],'Item Lookup'!C[-20]:C[-18],3,0)"

    Columns("A:U").AutoFilter

   

 

    If Application.WorksheetFunction.CountIfs(Range("L1:L" & Wrow), "", Range("R1:R" & Wrow), 0) = 0 Then

    GoTo Line1

    Else

    Range("A1:U" & Wrow).AutoFilter Field:=12, Criteria1:="="

    Range("A1:U" & Wrow).AutoFilter Field:=18, Criteria1:=0

    Sheets("Available Physical").Range("A2:U" & Wrow).SpecialCells(xlCellTypeVisible).EntireRow.Delete

    ActiveSheet.ShowAllData

    End If

 

 

Line1:

 

    Dim Wrow2 As Long

    Wrow2 = Sheets("Available Physical").Cells(Rows.Count, "A").End(xlUp).Row

    'Check if Manufacturer column contains any #N/A

    Dim rngFoundCell As Range

    Set rngFoundCell = Sheets("Available Physical").Range("U2:U" & Wrow2).Find(what:="#N/A")

    If rngFoundCell Is Nothing Then 'if no Manufacturer is missing

        GoTo Line2

    Else

 

        Dim ILrow As Long   'if Manufacturer is missing

        ILrow = Sheets("Item Lookup").Cells(Rows.Count, "A").End(xlUp).Row

        Worksheets("Available Physical").Range("A1:U" & Wrow).AutoFilter Field:=21, Criteria1:="#N/A"

        Sheets("Available Physical").Range("A2:A" & Wrow).SpecialCells(xlCellTypeVisible).Copy Destination:=Sheets("Item Lookup").Range("A" & ILrow + 1)

        Sheets("Item Lookup").Select

        Range("A" & ILrow + 1).Select

        Range(Selection, Selection.End(xlDown)).RemoveDuplicates Columns:=1, Header:=xlNo

 

        Application.Goto Reference:=Worksheets("Item Lookup").Range("A" & ILrow + 1), Scroll:=True: ActiveWindow.SmallScroll Down:=-9

       

        MsgBox "Missing Manufacturer found. Please update the Item Lookup sheet.", vbCritical

 

        Exit Sub

    End If

 

 

Line2:

 

    Application.DisplayAlerts = False

    On Error Resume Next

    Worksheets("PivotTable").Delete

    Worksheets("EG Inventory").Delete

    Worksheets("247 Inventory").Delete

    Worksheets("Inventory").Delete

    On Error GoTo 0

    Application.DisplayAlerts = True

   

    MsgBox "Checking Item Lookup."

   

    Dim ILrow2 As Long

    ILrow2 = Sheets("Item Lookup").Cells(Rows.Count, "A").End(xlUp).Row

    If Application.WorksheetFunction.CountA(Worksheets("Item Lookup").Range("B:B")) <> ILrow2 And Application.WorksheetFunction.CountA(Worksheets("Item Lookup").Range("C:C")) <> ILrow2 Then

    MsgBox "Missing Item name and Manufacturer info. Plese fill in the Item Lookup sheet.", vbCritical

    Exit Sub

    ElseIf Application.WorksheetFunction.CountA(Worksheets("Item Lookup").Range("B:B")) <> ILrow2 Then

    MsgBox "Missing Item name info. Plese fill in the Item Lookup sheet.", vbCritical

    Exit Sub

    ElseIf Application.WorksheetFunction.CountA(Worksheets("Item Lookup").Range("C:C")) <> ILrow2 Then

    MsgBox "Missing Manufacturer info. Plese fill in the Item Lookup sheet.", vbCritical

    Exit Sub

    End If

   

    'Check if data is filtered

    Dim rngFilter As Range

    Dim r As Long, f As Long

    Set rngFilter = Worksheets("Available Physical").AutoFilter.Range

    r = rngFilter.Rows.Count

    f = rngFilter.SpecialCells(xlCellTypeVisible).Count

    If r > f Then 'If filtered, show data

    Worksheets("Available Physical").ShowAllData

    End If

   

      

    'Remove POESA206333

    Sheets("Available Physical").Select

    Dim Wrow4 As Long

    Wrow4 = Sheets("Available Physical").Cells(Rows.Count, "A").End(xlUp).Row

    If Application.WorksheetFunction.CountIf(Range("E1:E" & Wrow4), "POESA206333*") = 0 Then

    Else

    '''''''''''''''''

    Application.DisplayAlerts = False

    On Error Resume Next

 

    Worksheets("POESA20633").Delete

    On Error GoTo 0

    Application.DisplayAlerts = True

    ''''''''''''''

    Worksheets.Add After:=Worksheets("Working")

    ActiveSheet.Name = "POESA20633"

    With Worksheets("POESA20633").Tab

        .ThemeColor = xlThemeColorAccent1

        .TintAndShade = 0

    End With

    Sheets("Available Physical").Cells.Copy

    Sheets("POESA20633").Cells.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

   

    Range("A1:U" & Wrow4).AutoFilter Field:=5, Criteria1:="<>POESA206333*"

    Range("A2:U" & Wrow4).SpecialCells(xlCellTypeVisible).EntireRow.Delete

    ActiveSheet.ShowAllData

    Application.Goto Range("A1"), True

   

    

    Sheets("Available Physical").Select

    Range("A1:U" & Wrow4).AutoFilter Field:=5, Criteria1:="POESA206333*"

    Range("A2:U" & Wrow4).SpecialCells(xlCellTypeVisible).EntireRow.Delete

    ActiveSheet.ShowAllData

    End If

 

   

    

    

    

    

    

    

    'Add dummies to the Available Physical sheet

    Worksheets("Available Physical").Select

    Rows("2:6").Insert Shift:=x5Down

    Range("A2") = "Dummy1": Range("T2") = "P1 (0 - 90 Days)"

    Range("A3") = "Dummy2": Range("T3") = "P2 (91 - 180 Days)"

    Range("A4") = "Dummy3": Range("T4") = "P3 (181 - 270 Days)"

    Range("A5") = "Dummy4": Range("T5") = "P4 (271 - 360 Days)"

    Range("A6") = "Dummy5": Range("T6") = "P5 (Over 360 Days)"

   

    'Marking negative physical inventory

    Dim Wrow3 As Long

    Wrow3 = Sheets("Available Physical").Cells(Rows.Count, "A").End(xlUp).Row

    If Application.WorksheetFunction.CountIf(Range("J1:J" & Wrow3), "<0") = 0 Then

    Else

    Range("A1:U" & Wrow3).AutoFilter Field:=10, Criteria1:="<0"

    Dim rngN As Range

    Set rngN = Range("A2:A" & Wrow3).SpecialCells(xlCellTypeVisible)

    For Each cell In rngN

        Range("U" & cell.Row).Value = Range("U" & cell.Row).Value

        Range("A" & cell.Row).Value = Range("A" & cell.Row).Value & " negative"

    Next cell

 

    ActiveSheet.ShowAllData

    End If

   

    

    

    

    

    ''''Create PivotTable

    Dim PSheet As Worksheet

    Dim DSheet As Worksheet

    Dim PCache As PivotCache

    Dim PTable As PivotTable

    Dim PRange As Range

    Dim LastRow As Long

    Dim LastCol As Long

    'Insert a New Blank Worksheet

    On Error Resume Next

    Application.DisplayAlerts = False

    Worksheets("PivotTable").Delete

    Sheets.Add After:=Sheets("Available Physical")

    ActiveSheet.Name = "PivotTable"

    With Worksheets("PivotTable").Tab

        .ThemeColor = xlThemeColorAccent1

        .TintAndShade = 0

    End With

    Application.DisplayAlerts = True

    Set PSheet = Worksheets("PivotTable")

    Set DSheet = Worksheets("Available Physical")

    'Define Data Range

    LastRow = DSheet.Cells(Rows.Count, 1).End(xlUp).Row

    LastCol = DSheet.Cells(1, Columns.Count).End(xlToLeft).Column

    Set PRange = DSheet.Cells(1, 1).Resize(LastRow, LastCol)

    'Define Pivot Cache

    Set PCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=PRange).CreatePivotTable(TableDestination:=PSheet.Cells(3, 1), TableName:="PivotTable1")

    'Insert Row Fields

    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Period")

        .Orientation = xlColumnField

        .Position = 1

    End With

    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Item number")

        .Orientation = xlRowField

        .Position = 1

    End With

    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Product name")

        .Orientation = xlRowField

        .Position = 2

    End With

    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Manufacturer")

        .Orientation = xlRowField

        .Position = 3

    End With

    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables("PivotTable1").PivotFields("Available physical"), "Sum of Available physical", xlSum

    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables("PivotTable1").PivotFields("Available Stock Value"), "Sum of Available Stock Value", xlSum

    Range("A7").Select

    With ActiveSheet.PivotTables("PivotTable1")

        .InGridDropZones = True

        .RowAxisLayout xlTabularRow

    End With

    Range("A6").Select

   ActiveSheet.PivotTables("PivotTable1").PivotFields("Item number").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)

    Range("B6").Select

    ActiveSheet.PivotTables("PivotTable1").PivotFields("Product name").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)

       

    Columns("A:P").ColumnWidth = 15

 

    ''''Add an "Inventory" sheet

    Worksheets.Add After:=Worksheets("PivotTable")

    ActiveSheet.Name = "Inventory"

    Worksheets("Inventory").Tab.ThemeColor = 49407

 

    Range("E:E,G:G,K:K,M:M,O:O").ColumnWidth = 13.5

    Sheets("PivotTable").Range("A3:O5").Copy Destination:=Sheets("Inventory").Range("A3")

    Sheets("Inventory").Select

    Range("N4:O4").Copy Range("N5:O5")

    Range("N4").Value = "Total Period": Range("O4").ClearContents

    Columns("A:A").NumberFormat = "@"

    Range("A3:O5").Font.Name = "Arial": Range("A3:O5").Font.Size = 9

    Range("M5,O4").HorizontalAlignment = xlFill

    Range("A5:O5").Style = "Accent1": Range("A5:O5").Font.Bold = True

    Range("A3:O3").Interior.Pattern = xlNone

    Range("A4:C4").HorizontalAlignment = xlCenter: Range("A4:C4").Merge

    Range("D4:E4").HorizontalAlignment = xlCenter: Range("D4:E4").Merge

    Range("F3:G3,F4:G4").HorizontalAlignment = xlCenter: Range("F3:G3,F4:G4").Merge

    Range("H3:I3,H4:I4").HorizontalAlignment = xlCenter: Range("H3:I3,H4:I4").Merge

    Range("J3:K3,J4:K4").HorizontalAlignment = xlCenter: Range("J3:K3,J4:K4").Merge

    Range("L3:M3,L4:M4").HorizontalAlignment = xlCenter: Range("L3:M3,L4:M4").Merge

    Range("N3:O3,N4:O4").HorizontalAlignment = xlCenter: Range("N3:O3,N4:O4").Merge

    Range("F3:O3,A4:O4").Borders.LineStyle = xlContinuous

    Range("D3:E3").BorderAround ColorIndex:=xlAutomatic, Weight:=xlThin

   

    

    Dim Prow As Long ' Count the number of row from Cell A1 to Grandtotal

    Prow = Sheets("PivotTable").Cells(Rows.Count, "A").End(xlUp).Row

    Sheets("PivotTable").Range("A6:O" & Prow - 1).Copy

    Sheets("Inventory").Range("A6").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

    Range("A6:O" & Prow).Font.Name = "Calibri": Range("A6:O" & Prow).Font.Size = 9

    Range("D:D,F:F,H:H,J:J,L:L,N:N").NumberFormat = "0_ ;[Red]-0 "

    Range("E:E,G:G,I:I,K:K,M:M,O:O").NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"

    'Create the Grand Total line

    Range("A" & Prow + 1).Value = "Grand Total"

   Range("D" & Prow + 1 & ":O" & Prow + 1).FormulaR1C1 = "=SUM(R[-" & Prow - 5 & "]C:R[-1]C)"

    Range("A" & Prow + 1 & ":O" & Prow + 1).Font.Size = 9

    Range("A" & Prow + 1 & ":O" & Prow + 1).Font.Bold = True

    With Range("A" & Prow + 1 & ":O" & Prow + 1).Borders(xlEdgeTop)

        .LineStyle = xlContinuous

        .ColorIndex = 0

        .TintAndShade = 0

        .Weight = xlThin

    End With

    With Range("A" & Prow + 1 & ":O" & Prow + 1).Borders(xlEdgeBottom)

        .LineStyle = xlDouble

        .ColorIndex = 0

        .TintAndShade = 0

        .Weight = xlThick

    End With

   

    

    Dim Nap As Long 'Count the number of item with Available physical less than or equal to 0

    'Check if there is any item with Available physical less than or equal to 0

    If Application.WorksheetFunction.CountIf(Range("N6:N" & Prow), "<=0") = 0 Then

    Else

    'Sort the Total Sum of Available physical in ascending order

    Range("A6:O" & Prow).Sort Key1:=Range("N6"), Order1:=xlAscending

    Range("A5:O5").AutoFilter

    Range("A5:O5").AutoFilter Field:=14, Criteria1:="<=0"

    Set rng = Range("A6:A" & Prow).SpecialCells(xlCellTypeVisible)

    For Each cell In rng

        Range("A" & cell.Row).Value = WorksheetFunction.Substitute(Range("A" & cell.Row), " negative", "")

    Next cell

    Nap = Range("A6:A" & Prow).SpecialCells(xlCellTypeVisible).Count - 1

    Rows("6:" & Prow - 1).SpecialCells(xlCellTypeVisible).Copy Destination:=Range("A" & Prow + 7)

    Rows("6:" & Prow - 1).SpecialCells(xlCellTypeVisible).EntireRow.Delete

    ActiveSheet.ShowAllData

    End If

   

    'Sort the Total Sum of Available Stock Value in descending order

    Range("A6:O" & Prow - Nap - 1).Sort Key1:=Range("O6"), Order1:=xlDescending

   

    'Delete the dummies

    Range("A5:O" & Prow - Nap - 1).AutoFilter Field:=1, Criteria1:="Dummy*"

    Range("A6:O" & Prow - Nap - 1).SpecialCells(xlCellTypeVisible).EntireRow.Delete

    ActiveSheet.ShowAllData

    Worksheets("Available Physical").Rows("2:6").Delete Shift:=xlUp

    Worksheets("PivotTable").PivotTables("PivotTable1").PivotCache.Refresh

   

    

    'Delete the negative markings on physical inventory

    Sheets("Available Physical").Select

    Wrow3 = Sheets("Available Physical").Cells(Rows.Count, "A").End(xlUp).Row

    If Application.WorksheetFunction.CountIf(Range("J1:J" & Wrow3), "<0") = 0 Then

    Else

    Range("A1:U" & Wrow3).AutoFilter Field:=10, Criteria1:="<0"

    Set rng = Range("A2:A" & Wrow3).SpecialCells(xlCellTypeVisible)

    For Each cell In rng

        Range("A" & cell.Row).Value = WorksheetFunction.Substitute(Range("A" & cell.Row), " negative", "")

        Range("U" & cell.Row).FormulaR1C1 = "=VLOOKUP(RC[-20],'Item Lookup'!C[-20]:C[-18],3,0)"

    Next cell

 

    ActiveSheet.ShowAllData

    End If

   
    

    Sheets("Inventory").Select

    'Rearrange worksheet order

    Application.DisplayAlerts = False

    Application.ScreenUpdating = False

    WSOrder = Array("RAW", "Item Lookup", "Period Lookup", "Working", "POESA20633", "Available Physical", "PivotTable", "Inventory", "Note")

    On Error Resume Next

    For wso = UBound(WSOrder) To LBound(WSOrder) Step -1

    Worksheets(WSOrder(wso)).Move Before:=Worksheets(1)

    Next wso

    On Error GoTo 0

    Application.DisplayAlerts = True

    Application.ScreenUpdating = True

   

    Worksheets("Available Physical").Select

    Range("R1:R1").EntireColumn.Select

    Worksheets("Inventory").Select

    With Worksheets("Inventory").Tab

        .Color = 49407

        .TintAndShade = 0

    End With

    If Worksheets("Item Lookup").Range("E1") = "EG Inventory" Or Worksheets("Item Lookup").Range("E1") = "247 Inventory" Then

    Worksheets("Inventory").Name = Worksheets("Item Lookup").Range("E1")

    End If

   

    

    Range("A1").NumberFormat = "General"

    Range("A1").FormulaR1C1 = "=TEXT('Available Physical'!RC[18],""yymmdd"")&"" ""&""Inventory Ageing Report""&"" - ""&SUBSTITUTE('Item Lookup'!RC[4],"" Inventory"","""")"

    Range("A1").Value = Range("A1").Value

    Range("A1").Font.Bold = True

   

    'Go back to Grand Total

    Application.Goto Range("O" & Prow - Nap - 4), True

    Range("O" & Prow - Nap - 4 & ":O" & Prow + 1).Select

    ActiveWindow.ScrollRow = Prow - Nap - 7

    ActiveWindow.ScrollColumn = 1

   

    

    'check if the total balance on the report match the total balance on the raw data

    Dim InvBal As Double

    Dim AvaPhyBal As Double

    Dim wsSheet As Worksheet

    For Each wsSheet In Worksheets

        If wsSheet.Name = "EG Inventory" Then

        InvBal = Round(Application.Sum(Worksheets("EG Inventory").Range("O" & Prow - Nap - 4 & ":O" & Prow + 1)), 2)

        Else

        If wsSheet.Name = "247 Inventory" Then

        InvBal = Round(Application.Sum(Worksheets("247 Inventory").Range("O" & Prow - Nap - 4 & ":O" & Prow + 1)), 2)

        Else

        If wsSheet.Name = "Inventory" Then

        InvBal = Round(Application.Sum(Worksheets("Inventory").Range("O" & Prow - Nap - 4 & ":O" & Prow + 1)), 2)

        End If

        End If

        End If

    Next wsSheet

 

    AvaPhyBal = Round(Application.WorksheetFunction.Sum(Worksheets("Available Physical").Columns("R:R")), 2)

   

    If InvBal = AvaPhyBal Then

    MsgBox "DATA SET MATCHED" & vbCrLf & vbCrLf & "Inventory Ageing balance: " & Format(InvBal, "$#,##0.00") & vbCrLf & "Available Physical balance: " & Format(AvaPhyBal, "$#,##0.00"), , "Inventory Ageing Report"

       

        

        'Save as
    
        Dim wbstring As String
    
        For Each wsSheet In Worksheets

        If wsSheet.Name = "EG Inventory" Then

            wbstring = Worksheets("EG Inventory").Range("A1")

        Else

            If wsSheet.Name = "247 Inventory" Then

                wbstring = Worksheets("247 Inventory").Range("A1")

            Else

                If wsSheet.Name = "Inventory" Then

                    wbstring = Worksheets("Inventory").Range("A1")

                End If

            End If

        End If

        Next wsSheet

   

    

        If InStr(ActiveWorkbook.FullName, Left(wbstring, 6)) Then

   
        
            Dim Msg, Style, Title, Help, Ctxt, Response, MyString
        
            Msg = "Want to save your changes to '" & ActiveWorkbook.Name & "'?"    ' Define message.
        
            Style = vbYesNo + vbExclamation + vbDefaultButton2    ' Define buttons.
        
            Title = "Microsoft Excel"    ' Define title.
        
            Help = ""    ' Define Help file.
        
            Ctxt = 1000    ' Define topic context.

             ' Display message.

            Response = MsgBox(Msg, Style, Title, Help, Ctxt)
        
            If Response = vbYes Then    ' User chose Yes.
        
                ActiveWorkbook.Save   ' Perform some action.
        
                MsgBox "Document has been saved."
        
            Else    ' User chose No.
        
                MsgBox "Document not saved."   ' Perform some action.
        
            End If

   

        Else

   

            Msg = "Want to save your document as '" & wbstring & "'" & " under " & ActiveWorkbook.Path & "?"    ' Define message.
        
            Style = vbYesNo + vbExclamation + vbDefaultButton2    ' Define buttons.
        
            Title = "Microsoft Excel"    ' Define title.
        
            Help = ""    ' Define Help file.
        
            Ctxt = 1000    ' Define topic context.

            ' Display message.

            Response = MsgBox(Msg, Style, Title, Help, Ctxt)

            If Response = vbYes Then    ' User chose Yes.
        
                ActiveWorkbook.SaveAs FileName:=Application.ActiveWorkbook.Path & "\" & wbstring   ' Perform some action.
        
                MsgBox "Document has been saved."
        
            Else    ' User chose No.
        
                MsgBox "Document not saved."   ' Perform some action.
        
            End If

   

        End If

       

    Else

        MsgBox "DATA SET NOT MATCHED" & vbCrLf & vbCrLf & "Inventory Ageing balance: " & Format(InvBal, "$#,##0.00") & vbCrLf & "Available Physical balance: " & Format(AvaPhyBal, "$#,##0.00"), vbCritical, "Inventory Ageing Report"

    End If

   

    

    

    

End Sub
Sub Customer_Ageing()

 

    'Make sure the Workbook has only "Customer Ageing", "Ageing Data" and "Note" worksheets.

    Dim N As Long

    N = Worksheets.Count

    If N <> 3 Then

    MsgBox "Please include only Customer Ageing, Ageing Data and Note worksheet!"

    Exit Sub

    End If

    For Each wsSheet In Worksheets

        If wsSheet.Name = "Customer Ageing" Or wsSheet.Name = "Ageing Data" Or wsSheet.Name = "Note" Then

        Else: MsgBox "Please include only Customer Ageing, Ageing Data and Note worksheet!"

            Exit Sub

        End If

    Next wsSheet

   

    'Make sure the format of the raw data is correct.

    Dim Agcol As Long

    Agcol = Worksheets("Ageing Data").Cells(1, Worksheets("Ageing Data").Columns.Count).End(xlToLeft).Column

    If Agcol = 12 Then

    Else: MsgBox "Please make sure Worksheet Ageing Data has a table with exact 12 columns!"

    Exit Sub: End If

   

 

    Worksheets("Customer Ageing").Select

    Set A247 = Range("A:A").Find(what:="24/7 Distribution Pty Ltd", LookIn:=xlValues, lookat:=xlWhole)

    Set E247 = Range("E:E").Find(what:="24/7 Distribution Pty Ltd", LookIn:=xlValues, lookat:=xlWhole)

    Set A247N = Range("A:A").Find(what:="24/7 Distribution NZ Pty Ltd", LookIn:=xlValues, lookat:=xlWhole)

    Set E247N = Range("E:E").Find(what:="24/7 Distribution NZ Pty Ltd", LookIn:=xlValues, lookat:=xlWhole)

    Set ABCA = Range("A:A").Find(what:="CONNECTANZ PTY LTD", LookIn:=xlValues, lookat:=xlWhole)

    Set EBCA = Range("E:E").Find(what:="CONNECTANZ PTY LTD", LookIn:=xlValues, lookat:=xlWhole)

    Set ACAL = Range("A:A").Find(what:="Connect ANZ Ltd T/A SNAP Business Connect", LookIn:=xlValues, lookat:=xlWhole)

    Set ECAL = Range("E:E").Find(what:="Connect ANZ Ltd T/A SNAP Business Connect", LookIn:=xlValues, lookat:=xlWhole)

    Set AEG = Range("A:A").Find(what:="Ethan Group Pty Ltd", LookIn:=xlValues, lookat:=xlWhole)

    Set EEG = Range("E:E").Find(what:="Ethan Group Pty Ltd", LookIn:=xlValues, lookat:=xlWhole)

    Set AETL = Range("A:A").Find(what:="ETHAN TALENT PTY LTD", LookIn:=xlValues, lookat:=xlWhole)

    Set EETL = Range("E:E").Find(what:="ETHAN TALENT PTY LTD", LookIn:=xlValues, lookat:=xlWhole)

    Set ASRM = Range("A:A").Find(what:="SimpleRoam Pty Ltd", LookIn:=xlValues, lookat:=xlWhole)

    Set ESRM = Range("E:E").Find(what:="SimpleRoam Pty Ltd", LookIn:=xlValues, lookat:=xlWhole)

    Set AE4I = Range("A:A").Find(what:="E4I INDIGENOUS PTY LTD", LookIn:=xlValues, lookat:=xlWhole)

    Set EE4I = Range("E:E").Find(what:="E4I INDIGENOUS PTY LTD", LookIn:=xlValues, lookat:=xlWhole)

   

 

   

    'Remove the old data

    If Split(E247.Address, "$")(2) - Split(A247.Address, "$")(2) > 6 Then

    Rows(Split(A247.Address, "$")(2) + 5 & ":" & Split(E247.Address, "$")(2) - 2).Delete

    Else: End If

    If Split(E247N.Address, "$")(2) - Split(A247N.Address, "$")(2) > 6 Then

    Rows(Split(A247N.Address, "$")(2) + 5 & ":" & Split(E247N.Address, "$")(2) - 2).Delete

    Else: End If

    If Split(EBCA.Address, "$")(2) - Split(ABCA.Address, "$")(2) > 6 Then

   Rows(Split(ABCA.Address, "$")(2) + 5 & ":" & Split(EBCA.Address, "$")(2) - 2).Delete

    Else: End If

    If Split(ECAL.Address, "$")(2) - Split(ACAL.Address, "$")(2) > 6 Then

    Rows(Split(ACAL.Address, "$")(2) + 5 & ":" & Split(ECAL.Address, "$")(2) - 2).Delete

    Else: End If

    If Split(EEG.Address, "$")(2) - Split(AEG.Address, "$")(2) > 6 Then

    Rows(Split(AEG.Address, "$")(2) + 5 & ":" & Split(EEG.Address, "$")(2) - 2).Delete

    Else: End If

    If Split(EETL.Address, "$")(2) - Split(AETL.Address, "$")(2) > 6 Then

    Rows(Split(AETL.Address, "$")(2) + 5 & ":" & Split(EETL.Address, "$")(2) - 2).Delete

    Else: End If

    If Split(ESRM.Address, "$")(2) - Split(ASRM.Address, "$")(2) > 6 Then

    Rows(Split(ASRM.Address, "$")(2) + 5 & ":" & Split(ESRM.Address, "$")(2) - 2).Delete

    Else: End If

    If Split(EE4I.Address, "$")(2) - Split(AE4I.Address, "$")(2) > 6 Then

    Rows(Split(AE4I.Address, "$")(2) + 5 & ":" & Split(EE4I.Address, "$")(2) - 2).Delete

    End If

 

    'Copy the data to each entity

    Worksheets("Ageing Data").Select

    If Worksheets("Ageing Data").ListObjects.Count > 0 Then

    Worksheets("Ageing Data").ListObjects(1).TableStyle = ""

    ActiveSheet.ListObjects(1).Unlist

    End If

    Columns("A:L").HorizontalAlignment = xlLeft

    Columns("E:K").NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"

   

    ''''''''''''''''''''

    Rows("2:2").Insert Shift:=xlDown

    Range("A2").FormulaR1C1 = "Dummy247"

    Range("L2").FormulaR1C1 = "247"

    '''''''''''''''''''''''

   

    Dim Drow As Long

    Drow = Cells(Cells.Rows.Count, "A").End(xlUp).Row

   

    With ActiveSheet.Sort

     .SortFields.Add Key:=Range("L1"), Order:=xlAscending

     .SetRange Range("A1:L" & Drow)

     .Header = xlYes

     .Apply

    End With

   

    Range("A1:L1").AutoFilter Field:=12, Criteria1:=Array("247"), Operator:=xlAnd

        Range("A2:L" & Drow).SpecialCells(xlCellTypeVisible).Copy

        Worksheets("Customer Ageing").Range("A" & Split(A247.Address, "$")(2) + 5).Insert Shift:=xlDown

    Range("A1:L1").AutoFilter Field:=12, Criteria1:=Array("247N"), Operator:=xlAnd

        Range("A2:L" & Drow).SpecialCells(xlCellTypeVisible).Copy

        Worksheets("Customer Ageing").Range("A" & Split(A247N.Address, "$")(2) + 5).Insert Shift:=xlDown

    Range("A1:L1").AutoFilter Field:=12, Criteria1:=Array("BCA"), Operator:=xlAnd

        Range("A2:L" & Drow).SpecialCells(xlCellTypeVisible).Copy

        Worksheets("Customer Ageing").Range("A" & Split(ABCA.Address, "$")(2) + 5).Insert Shift:=xlDown

    Range("A1:L1").AutoFilter Field:=12, Criteria1:=Array("CAL"), Operator:=xlAnd

        Range("A2:L" & Drow).SpecialCells(xlCellTypeVisible).Copy

        Worksheets("Customer Ageing").Range("A" & Split(ACAL.Address, "$")(2) + 5).Insert Shift:=xlDown

    Range("A1:L1").AutoFilter Field:=12, Criteria1:=Array("EG"), Operator:=xlAnd

        Range("A2:L" & Drow).SpecialCells(xlCellTypeVisible).Copy

        Worksheets("Customer Ageing").Range("A" & Split(AEG.Address, "$")(2) + 5).Insert Shift:=xlDown

    Range("A1:L1").AutoFilter Field:=12, Criteria1:=Array("ETL"), Operator:=xlAnd

        Range("A2:L" & Drow).SpecialCells(xlCellTypeVisible).Copy

        Worksheets("Customer Ageing").Range("A" & Split(AETL.Address, "$")(2) + 5).Insert Shift:=xlDown

    Range("A1:L1").AutoFilter Field:=12, Criteria1:=Array("SRM"), Operator:=xlAnd

        Range("A2:L" & Drow).SpecialCells(xlCellTypeVisible).Copy

        Worksheets("Customer Ageing").Range("A" & Split(ASRM.Address, "$")(2) + 5).Insert Shift:=xlDown

    Range("A1:L1").AutoFilter Field:=12, Criteria1:=Array("E4I"), Operator:=xlAnd

        Range("A2:L" & Drow).SpecialCells(xlCellTypeVisible).Copy

        Worksheets("Customer Ageing").Range("A" & Split(AE4I.Address, "$")(2) + 5).Insert Shift:=xlDown

    ActiveSheet.ShowAllData

 

   

    'Sums for each entity

    Worksheets("Customer Ageing").Select

    Range("F" & Split(E247.Address, "$")(2)).FormulaR1C1 = "=SUM(R[" & 0 - Split(E247.Address, "$")(2) + Split(A247.Address, "$")(2) + 5 & "]C:R[-1]C)"

        Range("F" & Split(E247.Address, "$")(2)).AutoFill Destination:=Range("F" & Split(E247.Address, "$")(2) & ":K" & Split(E247.Address, "$")(2)), Type:=xlFillDefault

    Range("F" & Split(E247N.Address, "$")(2)).FormulaR1C1 = "=SUM(R[" & 0 - Split(E247N.Address, "$")(2) + Split(A247N.Address, "$")(2) + 5 & "]C:R[-1]C)"

        Range("F" & Split(E247N.Address, "$")(2)).AutoFill Destination:=Range("F" & Split(E247N.Address, "$")(2) & ":K" & Split(E247N.Address, "$")(2)), Type:=xlFillDefault

    Range("F" & Split(EBCA.Address, "$")(2)).FormulaR1C1 = "=SUM(R[" & 0 - Split(EBCA.Address, "$")(2) + Split(ABCA.Address, "$")(2) + 5 & "]C:R[-1]C)"

        Range("F" & Split(EBCA.Address, "$")(2)).AutoFill Destination:=Range("F" & Split(EBCA.Address, "$")(2) & ":K" & Split(EBCA.Address, "$")(2)), Type:=xlFillDefault

    Range("F" & Split(ECAL.Address, "$")(2)).FormulaR1C1 = "=SUM(R[" & 0 - Split(ECAL.Address, "$")(2) + Split(ACAL.Address, "$")(2) + 5 & "]C:R[-1]C)"

        Range("F" & Split(ECAL.Address, "$")(2)).AutoFill Destination:=Range("F" & Split(ECAL.Address, "$")(2) & ":K" & Split(ECAL.Address, "$")(2)), Type:=xlFillDefault

    Range("F" & Split(EEG.Address, "$")(2)).FormulaR1C1 = "=SUM(R[" & 0 - Split(EEG.Address, "$")(2) + Split(AEG.Address, "$")(2) + 5 & "]C:R[-1]C)"

        Range("F" & Split(EEG.Address, "$")(2)).AutoFill Destination:=Range("F" & Split(EEG.Address, "$")(2) & ":K" & Split(EEG.Address, "$")(2)), Type:=xlFillDefault

    Range("F" & Split(EETL.Address, "$")(2)).FormulaR1C1 = "=SUM(R[" & 0 - Split(EETL.Address, "$")(2) + Split(AETL.Address, "$")(2) + 5 & "]C:R[-1]C)"

        Range("F" & Split(EETL.Address, "$")(2)).AutoFill Destination:=Range("F" & Split(EETL.Address, "$")(2) & ":K" & Split(EETL.Address, "$")(2)), Type:=xlFillDefault

    Range("F" & Split(ESRM.Address, "$")(2)).FormulaR1C1 = "=SUM(R[" & 0 - Split(ESRM.Address, "$")(2) + Split(ASRM.Address, "$")(2) + 5 & "]C:R[-1]C)"

        Range("F" & Split(ESRM.Address, "$")(2)).AutoFill Destination:=Range("F" & Split(ESRM.Address, "$")(2) & ":K" & Split(ESRM.Address, "$")(2)), Type:=xlFillDefault

    Range("F" & Split(EE4I.Address, "$")(2)).FormulaR1C1 = "=SUM(R[" & 0 - Split(EE4I.Address, "$")(2) + Split(AE4I.Address, "$")(2) + 5 & "]C:R[-1]C)"

        Range("F" & Split(EE4I.Address, "$")(2)).AutoFill Destination:=Range("F" & Split(EE4I.Address, "$")(2) & ":K" & Split(EE4I.Address, "$")(2)), Type:=xlFillDefault

   

    

    'Change the dates

    Range("F" & Split(A247.Address, "$")(2) + 4).FormulaR1C1 = "=TODAY()"

    Range("G" & Split(A247.Address, "$")(2) + 4).FormulaR1C1 = "=TODAY()"

    Range("H" & Split(A247.Address, "$")(2) + 4).FormulaR1C1 = "=RC[-1]-30"

    Range("I" & Split(A247.Address, "$")(2) + 4).FormulaR1C1 = "=RC[-1]-30"

    Range("J" & Split(A247.Address, "$")(2) + 4).FormulaR1C1 = "=RC[-1]-30"

    Range("K" & Split(A247.Address, "$")(2) + 4).FormulaR1C1 = "=RC[-1]-30"

    Range("J" & Split(A247.Address, "$")(2) + 3).Value = Range("K" & Split(A247.Address, "$")(2) + 4).Value

    Range("I" & Split(A247.Address, "$")(2) + 3).FormulaR1C1 = "=RC[1]+30"

    Range("H" & Split(A247.Address, "$")(2) + 3).FormulaR1C1 = "=RC[1]+30"

    Range("G" & Split(A247.Address, "$")(2) + 3).FormulaR1C1 = "=RC[1]+30"

    Range("F" & Split(A247.Address, "$")(2) + 3 & ":K" & Split(A247.Address, "$")(2) + 4).Value = Range("F" & Split(A247.Address, "$")(2) + 3 & ":K" & Split(A247.Address, "$")(2) + 4).Value

    Range("F" & Split(A247N.Address, "$")(2) + 3 & ":K" & Split(A247N.Address, "$")(2) + 4).Value = Range("F" & Split(A247.Address, "$")(2) + 3 & ":K" & Split(A247.Address, "$")(2) + 4).Value

    Range("F" & Split(ABCA.Address, "$")(2) + 3 & ":K" & Split(ABCA.Address, "$")(2) + 4).Value = Range("F" & Split(A247.Address, "$")(2) + 3 & ":K" & Split(A247.Address, "$")(2) + 4).Value

    Range("F" & Split(ACAL.Address, "$")(2) + 3 & ":K" & Split(ACAL.Address, "$")(2) + 4).Value = Range("F" & Split(A247.Address, "$")(2) + 3 & ":K" & Split(A247.Address, "$")(2) + 4).Value

    Range("F" & Split(AEG.Address, "$")(2) + 3 & ":K" & Split(AEG.Address, "$")(2) + 4).Value = Range("F" & Split(A247.Address, "$")(2) + 3 & ":K" & Split(A247.Address, "$")(2) + 4).Value

    Range("F" & Split(AETL.Address, "$")(2) + 3 & ":K" & Split(AETL.Address, "$")(2) + 4).Value = Range("F" & Split(A247.Address, "$")(2) + 3 & ":K" & Split(A247.Address, "$")(2) + 4).Value

    Range("F" & Split(ASRM.Address, "$")(2) + 3 & ":K" & Split(ASRM.Address, "$")(2) + 4).Value = Range("F" & Split(A247.Address, "$")(2) + 3 & ":K" & Split(A247.Address, "$")(2) + 4).Value

    Range("F" & Split(AE4I.Address, "$")(2) + 3 & ":K" & Split(AE4I.Address, "$")(2) + 4).Value = Range("F" & Split(A247.Address, "$")(2) + 3 & ":K" & Split(A247.Address, "$")(2) + 4).Value

      

    'check if the total balance on the report match the total balance on the raw data

    Dim CusAgeBal As Double

    Dim AgeDatBal As Double

    CusAgeBal = Range("F" & Split(E247.Address, "$")(2)) + Range("F" & Split(E247N.Address, "$")(2)) + Range("F" & Split(EBCA.Address, "$")(2)) + Range("F" & Split(ECAL.Address, "$")(2)) + Range("F" & Split(EEG.Address, "$")(2)) + Range("F" & Split(EETL.Address, "$")(2)) + Range("F" & Split(ESRM.Address, "$")(2)) + Range("F" & Split(EE4I.Address, "$")(2))

    AgeDatBal = WorksheetFunction.Sum(Worksheets("Ageing Data").Range("F:F"))

    If Format(CusAgeBal, "$#,##0.00") = Format(AgeDatBal, "$#,##0.00") Then

    MsgBox "DATA SET MATCHED" & vbCrLf & vbCrLf & "Customer Ageing balance: " & Format(CusAgeBal, "$#,##0.00") & vbCrLf & "Ageing Data balance: " & Format(AgeDatBal, "$#,##0.00"), , "Customer Ageing Report"

      

     

    'Take out the following customer accounts from the EG report (if any)

    Set CUS11606A = Range("A:A").Find(what:="CUS11606", LookIn:=xlValues, lookat:=xlWhole) 'CUS11606

    Set CUS11606B = Range("A:A").Find(what:="CUS11606:MSR40000848", LookIn:=xlValues, lookat:=xlWhole) 'CUS11606:MSR40000848

    Set CUS15281A = Range("A:A").Find(what:="CUS15281:MSR40000419", LookIn:=xlValues, lookat:=xlWhole) 'CUS15281:MSR40000419

    Set CUS15281B = Range("A:A").Find(what:="CUS15281:MSR40000931", LookIn:=xlValues, lookat:=xlWhole) 'CUS15281:MSR40000931

    Set Cus11328A = Range("A:A").Find(what:="Cus11328:SUB03", LookIn:=xlValues, lookat:=xlWhole) 'Cus11328:SUB03

    Set Cus3820 = Range("A:A").Find(what:="Cus3820", LookIn:=xlValues, lookat:=xlWhole) 'Cus3820

    Set CUS1249SUB01 = Range("A:A").Find(what:="CUS1249:SUB01", LookIn:=xlValues, lookat:=xlWhole) 'CUS1249:SUB01

    Set CUS1249SUB02 = Range("A:A").Find(what:="CUS1249:SUB02", LookIn:=xlValues, lookat:=xlWhole) 'CUS1249:SUB02

    Set CUS11328 = Range("A:A").Find(what:="CUS11328", LookIn:=xlValues, lookat:=xlWhole) 'CUS11328

   

        

        If Not Cus3820 Is Nothing Then

            Range(Split(Cus3820.Address, "$")(2) & ":" & Split(Cus3820.Address, "$")(2)).Cut

            Range(Split(EE4I.Address, "$")(2) + 6 & ":" & Split(EE4I.Address, "$")(2) + 6).Insert Shift:=xlDown

        Else

        End If

        If Not CUS1249SUB01 Is Nothing Then

            Range(Split(CUS1249SUB01.Address, "$")(2) & ":" & Split(CUS1249SUB01.Address, "$")(2)).Cut

            Range(Split(EE4I.Address, "$")(2) + 6 & ":" & Split(EE4I.Address, "$")(2) + 6).Insert Shift:=xlDown

        Else

        End If

        If Not CUS1249SUB02 Is Nothing Then

            Range(Split(CUS1249SUB02.Address, "$")(2) & ":" & Split(CUS1249SUB02.Address, "$")(2)).Cut

            Range(Split(EE4I.Address, "$")(2) + 6 & ":" & Split(EE4I.Address, "$")(2) + 6).Insert Shift:=xlDown

        Else

        End If

        If Not CUS11328 Is Nothing Then

            Range(Split(CUS11328.Address, "$")(2) & ":" & Split(CUS11328.Address, "$")(2)).Cut

            Range(Split(EE4I.Address, "$")(2) + 6 & ":" & Split(EE4I.Address, "$")(2) + 6).Insert Shift:=xlDown

        Else

        End If

        If Not Cus11328A Is Nothing Then

            Range(Split(Cus11328A.Address, "$")(2) & ":" & Split(Cus11328A.Address, "$")(2)).Cut

            Range(Split(EE4I.Address, "$")(2) + 6 & ":" & Split(EE4I.Address, "$")(2) + 6).Insert Shift:=xlDown

        Else

        End If

        If Not CUS15281B Is Nothing Then

            Range(Split(CUS15281B.Address, "$")(2) & ":" & Split(CUS15281B.Address, "$")(2)).Cut

            Range(Split(EE4I.Address, "$")(2) + 6 & ":" & Split(EE4I.Address, "$")(2) + 6).Insert Shift:=xlDown

        Else

        End If

        If Not CUS15281A Is Nothing Then

            Range(Split(CUS15281A.Address, "$")(2) & ":" & Split(CUS15281A.Address, "$")(2)).Cut

            Range(Split(EE4I.Address, "$")(2) + 6 & ":" & Split(EE4I.Address, "$")(2) + 6).Insert Shift:=xlDown

        Else

        End If

        If Not CUS11606B Is Nothing Then

            Range(Split(CUS11606B.Address, "$")(2) & ":" & Split(CUS11606B.Address, "$")(2)).Cut

            Range(Split(EE4I.Address, "$")(2) + 6 & ":" & Split(EE4I.Address, "$")(2) + 6).Insert Shift:=xlDown

        Else

        End If

        If Not CUS11606A Is Nothing Then

            Range(Split(CUS11606A.Address, "$")(2) & ":" & Split(CUS11606A.Address, "$")(2)).Cut

            Range(Split(EE4I.Address, "$")(2) + 6 & ":" & Split(EE4I.Address, "$")(2) + 6).Insert Shift:=xlDown

        Else

        End If

   

    'Delete dummy customer accounts

    Set Dummy247 = Range("A:A").Find(what:="Dummy247", LookIn:=xlValues, lookat:=xlWhole) 'Dummy247

        If Not Dummy247 Is Nothing Then

            Range(Split(Dummy247.Address, "$")(2) & ":" & Split(Dummy247.Address, "$")(2)).EntireRow.Delete

        Else

        End If

    Worksheets("Ageing Data").Select

    Set Dummy247 = Range("A:A").Find(what:="Dummy247", LookIn:=xlValues, lookat:=xlWhole) 'Dummy247

        If Not Dummy247 Is Nothing Then

            Range(Split(Dummy247.Address, "$")(2) & ":" & Split(Dummy247.Address, "$")(2)).EntireRow.Delete

        Else

        End If

       

    Worksheets("Customer Ageing").Select

     

    'Grouping

    If Not Split(A247.Address, "$")(2) + 5 - Split(E247.Address, "$")(2) - 2 = -3 Then

    Rows(Split(A247.Address, "$")(2) + 5 & ":" & Split(E247.Address, "$")(2) - 2).Group

    End If

    Rows(Split(A247N.Address, "$")(2) + 5 & ":" & Split(E247N.Address, "$")(2) - 2).Group

    Rows(Split(ABCA.Address, "$")(2) + 5 & ":" & Split(EBCA.Address, "$")(2) - 2).Group

    Rows(Split(ACAL.Address, "$")(2) + 5 & ":" & Split(ECAL.Address, "$")(2) - 2).Group

    Rows(Split(AEG.Address, "$")(2) + 5 & ":" & Split(EEG.Address, "$")(2) - 2).Group

    Rows(Split(AETL.Address, "$")(2) + 5 & ":" & Split(EETL.Address, "$")(2) - 2).Group

    Rows(Split(ASRM.Address, "$")(2) + 5 & ":" & Split(ESRM.Address, "$")(2) - 2).Group

    Rows(Split(AE4I.Address, "$")(2) + 5 & ":" & Split(EE4I.Address, "$")(2) - 2).Group

    Worksheets("Customer Ageing").Outline.ShowLevels RowLevels:=1

   

    'Go back to cell A1

    Application.Goto Worksheets("Customer Ageing").Range("A1"), True

   

    'Save as

    Dim wbstring As String

    wbstring = WorksheetFunction.Text(Range("F" & Split(A247.Address, "$")(2) + 4).Value, "yymmdd") & " Customer Ageing Report - All Companies"

    Range("A2").Value = wbstring

   

    If InStr(ActiveWorkbook.FullName, WorksheetFunction.Text(Range("F" & Split(A247.Address, "$")(2) + 4).Value, "yymmdd")) Then

   

    Dim Msg, Style, Title, Help, Ctxt, Response, MyString

    Msg = "Want to save your changes to '" & ActiveWorkbook.Name & "'?"    ' Define message.

    Style = vbYesNo + vbExclamation + vbDefaultButton2    ' Define buttons.

    Title = "Microsoft Excel"    ' Define title.

    Help = ""    ' Define Help file.

    Ctxt = 1000    ' Define topic context.

        ' Display message.

    Response = MsgBox(Msg, Style, Title, Help, Ctxt)

    If Response = vbYes Then    ' User chose Yes.

    ActiveWorkbook.Save   ' Perform some action.

    MsgBox "Document has been saved."

    Else    ' User chose No.

    MsgBox "Document not saved."   ' Perform some action.

    End If

   

    Else

   

    Msg = "Want to save your document as '" & wbstring & "'" & " under " & ActiveWorkbook.Path & "?"    ' Define message.

    Style = vbYesNo + vbExclamation + vbDefaultButton2    ' Define buttons.

    Title = "Microsoft Excel"    ' Define title.

    Help = ""    ' Define Help file.

    Ctxt = 1000    ' Define topic context.

        ' Display message.

    Response = MsgBox(Msg, Style, Title, Help, Ctxt)

    If Response = vbYes Then    ' User chose Yes.

    ActiveWorkbook.SaveAs FileName:=Application.ActiveWorkbook.Path & "\" & wbstring   ' Perform some action.

    MsgBox "Document has been saved."

    Else    ' User chose No.

    MsgBox "Document not saved."   ' Perform some action.

    End If

   

 

    End If

   

    Else

   

    MsgBox "DATA SET NOT MATCHED" & vbCrLf & vbCrLf & "Customer Ageing balance: " & Format(CusAgeBal, "$#,##0.00") & vbCrLf & "Ageing Data balance: " & Format(AgeDatBal, "$#,##0.00"), vbCritical, "Customer Ageing Report"

    End If

 

 

 

End Sub
Sub Bank_Rec()

'

'   ''''Workbook Bank rec

    Dim rfoundCell As Range

   

    Set rfoundCell = Rows(1).Find(what:="Balance", LookIn:=xlValues, lookat:=xlWhole)

   

    If Not rfoundCell Is Nothing Then

   

        Range(rfoundCell.Address(0, 0)).Value = "Amount"

       

    Dim Collet As String

    ACollet = Split(rfoundCell.Address, "$")(1)

    Dim Arow As Long

    Arow = Cells(Cells.Rows.Count, ACollet).End(xlUp).Row

    Dim Acol As Long

    Acol = Cells(1, Columns.Count).End(xlToLeft).Column

   

    Range(ACollet & "2:" & ACollet & Arow).FormulaR1C1 = "=RC[-1]-RC[-2]"

    Range(ACollet & "2:" & ACollet & Arow).Value = Range(ACollet & "2:" & ACollet & Arow).Value

  

    Range("A1:" & ACollet & Arow).Subtotal GroupBy:=4, Function:=xlSum, TotalList:=Array(Acol), Replace:=True, PageBreaks:=False, SummaryBelowData:=True

    ActiveSheet.Outline.ShowLevels RowLevels:=2

   

 

   

    

    'Go back to cell A1

    Application.Goto Range("A1"), True

       

    

    Else

   

    ''''Raw data from AX

    'Check if you are doing bank reconciliation

    If IsError(Application.Match("Date", Range("A1:J1"), 0)) Or IsError(Application.Match("Description", Range("A1:J1"), 0)) Or IsError(Application.Match("Amount in transaction currency", Range("A1:J1"), 0)) Then

            MsgBox "Please make sure your table contains Date, Description and Amount in transaction currency columns extracted from AX!", vbCritical

            Exit Sub

        End If

 

 

    Dim arrColOrder As Variant, ndx As Integer

    Dim Found As Range, counter As Integer

    'Place the column headers in the end result order you want.

    arrColOrder = Array("Date", "Voucher", "Description", "Amount")

    counter = 1

    Application.ScreenUpdating = False

    For ndx = LBound(arrColOrder) To UBound(arrColOrder)

        Set Found = Rows("1:1").Find(arrColOrder(ndx), LookIn:=xlValues, lookat:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, MatchCase:=False)

        If Not Found Is Nothing Then

           If Found.Column <> counter Then

                Found.EntireColumn.Cut

                Columns(counter).Insert Shift:=xlToRight

                Application.CutCopyMode = False

            End If

            counter = counter + 1

        End If

   Next ndx

   

    Dim Rrow As Long

    Rrow = Cells(Rows.Count, "A").End(xlUp).Row

    With Range("A2:A" & Rrow)

        .Sort Key1:=.Cells(1), Order1:=xlDescending, Orientation:=xlTopToBottom, Header:=xlNo

    End With

 

    Columns("E:J").Delete

    Range("Table[#All]").Copy

    Range("G1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False

    Columns("G:G").NumberFormat = "d/mm/yyyy"

    Columns("A:F").Delete Shift:=xlToLeft

        Columns("A:A").ColumnWidth = 10.5

    Columns("C:C").ColumnWidth = 50

    Columns("D:D").ColumnWidth = 30

 

 

    Range("A1:D" & Rrow).Subtotal GroupBy:=1, Function:=xlSum, TotalList:=Array(4), Replace:=True, PageBreaks:=False, SummaryBelowData:=True

 

    ActiveSheet.Outline.ShowLevels RowLevels:=2

    Dim Srow As Long

    Srow = Cells(Rows.Count, "A").End(xlUp).Row

    Range("E2:E" & Srow - 1).SpecialCells(xlCellTypeVisible).FormulaR1C1 = "=SUBSTITUTE(RC[-4],""/"","""")"

    Range("E2:E" & Srow - 1).Value = Range("E2:E" & Srow - 1).Value

    Range("G2:G" & Srow - 1).SpecialCells(xlCellTypeVisible).FormulaR1C1 = "=RC[-3]-RC[-1]"

   

    

    Set GT = Range("A:A").Find(what:="Grand Total", LookIn:=xlValues, lookat:=xlWhole)

    Range("G" & Split(GT.Address, "$")(2)).FormulaR1C1 = "=SUM(R[" & 1 - Split(GT.Address, "$")(2) & "]C:R[-1]C)"

   

    

    

    Dim i As Long, wbnames As String, bankrecname As String, entity As String

    For i = 1 To Workbooks.Count

        wbnames = wbnames & Workbooks(i).Name & vbLf

    Next

 

    If wbnames Like "*Bank rec*" Then

    bankrecname = Mid(wbnames, WorksheetFunction.Find("Bank rec", wbnames) - 5, 13)

    entity = Workbooks(bankrecname).ActiveSheet.Name

   

    MsgBox "Vlookup with " & entity & "."

   

    For Each cell In Range("E1:E" & Srow)

    If Application.WorksheetFunction.IsText(cell.Value) = True Then

    Range("F" & Split(cell.Address, "$")(2)).FormulaR1C1 = "=VLOOKUP(RC[-1],'[" & bankrecname & ".xlsx]" & entity & "'!C4:C8,5,0)"

    End If

    Next

   

    Else

   

    MsgBox "It is suggested to have the Bank rec workbook open first.", vbExclamation

 
 

End If



    

    'Go back to cell A1

    Application.Goto Range("A1"), True

   

    End If

End Sub
Sub Remove_Carriage_Returns() 'Remove Alt+Enter line breaks

    Dim MyRange As Range

    Application.ScreenUpdating = False

    Application.Calculation = xlCalculationManual

    For Each MyRange In ActiveSheet.UsedRange

        If 0 < InStr(MyRange, Chr(10)) Then

            MyRange = Replace(MyRange, Chr(10), "")

        End If

    Next

   Application.ScreenUpdating = True

    Application.Calculation = xlCalculationAutomatic

End Sub
Sub mySum()

'copy the sum that appears on the status bar when highlighting a range of cells

'http://www.stockkevin.com/2010/03/excel-tip-1-copypaste-sum-of-selected.html#.XhvdRcgza70

'Select 'References' from the 'Tool Menu' and make sure 'Microsoft Forms 2.0 Object Library' is selected. If it's not listed then click 'browse' and select 'Fm20.dll'

Dim MyDataObj As New DataObject

MyDataObj.SetText Application.Sum(Selection)

MyDataObj.PutInClipboard

End Sub
Sub Reorder_PO_Lines()

 

    If Range("U1").Value = "PurchId" Then

    Range("A:R").Delete Shift:=xlToLeft

    Exit Sub

    Else

   

    If Range("A1").Value <> "Purchase order lines" Or Range("C8").Value <> "Purchase order" Then

    MsgBox "Invalid report format!" & vbCrLf & "Please use the Purchase order lines raw report from AX!", vbCritical

    Exit Sub

    End If

    'Go back to cell A1

    Application.Goto Range("A1"), True

   

    Cells.WrapText = False

    Cells.MergeCells = False

 

    Range("1:7").Delete Shift:=xlUp

    Range("F:F,Y:Y,AA:AA").Delete Shift:=xlToLeft

 

    Columns("A:Y").AutoFilter

   

    Dim Rrow As Long

    Rrow = Cells(Cells.Rows.Count, "A").End(xlUp).Row

   

    If Application.WorksheetFunction.CountIf(Range("A2:A" & Rrow), "<1/1/2018") = 0 Then

    Else

    Range("A1:Y" & Rrow).AutoFilter Field:=1, Criteria1:="<1/1/2018"

    Range("A2:Y" & Rrow).SpecialCells(xlCellTypeVisible).EntireRow.Delete

    ActiveSheet.ShowAllData

    End If

 

    If Application.WorksheetFunction.CountIf(Range("I2:I" & Rrow), "Cancelled") = 0 Then

    Else

    Range("A1:Y" & Rrow).AutoFilter Field:=9, Criteria1:="Canceled"

    Range("A2:Y" & Rrow).Select

    Selection.SpecialCells(xlCellTypeVisible).EntireRow.Delete

    ActiveSheet.ShowAllData

    End If

 

    If Application.WorksheetFunction.CountIf(Range("I2:I" & Rrow), "Open order") = 0 Then

    Else

    Range("A1:Y" & Rrow).AutoFilter Field:=9, Criteria1:="Open order"

    Range("A2:Y" & Rrow).Select

    Selection.SpecialCells(xlCellTypeVisible).EntireRow.Delete

    ActiveSheet.ShowAllData

    End If

   

    If Application.WorksheetFunction.CountBlank(Range("J2:J" & Rrow)) = 0 Then

    Else

    Range("A1:Y" & Rrow).AutoFilter Field:=10, Criteria1:=""

    Range("A2:Y" & Rrow).SpecialCells(xlCellTypeVisible).EntireRow.Delete

    ActiveSheet.ShowAllData

    End If

   

    Range("A2:Y" & Rrow).Sort Key1:=Range("J2"), Order1:=xlAscending

   

    End If

 

   

End Sub
Sub Reorder_SO_Lines()

 

    If Range("B1").Value <> "Sales order lines" Or Range("E6").Value <> "Sales order" Then

   

    If Range("C2").Value = "Sales order" Then

    Columns("A:F").Delete

    Exit Sub

    End If

   

    MsgBox "Invalid report format!" & vbCrLf & "Please use the Sales order lines raw report from AX!", vbCritical

    Exit Sub

    End If

    'Go back to cell A1

    Application.Goto Range("A1"), True

   

    Cells.WrapText = False

    Cells.MergeCells = False

 

    Range("1:5").Delete Shift:=xlUp

    Range("B:B,C:C,G:G,I:I").Delete Shift:=xlToLeft

    Range("A:A,F:F,J:J").ColumnWidth = 8

 

    Columns("A:AA").AutoFilter

   

    

    Dim Rrow As Long

    Rrow = Cells(Cells.Rows.Count, "A").End(xlUp).Row

    Dim rng As Range

    Set rng = Range("A1").CurrentRegion

    'Check if the transactions are from more than one division

    Dim MyRange As Range

    Dim myValue

    Dim allSame As Boolean

    'Set column to check

    Set MyRange = Range("J2:J" & Rrow)

    'Get first value from myRange

    myValue = MyRange(1, 1).Value

    allSame = (WorksheetFunction.CountA(MyRange) = WorksheetFunction.CountIf(MyRange, myValue))

    If allSame = False Then

    Dim RemainStr As Variant

    'Keeping looping until getting a correct division

    Do

    'Retrieve an answer from the user

    RemainStr = Application.InputBox("Which Division to remain?", , "ADM/CRS/EBC/EPS/ESA", Type:=2)

    'Check if user selected cancel button

    If TypeName(RemainStr) = "Boolean" Then MsgBox "Please manually choose the division.": GoTo Line1

    Loop Until UCase(RemainStr) = "ADM" Or UCase(RemainStr) = "CRS" Or UCase(RemainStr) = "EBC" Or UCase(RemainStr) = "EPS" Or UCase(RemainStr) = "ESA"

    'Delete rows that are not from the division chosen

    rng.AutoFilter Field:=10, Criteria1:=Array("<>" & RemainStr), Operator:=xlAnd

    rng.Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete

    ActiveSheet.ShowAllData

    End If

 

Line1:

    If Application.WorksheetFunction.CountIf(Range("A2:A" & Rrow), "<1/1/2018") = 0 Then

    GoTo Line2

    Else

    Range("A1:A" & Rrow).AutoFilter Field:=1, Criteria1:="<1/1/2018"

    Range("A2:AA" & Rrow).SpecialCells(xlCellTypeVisible).EntireRow.Delete

    ActiveSheet.ShowAllData

    End If

 

Line2:

    If Application.WorksheetFunction.CountIfs(Range("F2:F" & Rrow), "Canceled", Range("V2:V" & Rrow), "None") = 0 Then

    GoTo Line3

    Else

    Range("A1:F" & Rrow).AutoFilter Field:=6, Criteria1:="Canceled"

    Range("A1:V" & Rrow).AutoFilter Field:=22, Criteria1:="None", Operator:=xlOr, Criteria2:="Canceled"

    Range("A2:AA" & Rrow).SpecialCells(xlCellTypeVisible).EntireRow.Delete

    ActiveSheet.ShowAllData

    End If

 

   

Line3:

    If Application.WorksheetFunction.CountIfs(Range("O2:O" & Rrow), 0, Range("V2:V" & Rrow), "None") = 0 Then

    Exit Sub

    Else

    Range("A1:O" & Rrow).AutoFilter Field:=15, Criteria1:="0.00"

    Range("A1:V" & Rrow).AutoFilter Field:=22, Criteria1:="None"

    Range("A2:AA" & Rrow).SpecialCells(xlCellTypeVisible).EntireRow.Delete

    ActiveSheet.ShowAllData

    End If

      

End Sub

Sub GST()

   

    'Make sure the format of the raw data is correct.

    Dim APcol As Long

    APcol = Worksheets("AP").Cells(1, Worksheets("AP").Columns.Count).End(xlToLeft).Column

    If APcol = 17 Then

    Else: MsgBox "Please make sure Worksheet AP has a table with exact 17 columns!"

    Exit Sub: End If

    If Worksheets("AP").Range("A1").Value = "Tax invoice account" And Worksheets("AP").Range("J1").Value = "GST" Then

    Else: MsgBox "Please makse sure on Worksheet AP, Cell A1 is 'Tax invoice account' and Cell J1 is 'GST'!"

    Exit Sub: End If

    If Worksheets("VL").Range("A1").Value = "Vendor account" And Worksheets("VL").Range("D1").Value = "GST group" Then

    Else: MsgBox "Please makse sure on Worksheet VL, Cell A1 is 'Vendor account' and Cell D1 is 'GST group'!"

    Exit Sub: End If

    Dim ARcol As Long

    ARcol = Worksheets("AR").Cells(1, Worksheets("AR").Columns.Count).End(xlToLeft).Column

    If ARcol = 17 Then

    Else: MsgBox "Please make sure Worksheet AR has a table with exact 17 columns!"

    Exit Sub: End If

    If Worksheets("AR").Range("A1").Value = "Tax invoice account" And Worksheets("AR").Range("I1").Value = "GST" Then

    Else: MsgBox "Please makse sure on Worksheet AR, Cell A1 is 'Tax invoice account' and Cell I1 is 'GST'!"

    Exit Sub: End If

    If Worksheets("CL").Range("C1").Value = "GST group" Then

    Else: MsgBox "Please makse sure on Worksheet CL, Cell C1 is 'GST group'!"

    Exit Sub: End If

    If Worksheets("GL").Range("B1").Value = "Voucher" And Worksheets("GL").Range("E1").Value = "Description" Then

    Else: MsgBox "Please makse sure on Worksheet GL, Cell B1 is 'Voucher' and Cell E1 is 'Description'!"

    Exit Sub: End If

    'Make sure no filiters have been applied on the raw data.

    On Error Resume Next

    Worksheets("AP").ShowAllData

    Worksheets("VL").ShowAllData

    Worksheets("AR").ShowAllData

    Worksheets("CL").ShowAllData

    Worksheets("GL").ShowAllData

    On Error GoTo 0

 

   

    

    '''GL

    Worksheets("GL").Select

    Range("I1").FormulaR1C1 = "Account Number"

    Range("I1").Interior.Color = 49407

    Range("I1").Font.ColorIndex = xlAutomatic

    If Not Columns("I").NumberFormat = "General" Then

    Columns("I").NumberFormat = "General"

    End If

   

    Dim GLrow0 As Long

    GLrow0 = Cells(Rows.Count, "A").End(xlUp).Row

    Range("I2:I" & GLrow0).FormulaR1C1 = "=LEFT([@[Ledger account]],5)"

   

    

    '''AP GSTCode

    Application.DisplayAlerts = False

    Worksheets("AP GSTCode").Delete

    Application.DisplayAlerts = True

    Worksheets("AP").Copy After:=Worksheets("ITI")

    ActiveSheet.Name = "AP GSTCode"

    With Worksheets("AP GSTCode").Tab

        .ThemeColor = xlThemeColorLight2

        .TintAndShade = 0.399975585192419

    End With

    Range("R1").Value = "GST Code from VL"

    Range("S1").Value = "Code Check"

    Range("T1").Value = "Comment"

    Range("R1:T1").Select

    With Selection.Interior

        .Color = 49407

    End With

    With Selection.Font

        .ColorIndex = xlAutomatic

    End With

    Range("R2").FormulaR1C1 = "=INDEX(VL!C[-14],MATCH('AP GSTCode'!RC1,VL!C1,0))"

    Range("S2").FormulaR1C1 = "=IF((RC[-9]<>0)*(RC[-1]=""GST"")+(RC[-9]=0)*(RC[-1]<>""GST""),""OK"",""Review"")"

    Range("A:C,G:G").ColumnWidth = 12

    Range("N:Q").ColumnWidth = 4

    Dim AProw As Long

    AProw = Cells(Cells.Rows.Count, "H").End(xlUp).Row

    Range("R2:T" & AProw).Interior.ThemeColor = xlThemeColorDark1

    'Optus Mobile Acct# RS949

    If Application.WorksheetFunction.CountIf(Range("A2:A" & AProw), "Ven1004") > 0 Then

    Range("A1:T" & AProw).AutoFilter Field:=1, Criteria1:="Ven1004"

    Range("T2:T" & AProw).SpecialCells(xlCellTypeVisible).Value = "GST in EG"

    ActiveSheet.ShowAllData

    End If

    'ASIC

    If Application.WorksheetFunction.CountIf(Range("A2:A" & AProw), "Ven24") > 0 Then

    Range("A1:T" & AProw).AutoFilter Field:=1, Criteria1:="Ven24"

    Range("T2:T" & AProw).SpecialCells(xlCellTypeVisible).Value = "AISC FEE no GST"

    ActiveSheet.ShowAllData

    End If

    'ACMA

    If Application.WorksheetFunction.CountIf(Range("A2:A" & AProw), "Ven253") > 0 Then

    Range("A1:T" & AProw).AutoFilter Field:=1, Criteria1:="Ven253"

    Range("T2:T" & AProw).SpecialCells(xlCellTypeVisible).Value = "NO GST"

    ActiveSheet.ShowAllData

    End If

    'ATO

    If Application.WorksheetFunction.CountIf(Range("A2:A" & AProw), "Ven38") > 0 Then

    Range("A1:T" & AProw).AutoFilter Field:=1, Criteria1:="Ven38"

    Range("T2:T" & AProw).SpecialCells(xlCellTypeVisible).Value = "NO GST"

    ActiveSheet.ShowAllData

    End If

    'Expense claim

    If Application.WorksheetFunction.CountIf(Range("H2:H" & AProw), "EXP20*") > 0 Then

    Range("A1:T" & AProw).AutoFilter Field:=8, Criteria1:="EXP20*"

    Range("T2:T" & AProw).SpecialCells(xlCellTypeVisible).Value = "Expense claim"

    ActiveSheet.ShowAllData

    End If

    '$0 Tax invoice amount

    If Application.WorksheetFunction.CountIfs(Range("K2:K" & AProw), 0, Range("J2:J" & AProw), 0) > 0 Then

    Range("A1:T" & AProw).AutoFilter Field:=11, Criteria1:="0.00"

    Range("A1:T" & AProw).AutoFilter Field:=10, Criteria1:="0.00"

    Application.DisplayAlerts = False

    Range("A2:T" & AProw).SpecialCells(xlCellTypeVisible).Delete

    Application.DisplayAlerts = True

    ActiveSheet.ShowAllData

 

    End If

   

    AProw = Cells(Cells.Rows.Count, "H").End(xlUp).Row

    If Application.WorksheetFunction.CountIf(Range("S2:S" & AProw), "Review") = 0 Then

    Else

    Range("A1:T" & AProw).AutoFilter Field:=19, Criteria1:="Review"

    End If

   

    'Go back to cell A1

    Application.Goto Range("A1"), True

   

    

    

    '''AR GSTCode

    Application.DisplayAlerts = False

    Worksheets("AR GSTCode").Delete

    Application.DisplayAlerts = True

    Worksheets("AR").Copy After:=Worksheets("AP GSTAmount")

    ActiveSheet.Name = "AR GSTCode"

    With Worksheets("AR GSTCode").Tab

        .ThemeColor = xlThemeColorLight2

        .TintAndShade = 0.399975585192419

    End With

    Range("R1").Value = "GST Code from VL"

    Range("S1").Value = "Code Check"

    Range("T1").Value = "Comment"

    Range("R1:T1").Interior.Color = 49407

    Range("R1:T1").Font.ColorIndex = xlAutomatic

    Range("R:T").NumberFormat = "General"

   

 

    Range("R2").FormulaR1C1 = "=INDEX(CL!C[-15],MATCH('AR GSTCode'!RC[-17],CL!C[-17],0))"

    Range("S2").FormulaR1C1 = "=IF((RC[-10]<>0)*(RC[-1]=""GST"")+(RC[-10]=0)*(RC[-1]<>""GST""),""OK"",""Review"")"

    Range("A:C,G:G").ColumnWidth = 12

    Range("N:Q").ColumnWidth = 4

    Dim ARrow As Long

    ARrow = Cells(Cells.Rows.Count, "H").End(xlUp).Row

    Range("R2:T" & ARrow).Select

    With Selection.Interior

        .ThemeColor = xlThemeColorDark1

    End With

   

    'Tax invoice amount < $0.06

    If Application.WorksheetFunction.CountIf(Range("J2:J" & ARrow), "Review") < 0.06 Then

    Range("A1:T" & ARrow).AutoFilter Field:=10, Criteria1:="<0.06"

    If Application.WorksheetFunction.CountIf(Range("S2:S" & ARrow), "Review") > 0 Then

    ARrow = Cells(Cells.Rows.Count, "H").End(xlUp).Row

    Range("A1:T" & ARrow).AutoFilter Field:=19, Criteria1:="Review"

    Range("T2:T" & ARrow).SpecialCells(xlCellTypeVisible).Value = "Tax invoice amount < $0.06"

    ActiveSheet.ShowAllData

    Else: ActiveSheet.ShowAllData

    End If

    End If

   

    '$0 Tax invoice amount

    If Application.WorksheetFunction.CountIfs(Range("J2:J" & ARrow), 0, Range("I2:I" & ARrow), 0) > 0 Then

    Range("A1:T" & ARrow).AutoFilter Field:=10, Criteria1:="0.00"

    Range("A1:T" & ARrow).AutoFilter Field:=9, Criteria1:="0.00"

    Application.DisplayAlerts = False

    Range("A2:T" & ARrow).SpecialCells(xlCellTypeVisible).Delete

    Application.DisplayAlerts = True

    ActiveSheet.ShowAllData

    End If

   

    

    ARrow = Cells(Cells.Rows.Count, "H").End(xlUp).Row

   

    If Application.WorksheetFunction.CountIf(Range("S2:S" & ARrow), "Review") > 0 Then

    Range("A1:T" & ARrow).AutoFilter Field:=19, Criteria1:="Review"

    End If

   

    'Go back to cell A1

   Application.Goto Range("A1"), True

   

    

    

    '''AP GSTAmount

    Worksheets("AP GSTAmount").Select

    Dim objPT As PivotTable, iCount As Integer

    For iCount = ActiveSheet.PivotTables.Count To 1 Step -1

    Set objPT = ActiveSheet.PivotTables(iCount)

    objPT.PivotSelect ""

    Selection.Clear

    Next iCount

   

    Range("N1:P2").ClearContents

     

    Dim PSheet As Worksheet

    Dim DSheet As Worksheet

    Dim PCache As PivotCache

    Dim PTable As PivotTable

    Dim PRange As Range

    Dim LastRow As Long

    Dim LastCol As Long

    'Insert a New Blank Worksheet

    On Error Resume Next

    Sheets.Add After:=Sheets("AP GSTAmount")

    ActiveSheet.Name = "PivotTable"

    Set PSheet = Worksheets("PivotTable")

    Set DSheet = Worksheets("AP GSTCode")

    'Define Data Range

    LastRow = DSheet.Cells(Rows.Count, 1).End(xlUp).Row

    LastCol = DSheet.Cells(1, Columns.Count).End(xlToLeft).Column

    Set PRange = DSheet.Cells(1, 1).Resize(LastRow, LastCol)

    'Define Pivot Cache

    Set PCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=PRange).CreatePivotTable(TableDestination:=PSheet.Cells(1, 3), TableName:="AP GSTCode")

    'Insert Row Fields

    With ActiveSheet.PivotTables("AP GSTCode").PivotFields("Voucher")

        .Orientation = xlRowField

        .Position = 1

    End With

    With ActiveSheet.PivotTables("AP GSTCode").PivotFields("GST Code from VL")

        .Orientation = xlRowField

        .Position = 2

    End With

    With ActiveSheet.PivotTables("AP GSTCode").PivotFields("Currency")

        .Orientation = xlRowField

        .Position = 3

    End With

    ActiveSheet.PivotTables("AP GSTCode").AddDataField ActiveSheet.PivotTables("AP GSTCode").PivotFields("GST"), "Sum of GST", xlSum

    Range("A6").Select

    With ActiveSheet.PivotTables("AP GSTCode")

        .InGridDropZones = True

        .RowAxisLayout xlTabularRow

    End With

    ActiveSheet.PivotTables("AP GSTCode").PivotFields("Voucher").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)

    Range("B6").Select

    ActiveSheet.PivotTables("AP GSTCode").PivotFields("GST Code from VL").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)

    ActiveSheet.PivotTables("AP GSTCode").PivotSelect "", xlDataAndLabel, True

    Selection.Copy

    Worksheets("AP GSTAmount").Range("A3").PasteSpecial

    

    Application.Wait (Now + TimeValue("0:00:1.5"))

   

    Application.DisplayAlerts = False

    Worksheets("PivotTable").Delete

    Application.DisplayAlerts = True

   

    

    'Insert a New Blank Worksheet

    On Error Resume Next

    Sheets.Add After:=Sheets("AP GSTAmount")

    ActiveSheet.Name = "PivotTable"

    Set PSheet = Worksheets("PivotTable")

    Set DSheet = Worksheets("GL")

    'Define Data Range

    LastRow = DSheet.Cells(Rows.Count, 1).End(xlUp).Row

    LastCol = DSheet.Cells(1, Columns.Count).End(xlToLeft).Column

    Set PRange = DSheet.Cells(1, 1).Resize(LastRow, LastCol)

    'Define Pivot Cache

    Set PCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=PRange).CreatePivotTable(TableDestination:=PSheet.Cells(1, 10), TableName:="GL")

    'Insert Row Fields

    With ActiveSheet.PivotTables("GL").PivotFields("Account Number")

        .Orientation = xlRowField

        .Position = 1

    End With

    With ActiveSheet.PivotTables("GL").PivotFields("Account Number")

        .Orientation = xlPageField

        .Position = 1

    End With

    With ActiveSheet.PivotTables("GL").PivotFields("Voucher")

        .Orientation = xlRowField

        .Position = 1

    End With

    With ActiveSheet.PivotTables("GL").PivotFields("Currency")

        .Orientation = xlRowField

        .Position = 2

    End With

    ActiveSheet.PivotTables("GL").AddDataField ActiveSheet.PivotTables("GL").PivotFields("Amount"), "Sum of Amount", xlSum

    Range("J7").Select

    With ActiveSheet.PivotTables("GL")

        .InGridDropZones = True

        .RowAxisLayout xlTabularRow

    End With

    Range("J10").Select

    ActiveSheet.PivotTables("GL").PivotFields("Voucher").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)

    ActiveSheet.PivotTables("GL").PivotSelect "", xlDataAndLabel, True

    ActiveSheet.PivotTables("GL").PivotFields("Account Number").ClearAllFilters

    ActiveSheet.PivotTables("GL").PivotFields("Account Number").CurrentPage = "16100"

    Selection.Copy

    Worksheets("AP GSTAmount").Range("J1").PasteSpecial

   

    Application.Wait (Now + TimeValue("0:00:1.5"))

   

    Application.DisplayAlerts = False

    Worksheets("PivotTable").Delete

    Application.DisplayAlerts = True

   

    Worksheets("AP GSTAmount").Select

    Dim Vrow As Long

    Vrow = Cells(Cells.Rows.Count, "F").End(xlUp).Row

    Range("F5:H" & Vrow).ClearContents

    Dim Trow As Long

    Trow = Cells(Cells.Rows.Count, "D").End(xlUp).Row

 

    Range("F5").Select

    Selection.FormulaR1C1 = "=VLOOKUP(RC[-5],C[4],1,FALSE)"

    Selection.AutoFill Destination:=Range("F5:F" & Trow - 1)

    Range("G5").Select

    Selection.FormulaR1C1 = "=VLOOKUP(RC[-6],C[3]:C[5],3,FALSE)"

    Selection.AutoFill Destination:=Range("G5:G" & Trow - 1)

    Range("H5").Select

    Selection.FormulaR1C1 = "=IFERROR(RC[-4]-RC[-1],RC[-4])"

    Selection.AutoFill Destination:=Range("H5:H" & Trow - 1)

   

    Vrow = Cells(Cells.Rows.Count, "N").End(xlUp).Row

    Range("N5:R" & Vrow).ClearContents

    Trow = Cells(Cells.Rows.Count, "L").End(xlUp).Row

   

    Range("N5").Select

    Selection.FormulaR1C1 = "=VLOOKUP(RC[-4],C[-13],1,FALSE)"

    Selection.AutoFill Destination:=Range("N5:N" & Trow - 1)

    Range("O5").Select

    Selection.FormulaR1C1 = "=VLOOKUP(RC[-5],C[-14]:C[-11],4,FALSE)"

    Selection.AutoFill Destination:=Range("O5:O" & Trow - 1)

    Range("P5").Select

    Selection.FormulaR1C1 = "=IFERROR(RC[-4]-RC[-1],RC[-4])"

    Selection.AutoFill Destination:=Range("P5:P" & Trow - 1)

    Range("Q5").Select

    Selection.FormulaR1C1 = "=IF(IFERROR(SEARCH(""PAYMENT BY AUTHORITY"",VLOOKUP(RC[-7],GL!C[-15]:C[-12],4,0)),0)+IFERROR(SEARCH(""WITHDRAWAL WESTPAC"",VLOOKUP(RC[-7],GL!C[-15]:C[-12],4,0)),0)>0,""Merchant fee"","""")"

    Selection.AutoFill Destination:=Range("Q5:Q" & Trow - 1)

    For Each cell In Range("J5:J" & Trow)

    If InStr(1, cell.Value, "CPA") > 0 Then

    Range("Q" & Split(cell.Address, "$")(2)).FormulaR1C1 = "=IF(ISNUMBER((SEARCH(""CPA"",VLOOKUP(RC[-7],GL!C[-15]:C[-12],1,0))))+ISNUMBER(SEARCH(""CUS"",VLOOKUP(RC[-7],GL!C[-15]:C[-12],4,0)))>1,""AAPT Invoice, GST under EG"","""")"

    End If

    Next

 

 

 

    Range("Q5:Q" & Trow - 1).Value = Range("Q5:Q" & Trow - 1).Value

   

    For Each cell In Range("N5:N" & Trow)

    If Application.WorksheetFunction.IsError(cell.Value) = True Then

    Range("R" & Split(cell.Address, "$")(2)).FormulaR1C1 = "=VLOOKUP(RC[-8],GL!C[-16]:C[-13],4,0)"

   

    Range("N1").Value = "Merchant fee"

    Range("N2").Value = "AAPT Invoice, GST under EG"

    Range("P1").FormulaR1C1 = "=SUMIF(R[4]C[1]:R[" & Trow - 2 & "]C[1],""Merchant fee"",R[4]C:R[" & Trow - 2 & "]C)"

    Range("P2").FormulaR1C1 = "=SUMIF(R[3]C[1]:R[" & Trow - 3 & "]C[1],""AAPT Invoice, GST under EG"",R[3]C:R[" & Trow - 2 & "]C)"

    End If

    Next

   

    For Each cell In Range("R5:R" & Trow)

    If InStr(1, cell.Value, "ConnectWise") > 0 = True Then

    Range("Q" & Split(cell.Address, "$")(2)).FormulaR1C1 = "=IF(ISNUMBER(SEARCH(""VPA"",RC[-7])),""AAPT Invoice, GST under EG"","""")"

   

    Range("N1").Value = "Merchant fee"

    Range("N2").Value = "AAPT Invoice, GST under EG"

    Range("P1").FormulaR1C1 = "=SUMIF(R[4]C[1]:R[" & Trow - 2 & "]C[1],""Merchant fee"",R[4]C:R[" & Trow - 2 & "]C)"

    Range("P2").FormulaR1C1 = "=SUMIF(R[3]C[1]:R[" & Trow - 3 & "]C[1],""AAPT Invoice, GST under EG"",R[3]C:R[" & Trow - 2 & "]C)"

    End If

    Next

   

    

    Range("Q5:Q" & Trow).Value = Range("Q5:Q" & Trow).Value

   

    

    Worksheets("AP GSTAmount").Select

    Columns("A:A,F:F,J:J,N:N").ColumnWidth = 14

    Columns("B:C").ColumnWidth = 6

    Columns("I:I").ColumnWidth = 10

    Columns("D:D,G:G,L:L,O:O,S:S").ColumnWidth = 5

    Columns("H:H,I:I,P:P,T:U").ColumnWidth = 8.5

    Columns("K:K").ColumnWidth = 7.5

    Columns("Q:Q").ColumnWidth = 26.5

 

   

 

    '''AR GSTAmount

    Worksheets("AR GSTAmount").Select

    For iCount = ActiveSheet.PivotTables.Count To 1 Step -1

    Set objPT = ActiveSheet.PivotTables(iCount)

    objPT.PivotSelect ""

    Selection.Clear

    Next iCount

    Range("L1").ClearContents

   

 

    'Insert a New Blank Worksheet

    On Error Resume Next

    Sheets.Add After:=Sheets("AR GSTAmount")

    ActiveSheet.Name = "PivotTable"

    Set PSheet = Worksheets("PivotTable")

    Set DSheet = Worksheets("AR GSTCode")

    'Define Data Range

    LastRow = DSheet.Cells(Rows.Count, 1).End(xlUp).Row

    LastCol = DSheet.Cells(1, Columns.Count).End(xlToLeft).Column

    Set PRange = DSheet.Cells(1, 1).Resize(LastRow, LastCol)

    'Define Pivot Cache

    Set PCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=PRange).CreatePivotTable(TableDestination:=PSheet.Cells(1, 3), TableName:="AR GSTCode")

    'Insert Row Fields

    With ActiveSheet.PivotTables("AR GSTCode").PivotFields("Voucher")

        .Orientation = xlRowField

        .Position = 1

    End With

    With ActiveSheet.PivotTables("AR GSTCode").PivotFields("GST Code from VL")

        .Orientation = xlRowField

        .Position = 2

    End With

    With ActiveSheet.PivotTables("AR GSTCode").PivotFields("Currency")

        .Orientation = xlRowField

        .Position = 3

    End With

    ActiveSheet.PivotTables("AR GSTCode").AddDataField ActiveSheet.PivotTables("AR GSTCode").PivotFields("GST"), "Sum of GST", xlSum

    Range("A6").Select

    With ActiveSheet.PivotTables("AR GSTCode")

        .InGridDropZones = True

        .RowAxisLayout xlTabularRow

    End With

    ActiveSheet.PivotTables("AR GSTCode").PivotFields("Voucher").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)

    Range("B6").Select

    ActiveSheet.PivotTables("AR GSTCode").PivotFields("GST Code from VL").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)

    ActiveSheet.PivotTables("AR GSTCode").PivotSelect "", xlDataAndLabel, True

    Selection.Copy

    Worksheets("AR GSTAmount").Range("A3").PasteSpecial

   

    Application.Wait (Now + TimeValue("0:00:1.5"))

   

    Application.DisplayAlerts = False

    Worksheets("PivotTable").Delete

    Application.DisplayAlerts = True

   

    'Insert a New Blank Worksheet

    On Error Resume Next

    Sheets.Add After:=Sheets("AR GSTAmount")

    ActiveSheet.Name = "PivotTable"

    Set PSheet = Worksheets("PivotTable")

    Set DSheet = Worksheets("GL")

    'Define Data Range

    LastRow = DSheet.Cells(Rows.Count, 1).End(xlUp).Row

    LastCol = DSheet.Cells(1, Columns.Count).End(xlToLeft).Column

    Set PRange = DSheet.Cells(1, 1).Resize(LastRow, LastCol)

    'Define Pivot Cache

    Set PCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=PRange).CreatePivotTable(TableDestination:=PSheet.Cells(1, 10), TableName:="GL")

    'Insert Row Fields

    With ActiveSheet.PivotTables("GL").PivotFields("Account Number")

        .Orientation = xlRowField

        .Position = 1

    End With

    With ActiveSheet.PivotTables("GL").PivotFields("Account Number")

        .Orientation = xlPageField

        .Position = 1

    End With

    With ActiveSheet.PivotTables("GL").PivotFields("Voucher")

        .Orientation = xlRowField

        .Position = 1

    End With

    With ActiveSheet.PivotTables("GL").PivotFields("Currency")

        .Orientation = xlRowField

        .Position = 2

    End With

    ActiveSheet.PivotTables("GL").AddDataField ActiveSheet.PivotTables("GL").PivotFields("Amount"), "Sum of Amount", xlSum

    Range("J7").Select

    With ActiveSheet.PivotTables("GL")

        .InGridDropZones = True

        .RowAxisLayout xlTabularRow

    End With

    Range("J10").Select

    ActiveSheet.PivotTables("GL").PivotFields("Voucher").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)

    ActiveSheet.PivotTables("GL").PivotSelect "", xlDataAndLabel, True

    ActiveSheet.PivotTables("GL").PivotSelect "", xlDataAndLabel, True

    ActiveSheet.PivotTables("GL").PivotFields("Account Number").ClearAllFilters

    ActiveSheet.PivotTables("GL").PivotFields("Account Number").CurrentPage = "24100"

    Selection.Copy

    Worksheets("AR GSTAmount").Range("J1").PasteSpecial

   

    Application.Wait (Now + TimeValue("0:00:1.5"))

   

    Application.DisplayAlerts = False

    Worksheets("PivotTable").Delete

    Application.DisplayAlerts = True

   

    Worksheets("AR GSTAmount").Select

    Vrow = Cells(Cells.Rows.Count, "F").End(xlUp).Row

    Range("F5:H" & Vrow).ClearContents

 

    Trow = Cells(Cells.Rows.Count, "D").End(xlUp).Row

 

    Range("F5").Select

    Selection.FormulaR1C1 = "=VLOOKUP(RC[-5],C[4],1,FALSE)"

    Selection.AutoFill Destination:=Range("F5:F" & Trow - 1)

    Range("G5").Select

    Selection.FormulaR1C1 = "=VLOOKUP(RC[-6],C[3]:C[5],3,FALSE)"

    Selection.AutoFill Destination:=Range("G5:G" & Trow - 1)

    Range("H5").Select

    Selection.FormulaR1C1 = "=IFERROR(RC[-4]+RC[-1],RC[-4])"

    Selection.AutoFill Destination:=Range("H5:H" & Trow - 1)

   

    Vrow = Cells(Cells.Rows.Count, "N").End(xlUp).Row

    Range("N5:Q" & Vrow).ClearContents

    Trow = Cells(Cells.Rows.Count, "L").End(xlUp).Row

   

    Range("N5").Select

    Selection.FormulaR1C1 = "=VLOOKUP(RC[-4],C[-13],1,FALSE)"

    Selection.AutoFill Destination:=Range("N5:N" & Trow - 1)

    Range("O5").Select

    Selection.FormulaR1C1 = "=VLOOKUP(RC[-5],C[-14]:C[-11],4,FALSE)"

    Selection.AutoFill Destination:=Range("O5:O" & Trow - 1)

    Range("P5").Select

    Selection.FormulaR1C1 = "=IFERROR(RC[-4]+RC[-1],RC[-4])"

    Selection.AutoFill Destination:=Range("P5:P" & Trow - 1)

   

    Worksheets("AR GSTAmount").Select

    Columns("A:A,F:F,J:J,N:N").ColumnWidth = 14

    Columns("B:C").ColumnWidth = 6

    Columns("I:I").ColumnWidth = 10

    Columns("D:D,G:G,L:L,O:O,S:S").ColumnWidth = 5

    Columns("H:H,I:I,P:P,T:U").ColumnWidth = 8.5

    Columns("K:K").ColumnWidth = 7.5

    Columns("Q:Q").ColumnWidth = 26.5

   

    

    

    

 

    '''GSTAdj

Line1:

    Application.DisplayAlerts = False

    Worksheets("GSTAdj").Delete

    Application.DisplayAlerts = True

    Worksheets("GL").Copy After:=Worksheets("AR GSTAmount")

    ActiveSheet.Name = "GSTAdj"

    With Worksheets("GSTAdj").Tab

        .ThemeColor = xlThemeColorLight2

        .TintAndShade = 0.399975585192419

    End With

    If Worksheets("GSTAdj").ListObjects.Count > 0 Then

    ActiveSheet.ListObjects(1).Unlist

    End If

    Dim GLrow As Long

    GLrow = Cells(Cells.Rows.Count, "F").End(xlUp).Row

    Range("A1:I" & GLrow).AutoFilter

    Range("A" & GLrow + 2).Value = "Grand Total"

    Range("H" & GLrow + 2).NumberFormat = "General"

    Range("H" & GLrow + 2).FormulaR1C1 = "=SUM(R[" & -GLrow & "]C:R[-1]C)"

    Range("A" & GLrow + 2 & ":H" & GLrow + 2).Select

    With Selection.Borders(xlEdgeTop)

        .LineStyle = xlContinuous

        .Weight = xlThin

    End With

    With Selection.Borders(xlEdgeBottom)

        .LineStyle = xlDouble

        .Weight = xlThick

    End With

    Range("A" & GLrow + 2 & ":H" & GLrow + 2).Font.Bold = True

   

    If Application.WorksheetFunction.CountIf(Range("I2:I" & GLrow), "24300") > 0 Then

    Range("A1:I" & GLrow).AutoFilter Field:=9, Criteria1:="24300"

    ''''''''Delete hidden rows''''''''

    Dim oRow As Range, rng As Range

    Dim myRows As Range

    With Sheets("GSTAdj")

        Set myRows = Intersect(.Range("A:A").EntireRow, .UsedRange)

        If myRows Is Nothing Then GoTo Line2

    End With

    For Each oRow In myRows.Columns(1).Cells

        If oRow.EntireRow.Hidden Then

            If rng Is Nothing Then

                Set rng = oRow

            Else

                Set rng = Union(rng, oRow)

            End If

        End If

    Next

    If Not rng Is Nothing Then rng.EntireRow.Delete

    ''''''''Delete hidden rows''''''''

    Else: Range("A2:I" & GLrow).EntireRow.Delete

    End If

    Application.Goto Range("A1"), True

   

    '''FA

Line2:

    Application.DisplayAlerts = False

    Worksheets("FA").Delete

    Application.DisplayAlerts = True

    Worksheets("GL").Copy After:=Worksheets("GSTAdj")

    ActiveSheet.Name = "FA"

    With Worksheets("FA").Tab

        .ThemeColor = xlThemeColorLight2

        .TintAndShade = 0.399975585192419

    End With

    If Worksheets("FA").ListObjects.Count > 0 Then

    ActiveSheet.ListObjects(1).Unlist

    End If

    GLrow = Cells(Cells.Rows.Count, "F").End(xlUp).Row

    Range("A1:I" & GLrow).AutoFilter

    Range("A" & GLrow + 2).Value = "Grand Total"

    Range("H" & GLrow + 2).NumberFormat = "General"

    Range("H" & GLrow + 2).FormulaR1C1 = "=SUM(R[" & -GLrow & "]C:R[-1]C)"

    Range("A" & GLrow + 2 & ":H" & GLrow + 2).Select

    With Selection.Borders(xlEdgeTop)

        .LineStyle = xlContinuous

        .Weight = xlThin

    End With

    With Selection.Borders(xlEdgeBottom)

        .LineStyle = xlDouble

        .Weight = xlThick

    End With

    Range("A" & GLrow + 2 & ":H" & GLrow + 2).Font.Bold = True

   

    If Application.WorksheetFunction.CountIf(Range("I2:I" & GLrow), "17300") > 0 Then

    Range("A1:I" & GLrow).AutoFilter Field:=9, Criteria1:="17300"

    ''''''''Delete hidden rows''''''''

    Dim oRow2 As Range, rng2 As Range

    Dim myRows2 As Range

    With Sheets("FA")

        Set myRows2 = Intersect(.Range("A:A").EntireRow, .UsedRange)

        If myRows2 Is Nothing Then GoTo Line3

    End With

    For Each oRow2 In myRows2.Columns(1).Cells

        If oRow2.EntireRow.Hidden Then

            If rng2 Is Nothing Then

                Set rng2 = oRow2

            Else

                Set rng2 = Union(rng2, oRow2)

            End If

        End If

    Next

    If Not rng2 Is Nothing Then rng2.EntireRow.Delete

    ''''''''Delete hidden rows''''''''

    Else: Range("A2:I" & GLrow).EntireRow.Delete

    End If

    Application.Goto Range("A1"), True

   

    '''EXP

Line3:

    Application.DisplayAlerts = False

    Worksheets("EXP").Delete

    Application.DisplayAlerts = True

    Worksheets("GL").Copy After:=Worksheets("FA")

    ActiveSheet.Name = "EXP"

    With Worksheets("EXP").Tab

        .ThemeColor = xlThemeColorLight2

        .TintAndShade = 0.399975585192419

    End With

    If Worksheets("EXP").ListObjects.Count > 0 Then

    ActiveSheet.ListObjects(1).Unlist

    End If

    GLrow = Cells(Cells.Rows.Count, "F").End(xlUp).Row

    Range("A1:I" & GLrow).AutoFilter

    Range("A" & GLrow + 2).Value = "Grand Total"

    Range("H" & GLrow + 2).NumberFormat = "General"

    Range("H" & GLrow + 2).FormulaR1C1 = "=SUM(R[" & -GLrow & "]C:R[-1]C)"

    Range("A" & GLrow + 2 & ":H" & GLrow + 2).Select

    With Selection.Borders(xlEdgeTop)

        .LineStyle = xlContinuous

        .Weight = xlThin

    End With

    With Selection.Borders(xlEdgeBottom)

        .LineStyle = xlDouble

        .Weight = xlThick

    End With

    Range("A" & GLrow + 2 & ":H" & GLrow + 2).Font.Bold = True

   

    If Application.WorksheetFunction.CountIf(Range("B2:B" & GLrow), "EXP*") > 0 Then

    Range("A1:I" & GLrow).AutoFilter Field:=2, Criteria1:="EXP*"

    ''''''''Delete hidden rows''''''''

    Dim oRow3 As Range, rng3 As Range

    Dim myRows3 As Range

    With Sheets("EXP")

        Set myRows3 = Intersect(.Range("A:A").EntireRow, .UsedRange)

        If myRows3 Is Nothing Then GoTo Line4

    End With

    For Each oRow3 In myRows3.Columns(1).Cells

        If oRow3.EntireRow.Hidden Then

            If rng3 Is Nothing Then

                Set rng3 = oRow3

            Else

                Set rng3 = Union(rng3, oRow3)

            End If

        End If

    Next

    If Not rng3 Is Nothing Then rng3.EntireRow.Delete

    ''''''''Delete hidden rows''''''''

    Else: Range("A2:I" & GLrow).EntireRow.Delete

    End If

    Application.Goto Range("A1"), True

   

Line4:

    ''''BCA

    Worksheets("BCA").Range("G47").FormulaR1C1 = "=VLOOKUP(""Grand Total"",EXP!C[-6]:C[3],10,FALSE)"

    Worksheets("BCA").Range("G48").FormulaR1C1 = "=VLOOKUP(""Grand Total"",FA!C[-6]:C[1],8,FALSE)"

    ''''SRM

    Worksheets("SRM").Range("G47").FormulaR1C1 = "=VLOOKUP(""Grand Total"",EXP!C[-6]:C[3],10,FALSE)"

    Worksheets("SRM").Range("G48").FormulaR1C1 = "=VLOOKUP(""Grand Total"",FA!C[-6]:C[1],8,FALSE)"

   

    Application.Goto Worksheets("AP GSTAmount").Range("A1"), True

   

End Sub
