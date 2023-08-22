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

   

    

    Dim Rng As Range

    Set Rng = Range("A1").CurrentRegion

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

    Rng.AutoFilter Field:=11, Criteria1:=Array("<>" & RemainStr), Operator:=xlAnd

    Rng.Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete

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

    Rng.AutoFilter Field:=23, Criteria1:=Array("<>EG*"), Operator:=xlAnd

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

   

    Dim Rng As Range

    Set Rng = Range("A1").CurrentRegion

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

    Rng.AutoFilter Field:=11, Criteria1:=Array("<>" & RemainStr), Operator:=xlAnd

    Rng.Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete

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

    Set Rng = Range("A6:A" & Prow).SpecialCells(xlCellTypeVisible)

    For Each cell In Rng

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

    Set Rng = Range("A2:A" & Wrow3).SpecialCells(xlCellTypeVisible)

    For Each cell In Rng

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

    Dim Rng As Range

    Set Rng = Range("A1").CurrentRegion

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

    Rng.AutoFilter Field:=10, Criteria1:=Array("<>" & RemainStr), Operator:=xlAnd

    Rng.Offset(1, 0).SpecialCells(xlCellTypeVisible).EntireRow.Delete

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
