Attribute VB_Name = "Data_Export1"
'Option Explicit
'Public Sub SaveAsExcel(ByVal rs As ADODB.Recordset, ByVal filename _
' As String, Optional Ffmt As XlFileFormat = xlWorkbookNormal, _
' Optional bHeaders As Boolean = True)
' '***********************************************************
'
' ' Exports a Recordset data into a Microsoft Excel Sheet and
' ' then can save as new file
' ' with a given format such Lotus, Q-Pro, dBase, Text
' '
' ' Arguments:
' '
' ' rs : Recordset object (ADODB) containing data.
' ' filename: Name of the file.
' ' Ffmt: File Format the default value is the
'  'MS-Excel current version.
' ' bHeaders: If true the name of the fields will be inserted
' 'in the first row of each column.
' '
'
'Dim xlApp As Excel.Application
'Dim xlBook As Excel.Workbook
'Dim xlSheet As Excel.Worksheet
'
'Dim i As Integer
''Field object
'Dim fd As Field
'
''Cell count, the cells we can use
'Dim CellCnt As Integer
'
''File Extension Type
'Dim Fet As String
'
' Screen.MousePointer = vbHourglass
'' Assign object references to the variables. Use
'' Add methods to create new workbook and worksheet
'' objects.
'Set xlApp = New Excel.Application
'Set xlBook = xlApp.Workbooks.Add
'Set xlSheet = xlBook.Worksheets.Add
'
''Get the field names
'If bHeaders Then
'     CellCnt = 1
'     For Each fd In rs.Fields
''        Select Case fd.Type
''        Case dbBinary, dbGUID, dbLongBinary, dbVarBinary
''            ' This type of data can't export to excel
' '       Case Else
'            xlSheet.cells(1, CellCnt).Value = fd.Name
'            xlSheet.cells(1, CellCnt).Interior.ColorIndex = 33
'            xlSheet.cells(1, CellCnt).Font.Bold = True
'            xlSheet.cells(1, CellCnt).BorderAround xlContinuous
'            CellCnt = CellCnt + 1
''        End Select
'     Next
'End If
'
''Rewind the rescordset
'rs.MoveFirst
'i = 2
'Do While Not rs.EOF()
'     CellCnt = 1
'     For Each fd In rs.Fields
'    '    Select Case fd.Type
'     '   Case dbBinary, dbGUID, dbLongBinary, dbVarBinary
'            ' This type of data can't export to excel
'      '  Case Else
'            xlSheet.cells(i, CellCnt).Value = _
'                rs.Fields(fd.Name).Value
'            'xlSheet.Columns().AutoFit
'            CellCnt = CellCnt + 1
'       ' End Select
'     Next
'     rs.MoveNext
'     i = i + 1
' Loop
'
''Fit all columns
'CellCnt = 1
'For Each fd In rs.Fields
'
'    ' Select Case fd.Type
'        ' Case dbBinary, dbGUID, dbLongBinary, _
'         '        dbVarBinary
'                  ' This type of data can't export to excel
'          'Case Else
'                  xlSheet.Columns(CellCnt).AutoFit
'                  CellCnt = CellCnt + 1
'          'End Select
'Next
'
''Get the file extension
'''Select Case Ffmt
'''     Case xlSYLK
'''         Fet = "slk"
'''     Case xlWKS
'''         Fet = "wks"
'''     Case xlWK1, xlWK1ALL, xlWK1FMT
'''         Fet = "wk1"
'''     Case xlCSV, xlCSVMac, xlCSVWindows 'xlCSVdos,
'''         Fet = "csv"
'''     Case xlDBF2, xlDBF3, xlDBF4
'''         Fet = "dbf"
'''     Case xlWorkbookNormal, xlExcel2FarEast, xlExcel3, _
'''         xlExcel4, xlExcel4Workbook, xlExcel5, xlExcel6, _
'''         xlExcel7, xlExcel9795
'''         Fet = "xls"
'''     Case xlHtml
'''         Fet = "htm"
'''     Case xlTextMac, xlTextdos, xlTextWindows, xlUnicodeText, _
'''           xlCurrentPlatformText
'''         Fet = "txt"
'''     Case xlTextPrinter
'''         Fet = "prn"
'''     Case Else
'''         Fet = "dat"
''' End Select
'''
'' Save the Worksheet.
'If InStr(1, filename, ".") = 0 Then filename = _
'   filename + "." + "xls"
'xlSheet.SaveAs filename, Ffmt
'
'' Close the Workbook
'xlBook.Close
'' Close Microsoft Excel with the Quit method.
'xlApp.Quit
'
'' Release the objects.
'Set xlApp = Nothing
'Set xlBook = Nothing
'Set xlSheet = Nothing
'MsgBox "Search Result saved successfully as " + filename, vbExclamation, "Export Successful"
'Screen.MousePointer = vbDefault
'End Sub
'
'
'
'
'
