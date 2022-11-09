Attribute VB_Name = "Crea_Planilla_Excel"
Public objExcel As Excel.Application
Public Numero_Registros2 As Long
Public Function Inicio_Excel() As Boolean
Dim i As Integer
Dim j As Integer
Set objExcel = New Excel.Application
objExcel.Visible = True 'lo hacemos visible
objExcel.SheetsInNewWorkbook = 1 'decimos cuantas hojas queremos en el nuevo documento
objExcel.Workbooks.Add ' añadimos el objeto al workbook
End Function
Public Function Formato_Excel(Num_Campos As Integer, Nombre_Campos() As String) As Boolean
With objExcel.ActiveSheet
         objExcel.Cells.Font.Size = 8
         objExcel.Cells.Font.name = "Draft 17cpi"

        .Range(.Cells(1, 1), .Cells(1, 8)).Borders.LineStyle = xlContinuous
        .Range(.Cells(1, 1), .Cells(1, 8)).Borders.LineStyle = xlContinuous
        .Cells(1, 1).Interior.ColorIndex = 15
        .Cells(1, 2).Interior.ColorIndex = 15
        .Cells(1, 3).Interior.ColorIndex = 15
        .Cells(1, 4).Interior.ColorIndex = 15
        .Cells(1, 5).Interior.ColorIndex = 15
        .Cells(1, 6).Interior.ColorIndex = 15
        .Cells(1, 7).Interior.ColorIndex = 15
        .Cells(1, 8).Interior.ColorIndex = 15
        .Range(.Cells(1, 1), .Cells(1, 8)).Font.color = vbBlue
    For i = 1 To Num_Campos Step 1
        .Cells(1, i) = Nombre_Campos(i)
    Next i
        .Columns("A").columnWidth = 4
        .Columns("B").columnWidth = 6
        .Columns("C").columnWidth = 17.5
        .Columns("D").columnWidth = 18.3
        .Columns("E").columnWidth = 17.7
        .Columns("F").columnWidth = 8
        .Columns("G").columnWidth = 10
        .Columns("H").columnWidth = 8
End With
End Function

