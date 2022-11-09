Attribute VB_Name = "ModReportes"
Option Explicit

Public gobjReporte As Object
'************************************************************************
'* Nombre : ShowMultiReport
'* Descripción : Muestra un Reporte con o sin SubReportes
'************************************************************************
Public Function ShowMultiReport( _
            irsDatos As Object, _
            istrNombreArchivo As String, _
            Optional arrParametros As Variant = Null, _
            Optional istrPath As String = "", _
            Optional istrTitulo As String = "", _
            Optional iblnParametroUsuario As Boolean = True, _
            Optional iblnPreview As Boolean = True, _
            Optional iblnPermitirVacio As Boolean = False, _
            Optional irsDatos1 As Object = Nothing, _
            Optional istrNombreArchivo1 As String = "", _
            Optional irsDatos2 As Object = Nothing, _
            Optional istrNombreArchivo2 As String = "", _
            Optional irsDatos3 As Object = Nothing, _
            Optional istrNombreArchivo3 As String = "", _
            Optional irsDatos4 As Object = Nothing, _
            Optional istrNombreArchivo4 As String = "", _
            Optional irsDatos5 As Object = Nothing, _
            Optional istrNombreArchivo5 As String = "", _
            Optional irsDatos6 As Object = Nothing, _
            Optional istrNombreArchivo6 As String = "" _
            ) As Boolean

On Error GoTo ShowMultiReportErr
    Dim RepApp As Object
    Dim CDOSet As Object
    Dim RepDb As Object
    Dim RepTables As Object
    Dim RepTable As Object
    Dim intLabelCount As Integer
    Dim objSubReporte1 As Object
    Dim RepApp2 As Object
    Dim CDOSet1 As Object
    Dim RepDb1 As Object
    Dim RepTables1 As Object
    Dim RepTable1 As Object
    Dim intLabelCount1 As Integer
    Dim objSubReporte2 As Object

    Dim CDOSet2 As Object
    Dim RepDb2 As Object
    Dim RepTables2 As Object
    Dim RepTable2 As Object
    Dim intLabelCount2 As Integer
    Dim objParametros As Object
    Dim objParametro As Object
    Dim RepApp1 As Object
    Dim strFN As String
    Dim i As Integer
    
    Dim CDOSet3 As Object
    Dim intLabelCount3 As Integer
    Dim RepDb3 As Object
    Dim objSubReporte3 As Object
    Dim RepTables3 As Object
    Dim RepTable3 As Object
    Dim RepApp3 As Object
    
    Dim CDOSet4 As Object
    Dim intLabelCount4 As Integer
    Dim RepDb4 As Object
    Dim objSubReporte4 As Object
    Dim RepTables4 As Object
    Dim RepTable4 As Object
    Dim RepApp4 As Object
    
    Dim CDOSet5 As Object
    Dim intLabelCount5 As Integer
    Dim RepDb5 As Object
    Dim objSubReporte5 As Object
    Dim RepTables5 As Object
    Dim RepTable5 As Object
    Dim RepApp5 As Object
    
    Dim CDOSet6 As Object
    Dim intLabelCount6 As Integer
    Dim RepDb6 As Object
    Dim objSubReporte6 As Object
    Dim RepTables6 As Object
    Dim RepTable6 As Object
    Dim RepApp6 As Object
    
    If istrTitulo = "" Then istrTitulo = istrNombreArchivo
    'If istrPath = "" Then istrPath = gstrPathReportes
    
    Set RepApp = CreateObject("Crystal.CRPE.Application")
    Set gobjReporte = RepApp.OpenReport(istrPath & istrNombreArchivo & ".rpt")
    Set CDOSet = CreateObject("CrystalDataObject.CrystalComObject")
    Set CDOSet1 = CreateObject("CrystalDataObject.CrystalComObject")
    Set CDOSet2 = CreateObject("CrystalDataObject.CrystalComObject")
    Set CDOSet3 = CreateObject("CrystalDataObject.CrystalComObject")
    Set CDOSet4 = CreateObject("CrystalDataObject.CrystalComObject")
    Set CDOSet5 = CreateObject("CrystalDataObject.CrystalComObject")
    Set CDOSet6 = CreateObject("CrystalDataObject.CrystalComObject")
    
    '...Ingresa estructura de datos del archivo ttx a la estructura CDOSet
        strFN = istrPath & istrNombreArchivo & ".ttx"
        intLabelCount = SetStructuredTTX(CDOSet, strFN)
    
    '...Ahora para el detalle 1 a CDOSet1
    If istrNombreArchivo1 <> "" Then
        strFN = istrPath & istrNombreArchivo1 & ".ttx"
        intLabelCount1 = SetStructuredTTX(CDOSet1, strFN)
    End If
    '...Ahora para el detalle 2 a CDOSet2
    If istrNombreArchivo2 <> "" Then
        strFN = istrPath & istrNombreArchivo2 & ".ttx"
        intLabelCount2 = SetStructuredTTX(CDOSet2, strFN)
    End If
    
    If istrNombreArchivo3 <> "" Then
        strFN = istrPath & istrNombreArchivo3 & ".ttx"
        intLabelCount3 = SetStructuredTTX(CDOSet3, strFN)
    End If
    
    If istrNombreArchivo4 <> "" Then
        strFN = istrPath & istrNombreArchivo4 & ".ttx"
        intLabelCount4 = SetStructuredTTX(CDOSet4, strFN)
    End If
    
    If istrNombreArchivo5 <> "" Then
        strFN = istrPath & istrNombreArchivo5 & ".ttx"
        intLabelCount5 = SetStructuredTTX(CDOSet5, strFN)
    End If
    
    If istrNombreArchivo6 <> "" Then
        strFN = istrPath & istrNombreArchivo6 & ".ttx"
        intLabelCount6 = SetStructuredTTX(CDOSet6, strFN)
    End If
    
    
    '...Ingreso de datos al objeto CDOSet
    If Not (InsertDataRS(CDOSet, irsDatos, intLabelCount) Or iblnPermitirVacio) Then
        Screen.MousePointer = vbDefault
        MsgBox "El reporte no tiene datos.", vbInformation, istrTitulo
    End If
    '...Asociando la data del objeto CDOSet como tabla del reporte
    Set RepDb = gobjReporte.Database
    Set RepTables = RepDb.Tables
    Set RepTable = RepTables(1)
    Call RepTable.SetPrivateData(3, CDOSet)
    gobjReporte.DiscardSavedData
    
    '...Existe detalle 1 ?
    If istrNombreArchivo1 <> "" Then
        '...Ingreso de datos al detalle 1
        InsertDataRS CDOSet1, irsDatos1, intLabelCount1
        '...Asociando la data del objeto CDOSet1 a la tabla del SubReporte 1
        'Set objSubReporte1 = gobjReporte.OpenSubreport("Detalle1")
        Set objSubReporte1 = gobjReporte.OpenSubreport(istrNombreArchivo1 & ".rpt") 'Ronald
        Set RepDb1 = objSubReporte1.Database
        Set RepTables1 = RepDb1.Tables
        Set RepTable1 = RepTables1(1)
        Call RepTable1.SetPrivateData(3, CDOSet1)
        objSubReporte1.DiscardSavedData
    End If
    
    '...Existe detalle 2 ?
    If istrNombreArchivo2 <> "" Then
        '...Ingreso de datos al detalle 2
        InsertDataRS CDOSet2, irsDatos2, intLabelCount2
        '...Asociando la data del objeto CDOSet1 a la tabla del SubReporte 2
        'Set objSubReporte2 = gobjReporte.OpenSubreport("Detalle2")
        Set objSubReporte2 = gobjReporte.OpenSubreport(istrNombreArchivo2 & ".rpt") 'Ronald.
        Set RepDb2 = objSubReporte2.Database
        Set RepTables2 = RepDb2.Tables
        Set RepTable2 = RepTables2(1)
        Call RepTable2.SetPrivateData(3, CDOSet2)
        objSubReporte2.DiscardSavedData
    End If
    
    
    If istrNombreArchivo3 <> "" Then
        '...Ingreso de datos al detalle 2
        InsertDataRS CDOSet3, irsDatos3, intLabelCount3
        '...Asociando la data del objeto CDOSet1 a la tabla del SubReporte 2
        'Set objSubReporte2 = gobjReporte.OpenSubreport("Detalle2")
        Set objSubReporte3 = gobjReporte.OpenSubreport(istrNombreArchivo3 & ".rpt") 'Ronald.
        Set RepDb3 = objSubReporte3.Database
        Set RepTables3 = RepDb3.Tables
        Set RepTable3 = RepTables3(1)
        Call RepTable3.SetPrivateData(3, CDOSet3)
        objSubReporte3.DiscardSavedData
    End If

    If istrNombreArchivo4 <> "" Then
        '...Ingreso de datos al detalle 2
        InsertDataRS CDOSet4, irsDatos4, intLabelCount4
        '...Asociando la data del objeto CDOSet1 a la tabla del SubReporte 2
        'Set objSubReporte2 = gobjReporte.OpenSubreport("Detalle2")
        Set objSubReporte4 = gobjReporte.OpenSubreport(istrNombreArchivo4 & ".rpt") 'Ronald.
        Set RepDb4 = objSubReporte4.Database
        Set RepTables4 = RepDb4.Tables
        Set RepTable4 = RepTables4(1)
        Call RepTable4.SetPrivateData(3, CDOSet4)
        objSubReporte4.DiscardSavedData
    End If
    
    If istrNombreArchivo5 <> "" Then
        '...Ingreso de datos al detalle 2
        InsertDataRS CDOSet5, irsDatos5, intLabelCount5
        '...Asociando la data del objeto CDOSet1 a la tabla del SubReporte 2
        'Set objSubReporte2 = gobjReporte.OpenSubreport("Detalle2")
        Set objSubReporte5 = gobjReporte.OpenSubreport(istrNombreArchivo5 & ".rpt") 'Ronald.
        Set RepDb5 = objSubReporte5.Database
        Set RepTables5 = RepDb5.Tables
        Set RepTable5 = RepTables5(1)
        Call RepTable5.SetPrivateData(3, CDOSet5)
        objSubReporte5.DiscardSavedData
    End If
    
    If istrNombreArchivo6 <> "" Then
        '...Ingreso de datos al detalle 2
        InsertDataRS CDOSet6, irsDatos6, intLabelCount6
        '...Asociando la data del objeto CDOSet1 a la tabla del SubReporte 2
        'Set objSubReporte2 = gobjReporte.OpenSubreport("Detalle2")
        Set objSubReporte6 = gobjReporte.OpenSubreport(istrNombreArchivo6 & ".rpt") 'Ronald.
        Set RepDb6 = objSubReporte6.Database
        Set RepTables6 = RepDb6.Tables
        Set RepTable6 = RepTables6(1)
        Call RepTable6.SetPrivateData(3, CDOSet6)
        objSubReporte6.DiscardSavedData
    End If
    
    
    '...Procesando los datos ingresados en el arreglo arrParametros
    Set objParametros = gobjReporte.ParameterFields

    For Each objParametro In objParametros  '...Para cada parámetro del Reporte
        If UCase(objParametro.ParameterFieldName) = "USUARIO" Then
            'objParametro.SetCurrentValue IIf(iblnParametroUsuario, gstrUsuario, "")
            objParametro.SetCurrentValue IIf(iblnParametroUsuario, 0, "")
        End If
        If Not IsNull(arrParametros) Then   'Si se ha pasado un arreglo de parámetros
            For i = LBound(arrParametros, 1) To UBound(arrParametros, 1)
                If objParametro.ParameterFieldName = CStr(arrParametros(i, 1)) Then
                    objParametro.SetCurrentValue CStr(arrParametros(i, 2))
                End If
            Next    '...Buscando el parámetro
        End If
    Next
    
   ' If iblnPreview Then '...Se muestra una presentación preliminar
        'gobjReporte.Preview istrTitulo, , , , , , frmRptListaHosp.hWnd
        'gobjReporte.Preview istrTitulo, , , , , , frmReporteMovXHosp.hWnd
        gobjReporte.Preview istrTitulo, , , , , , MDIFrmPrincipal.hwnd
        'gobjReporte.Preview istrTitulo, , , , , , MDIFrmPrincipal.hWnd
   ' Else     '...sino imprimir directamente
        'gobjReporte.Preview istrTitulo, , , , , , MDIFrmPrincipal.hWnd 'comentada antes
        'gobjReporte.PrintOut False, 1, False
   ' End If
   
    ShowMultiReport = True
    'Destruimos los objetos
    Set RepApp = Nothing
    Set CDOSet = Nothing
    Set RepDb = Nothing
    Set RepTables = Nothing
    Set RepTable = Nothing
    Set objParametros = Nothing
    Set objParametro = Nothing
    Set RepApp1 = Nothing
    Set CDOSet1 = Nothing
    Set RepDb1 = Nothing
    Set RepTables1 = Nothing
    Set RepTable1 = Nothing
    Set objSubReporte1 = Nothing
    Set RepApp2 = Nothing
    Set CDOSet2 = Nothing
    Set RepDb2 = Nothing
    Set RepTables2 = Nothing
    Set RepTable2 = Nothing
    Set objSubReporte2 = Nothing
    
    Set RepApp3 = Nothing
    Set CDOSet3 = Nothing
    Set RepDb3 = Nothing
    Set RepTables3 = Nothing
    Set RepTable3 = Nothing
    Set objSubReporte3 = Nothing
    
    Set RepApp4 = Nothing
    Set CDOSet4 = Nothing
    Set RepDb4 = Nothing
    Set RepTables4 = Nothing
    Set RepTable4 = Nothing
    Set objSubReporte4 = Nothing
    
    Set RepApp5 = Nothing
    Set CDOSet5 = Nothing
    Set RepDb5 = Nothing
    Set RepTables5 = Nothing
    Set RepTable5 = Nothing
    Set objSubReporte5 = Nothing
    
    Set RepApp6 = Nothing
    Set CDOSet6 = Nothing
    Set RepDb6 = Nothing
    Set RepTables6 = Nothing
    Set RepTable6 = Nothing
    Set objSubReporte6 = Nothing
    Exit Function
    
ShowMultiReportErr:
    Screen.MousePointer = vbDefault
    ShowMultiReport = False
    Set RepApp = Nothing
    Set CDOSet = Nothing
    Set RepDb = Nothing
    Set RepTables = Nothing
    Set RepTable = Nothing
    Set objParametros = Nothing
    Set objParametro = Nothing
    Set RepApp1 = Nothing
    Set CDOSet1 = Nothing
    Set RepDb1 = Nothing
    Set RepTables1 = Nothing
    Set RepTable1 = Nothing
    Set objSubReporte1 = Nothing
    Set RepApp2 = Nothing
    Set CDOSet2 = Nothing
    Set RepDb2 = Nothing
    Set RepTables2 = Nothing
    Set RepTable2 = Nothing
    Set objSubReporte2 = Nothing
    
    Set RepApp3 = Nothing
    Set CDOSet3 = Nothing
    Set RepDb3 = Nothing
    Set RepTables3 = Nothing
    Set RepTable3 = Nothing
    Set objSubReporte3 = Nothing

    Set RepApp4 = Nothing
    Set CDOSet4 = Nothing
    Set RepDb4 = Nothing
    Set RepTables4 = Nothing
    Set RepTable4 = Nothing
    Set objSubReporte4 = Nothing
    
    Set RepApp5 = Nothing
    Set CDOSet5 = Nothing
    Set RepDb5 = Nothing
    Set RepTables5 = Nothing
    Set RepTable5 = Nothing
    Set objSubReporte5 = Nothing
    
    Set RepApp6 = Nothing
    Set CDOSet6 = Nothing
    Set RepDb6 = Nothing
    Set RepTables6 = Nothing
    Set RepTable6 = Nothing
    Set objSubReporte6 = Nothing
    
    If Err.Number = 429 Then 'error del create object del crystal
        MsgBox "No tiene instalado correctamente el aplicativo Crystal Reports, no puede imprimirse o mostrarse una vista previa del reporte.", vbInformation, "Error de Instalación del Sistema"
    ElseIf Err.Number = 20545 Then 'error al ser cancelada la impresión por el usuario
        MsgBox "La impresión ha sido cancelada por el usuario.", vbInformation, "Impresión Cancelada"
    ElseIf Err.Number = 20507 Then 'error al no encontrar archivo
        MsgBox "No se ha encontrado un archivo necesario para la impresión.  Verifique que el sistema esté correctamente instalado.", vbInformation, "No se Encontró Archivo Necesario para Impresión"
    Else
        MsgBox Err.Number & " " & Err.Description, vbInformation, MSGERROR
        'Call ErrorMessage(MGeneral_ShowMultiReport, Err.Source & " MGeneral:ShowMultiReport", Err.Description)
    End If
End Function


'*************************** llenar structura de base de datos***********************

'************************************************************************
'* Nombre : SetStructuresTTX
'* Descripción : Configurar la Estructura de Datos de un Reporte
'************************************************************************
Function SetStructuredTTX(ByRef CDOSet As Object, ByVal istrFN As String) As Integer
On Error GoTo errHandler
Dim intFN As Integer
Dim strLine As String
Dim intLabelCount As Integer
    intFN = FreeFile
    Open istrFN For Input As intFN
    Do While Not EOF(intFN)
        Line Input #intFN, strLine
        If Len(strLine) <> 0 And Right(strLine, 2) <> "%%" Then
            CDOSet.AddField Split(strLine, vbTab)(0), vbString
            intLabelCount = intLabelCount + 1
        End If
    Loop
    Close intFN
    SetStructuredTTX = intLabelCount
    Exit Function
errHandler:
    SetStructuredTTX = 0
    MsgBox Err.Number & " " & Err.Description, vbInformation, MSGERROR
    'Call ErrorMessage(MGeneral_SetStructuredTTX, Err.Source & " MGeneral:SetStructuredTTX", Err.Description)
End Function


'**********************insertar data***********************************
'************************************************************************
'* Nombre : InsertDataRS
'* Descripción : Enlaza los Datos del Recordset al Reporte
'************************************************************************
Function InsertDataRS(ByRef CDOSet As Object, ByVal irsDatos As Object, ByVal intLabelCount As Integer) As Boolean
On Error GoTo InsertDataErr
Dim LabelRows() As Variant
Dim intX As Integer
Dim intC As Integer
' TODO DPB debe ser parametro iblnPermitirVacio
Dim iblnPermitirVacio As Boolean
    If irsDatos.RecordCount > 0 Or iblnPermitirVacio Then
        With irsDatos
            If Not (.EOF And .BOF) Then
                .MoveLast
                '... LabelRows contiene toda la data del recordset irsDatos
                ReDim LabelRows(.RecordCount - 1, intLabelCount - 1)
                .MoveFirst
                For intX = LBound(LabelRows) To UBound(LabelRows)
                    For intC = 0 To .Fields.Count - 1
                        LabelRows(intX, intC) = CStr(IIf(IsNull(.Fields(intC).Value), "", .Fields(intC).Value))
                    Next 'intC
                    .MoveNext
                Next 'intX
                '...Añade las filas de data del arreglo LabelRows al objeto CDOSet
                CDOSet.AddRows LabelRows
            End If
        End With
    End If
    InsertDataRS = True
    Exit Function
InsertDataErr:
    InsertDataRS = False
    MsgBox Err.Number & " " & Err.Description, vbInformation, MSGERROR
   ' Call ErrorMessage(MGeneral_InsertDataRS, Err.Source & "MGeneral:InsertDataRS", Err.Description)
End Function


