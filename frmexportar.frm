VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmexportar 
   BorderStyle     =   0  'None
   Caption         =   "IMPORTACION DE DATA"
   ClientHeight    =   6765
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13860
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   13860
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "EXPORTARMOVIMIENTOS"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4560
      TabIndex        =   4
      Top             =   720
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "IMPORTAR MOVIMIENTOS"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   4335
   End
   Begin VB.CommandButton cmdExportar 
      Caption         =   "EXPORTAR PLAN CONTABLE"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4560
      TabIndex        =   2
      Top             =   120
      Width           =   2655
   End
   Begin MSComDlg.CommonDialog dlgGuardar 
      Left            =   9600
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdImportar 
      Caption         =   "IMPORTAR PLAN CONTABLE"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   11280
      OleObjectBlob   =   "frmexportar.frx":0000
      Top             =   240
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid hfproductos 
      Height          =   5175
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   9128
      _Version        =   393216
      ForeColor       =   8388608
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      ForeColorFixed  =   8388608
      BackColorBkg    =   16777215
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "frmexportar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExportar_Click()
Dim sPathDB        As String
Dim Consulta    As String
Dim RutaArchivo As String


RutaArchivo = SelArchivo
If RutaArchivo = vbNullString Then
        Exit Sub
End If
    
    
    
    Dim codigo As Integer
strCadena = "DELETE FROM importacion_plan_contable"
CnBd.Execute (strCadena)
'On Error GoTo salir
For i = 0 To Me.hfproductos.Rows - 1
      
        If i < Me.hfproductos.Rows Then
            in_cuenta = Val(Format(Trim(Me.hfproductos.TextMatrix(i, 0)), "00000"))
            in_descripcion = Replace(UCase(Trim(Me.hfproductos.TextMatrix(i, 1))), "'", " ")
            in_nivel = UCase(Trim(Me.hfproductos.TextMatrix(i, 2)))
            ctaref = UCase(Trim(Me.hfproductos.TextMatrix(i, 3)))
            ctaref2 = UCase(Trim(Me.hfproductos.TextMatrix(i, 4)))
            ctapte = UCase(Trim(Me.hfproductos.TextMatrix(i, 5)))
            porcr1 = UCase(Trim(Me.hfproductos.TextMatrix(i, 6)))
            porcr2 = ""
     
        'in_precio_costo = Format(Me.hfproductos.TextMatrix(i, 7), "###0.00")
        
        
        strCadena = "INSERT INTO importacion_plan_contable(`id_cuenta`,`descripcion`,`nivel`,`ctaref`,`ctaref2`,`ctapte`,`porcr1`,`porcr2`)VALUES" & _
        "('" & in_cuenta & "','" & in_descripcion & "','" & in_nivel & "','" & ctaref & "','" & ctaref2 & "','" & ctapte & "','" & porcr1 & "','" & porcr2 & "')"
        CnBd.Execute (strCadena)
    End If
Next i
    
    
    
    If Exportar_ADO_Excel(sPathDB, Consulta, RutaArchivo, "importacion_plan_contable") Then
       MsgBox "Exportacion Correcta", vbInformation
       'Call AbreArchivo(RutaArchivo)
    End If
End Sub

Private Sub cmdImportar_Click()
Dim sPathDB        As String
Dim Consulta    As String
Dim RutaArchivo As String

Dim Archivo As String
Archivo = Trim("LEmigracion.xls")
      'Dim obj As New get_excel
      Set Me.hfproductos.DataSource = Leer_Excel(App.Path & "\comparar_percy\" & Archivo, "Sheet1")
     
      'Set obj = Nothing
      
      
      
Exit Sub



End Sub


Public Function Leer_Excel(ByVal PathXls As String, Hoja As String) As ADODB.Recordset

      On Error GoTo ErrorFunction
      Dim rs As ADODB.Recordset
      Set rs = New ADODB.Recordset
      Dim cs As String
    
      rs.CursorLocation = adUseClient
      rs.CursorType = adOpenKeyset
      rs.LockType = adLockBatchOptimistic

      cs = "DRIVER=Microsoft Excel Driver (*.xls);" & "DBQ=" & PathXls
      
      Hoja = "[" & Hoja & "$" & "]"
      
      rs.Open "SELECT * FROM " & Hoja, cs
      Set Leer_Excel = rs
      Set rs = Nothing
      Exit Function
ErrorFunction:
      MsgBox Err.Description, vbCritical
      Err.Clear
End Function

'devuelve un objeto Recordset con los datos del txt

Public Function LeerTxt(Directorio As String) As ADODB.Recordset
      On Error GoTo ErrorFunction
      Dim rs As ADODB.Recordset
      Set rs = New ADODB.Recordset
      Dim cn As ADODB.Connection
      Set cn = New ADODB.Connection
      
      cn.Open "DRIVER={Microsoft Text Driver (*.txt; *.csv)};" & _
                         "DBQ=" & Directorio & ";", "", ""
      rs.Open "select * from [archivo#txt]", cn, adOpenStatic, adLockReadOnly, adCmdText
      Set LeerTxt = rs
      
      Set rs = Nothing
      Set cn = Nothing
      
      Exit Function
ErrorFunction:
      MsgBox Err.Description, vbCritical
      Err.Clear
End Function






Private Function SelArchivo() As String
    On Error Resume Next
    
    dlgGuardar.ShowSave
    If Err.Number = cdlCancel Then
        ' The user canceled.
        Exit Function
    ElseIf Err.Number <> 0 Then
        ' Unknown error.
        MsgBox "Error " & Format$(Err.Number) & _
            " seleccionando ruta." & vbCrLf & _
            Err.Description
        Exit Function
    End If
    SelArchivo = dlgGuardar.FileName
End Function

Private Sub Command1_Click()
Dim sPathDB        As String
Dim Consulta    As String
Dim RutaArchivo As String

Dim Archivo As String

strCadena = "DELETE FROM importacion_asiento"
CnBd.Execute (strCadena)

Call get_txt
Exit Sub
Archivo = Trim("LEDATA.xls")
      'Dim obj As New get_excel
      Set Me.hfproductos.DataSource = Leer_Excel(App.Path & "\comparar_percy\" & Archivo, "Sheet1")
     
      'Set obj = Nothing
      
      
      
Exit Sub

End Sub
Private Sub get_txt()
Dim registro() As String
strRuta_ini = App.Path & "\comparar_percy\LEDATA.txt"
'strRuta_ini = App.Path & "\vitekey.ini"
FileName = Dir(strRuta_ini)
If FileName = "" Then
    strServer_ini = "158.69.74.64"
    Open strRuta_ini For Output As #1
    Print #1, strServer_ini
    sys_Server = strServer_ini
    Close #1
Else
    'Open strRuta_ini For Input As #1
    'Line Input #1, sys_Server
    'sys_Server = sys_Server
    'Close #1
    
    fnum = FreeFile
    Open strRuta_ini For Input As fnum
    i = 0
    Do While Not EOF(fnum)
        'registro = Split()
        Line Input #fnum, file_line
        registro = Split(file_line, "|")
        
            in_movimiento = registro(1)
            ano_as = Year(Format(registro(12), "dd-mm-YYYY"))
            mes_as = Format(Month(Format(registro(12), "dd-mm-YYYY")), "00")
            dia_as = Format(Day(Format(registro(12), "dd-mm-YYYY")), "00")
            compro = Format(registro(11), "000000")
            in_cuenta = Trim(registro(3))
            coddoc = Trim(registro(9))
            nrodoc = Trim(registro(10)) & "-" & Trim(registro(11))
            docref = coddoc
            fechao = Format(registro(12), "YYYY-mm-dd")
            fvenci = Format(registro(14), "YYYY-mm-dd")
            nomref = ""
            rucref = ""
            glosa = Trim(registro(15))
            If Val(registro(17)) > 0 Then
                tmovim = "D"
            Else
                tmovim = "H"
            End If
            c1 = registro(2)
            debe = Val(registro(17))
            haber = Val(registro(18))
            moneda = "S"
            tipcam = 0
            inaigv = "S"
            idcompro = registro(1)
            fpago = Format(registro(12), "dd-mm-YYYY")
            obser = ""
            detfec = "1900-01-01"
            detnro = ""
            detimp = 0
            dettas = "12"
            crefe = ""
            nrefe = ""
            frefe = "1900-01-01"
            
        
        strCadena = "INSERT INTO importacion_asiento " & _
        "(id_movimiento,c1,`ano_as`,`mes_as`,`dia_as`,`compro`,`origen`,`cuenta`,`coddoc`,`nrodoc`,`docref`,`fechao`,`fvenci`,`nomref`,`rucref`," & _
        "`glosa`,`tmovim`,`debe`,`haber`,`moneda`,`tipcam`,`inaigv`,`idcompro`,`fpago`,`obser`,`detfec`,`detnro`,`detimp`,`dettas`," & _
        "`crefe`,`nrefe`,`frefe`)VALUES" & _
        "('" & in_movimiento & "','" & c1 & "','" & ano_as & "','" & mes_as & "','" & dia_as & "','" & compro & "',' ','" & in_cuenta & "'," & _
        "'" & coddoc & "','" & nrodoc & "','" & docref & "','" & Format(fechao, "YYYY-mm-dd") & "','" & Format(fvenci, "YYYY-mm-dd") & "','" & nomref & "','" & rucref & "'," & _
        "'" & glosa & "','" & tmovim & "','" & debe & "','" & haber & "','" & moneda & "','" & tipcam & "','" & inaigv & "'," & _
        "'" & idcompro & "','" & Format(fpago, "YYYY-mm-dd") & "','" & obser & "','" & Format(detfec, "YYYY-mm-dd") & "','" & detnro & "','" & detimp & "','" & dettas & "'," & _
        "'" & crefe & "','" & nrefe & "','" & Format(frefe, "YYYY-mm-dd") & "')"
        CnBd.Execute (strCadena)
       
        i = i + 1
        DoEvents
        Command1.Caption = Str(i)
    Loop
    Close #fnum
    
    End If
End Sub

Private Sub get_txt_registro_venta()
Dim registro() As String
strRuta_ini = App.Path & "\comparar_percy\LEDATAVENTA.txt"
'strRuta_ini = App.Path & "\vitekey.ini"
FileName = Dir(strRuta_ini)
If FileName = "" Then
    strServer_ini = "158.69.74.64"
    Open strRuta_ini For Output As #1
    Print #1, strServer_ini
    sys_Server = strServer_ini
    Close #1
Else
    'Open strRuta_ini For Input As #1
    'Line Input #1, sys_Server
    'sys_Server = sys_Server
    'Close #1
    
    fnum = FreeFile
    Open strRuta_ini For Input As fnum
    i = 0
    Do While Not EOF(fnum)
        'registro = Split()
        Line Input #fnum, file_line
        registro = Split(file_line, "|")
        
            in_movimiento = registro(1)
            ano_as = Year(Format(registro(12), "dd-mm-YYYY"))
            mes_as = Format(Month(Format(registro(12), "dd-mm-YYYY")), "00")
            dia_as = Format(Day(Format(registro(12), "dd-mm-YYYY")), "00")
            compro = Format(registro(11), "000000")
            in_cuenta = Trim(registro(3))
            coddoc = Trim(registro(9))
            nrodoc = Trim(registro(10)) & "-" & Trim(registro(11))
            docref = coddoc
            fechao = Format(registro(12), "YYYY-mm-dd")
            fvenci = Format(registro(14), "YYYY-mm-dd")
            nomref = ""
            rucref = ""
            glosa = Trim(registro(15))
            If Val(registro(17)) > 0 Then
                tmovim = "D"
            Else
                tmovim = "H"
            End If
            
            debe = Val(registro(17))
            haber = Val(registro(18))
            moneda = "S"
            tipcam = 0
            inaigv = "S"
            idcompro = registro(1)
            fpago = Format(registro(12), "dd-mm-YYYY")
            obser = ""
            detfec = ""
            detnro = ""
            detimp = 0
            dettas = "12"
            crefe = ""
            nrefe = ""
            frefe = 0
            
        
        strCadena = "INSERT INTO importacion_asiento " & _
        "(id_movimiento,`ano_as`,`mes_as`,`dia_as`,`compro`,`origen`,`cuenta`,`coddoc`,`nrodoc`,`docref`,`fechao`,`fvenci`,`nomref`,`rucref`," & _
        "`glosa`,`tmovim`,`debe`,`haber`,`moneda`,`tipcam`,`inaigv`,`idcompro`,`fpago`,`obser`,`detfec`,`detnro`,`detimp`,`dettas`," & _
        "`crefe`,`nrefe`,`frefe`)VALUES" & _
        "('" & in_movimiento & "','" & ano_as & "','" & mes_as & "','" & dia_as & "','" & compro & "',' ','" & in_cuenta & "'," & _
        "'" & coddoc & "','" & nrodoc & "','" & docref & "','" & Format(fechao, "YYYY-mm-dd") & "','" & Format(fvenci, "YYYY-mm-dd") & "','" & nomref & "','" & rucref & "'," & _
        "'" & glosa & "','" & tmovim & "','" & debe & "','" & haber & "','" & moneda & "','" & tipcam & "','" & inaigv & "'," & _
        "'" & idcompro & "','" & Format(fpago, "YYYY-mm-dd") & "','" & obser & "','" & detfec & "','" & detnro & "','" & detimp & "','" & dettas & "'," & _
        "'" & crefe & "','" & nrefe & "','" & frefe & "')"
        CnBd.Execute (strCadena)
       
        i = i + 1
    Loop
    Close #fnum
    
    End If
End Sub

Private Sub Command2_Click()
Dim sPathDB        As String
Dim Consulta    As String
Dim RutaArchivo As String
Dim in_cuenta As String
Dim in_movimiento As Double
RutaArchivo = SelArchivo
If RutaArchivo = vbNullString Then
        Exit Sub
End If
    
Dim codigo As Integer

strCadena = "SELECT * FROM  importacion_asiento ORDER BY id ASC "
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
For i = 0 To rst.RecordCount - 1
      
            
            in_cuenta = Mid(rst("cuenta"), 1, 3)
            If (in_cuenta = "1212" And rst("c1") = "M1") Or (in_cuenta = "401" And rst("c1") = "M2") Or (in_cuenta = "701" And rst("c1") = "M3") Then
                origen = "03" ' VENTAS
            End If
             If (in_cuenta = "101" And rst("cuenta") = "M1") Or (in_cuenta = "1212" And rst("c1") = "M2") Then
                origen = "05" ' CAJA
            End If
             
             If (in_cuenta = "104" And rst("c1") = "M2") Or (in_cuenta = "641" And rst("c1") = "M1") Then
                origen = "06" ' BANCO
            End If
            
             If (in_cuenta = "40113" And rst("c1") = "M1") Or (in_cuenta = "4212" And rst("c1") = "M1") Or (in_cuenta = "104135" And rst("c1") = "M3") Then
                origen = "06" ' BANCO
            End If
            
            
           strCadena = "UPDATE importacion_asiento SET origen='" & origen & "' WHERE id='" & rst("id") & "'"
           CnBd.Execute (strCadena)
     
           rst.MoveNext
           DoEvents
           Me.Command2.Caption = Str(i)
        
       
     
    
Next i
    
End If
    
    If Exportar_ADO_Excel(sPathDB, Consulta, RutaArchivo, "importacion_asiento") Then
       MsgBox "Exportacion Correcta", vbInformation
       'Call AbreArchivo(RutaArchivo)
    End If
End Sub

Private Sub Form_Load()

Skin1.LoadSkin App.Path & "\Skins\BS.skn"
Skin1.ApplySkin Me.hWnd
CenterForm Me



dlgGuardar.InitDir = "%HOMEDRIVE%"
    dlgGuardar.Filter = "Archivo de Excel (*.xlsx)|*.xlsx"
    dlgGuardar.FilterIndex = 1
    dlgGuardar.DialogTitle = "Guardar reporte como..."
    
    

End Sub
Private Function Exportar_ADO_Excel(sPathDB As String, Sql As String, sOutputPathXLS As String, ByVal in_tabla As String) As Boolean
      
      
    On Error GoTo errSub
      
    Dim cn          As New ADODB.Connection
   
    Dim Excel       As Object
    Dim Libro       As Object
    Dim Hoja        As Object
    Dim arrData     As Variant
    Dim iRec        As Long
    Dim iCol        As Integer
    Dim iRow        As Integer
      
    Me.Enabled = False
    
    
    
    
    
    
    
   
    
    
    
    
    
    
    
      
  strCadena = "SELECT * FROM " & in_tabla & ""
  Call ConfiguraRst(strCadena)
    ' -- Crear los objetos para utilizar el Excel
    Set Excel = CreateObject("Excel.Application")
    Set Libro = Excel.Workbooks.Add
      
    ' -- Hacer referencia a la hoja
    Set Hoja = Libro.Worksheets(1)
      
    Excel.Visible = True: Excel.UserControl = True
    iCol = rst.Fields.Count
    For iCol = 1 To rst.Fields.Count
        Hoja.Cells(1, iCol).Value = rst.Fields(iCol - 1).Name
    Next
      
    If Val(Mid(Excel.Version, 1, InStr(1, Excel.Version, ".") - 1)) > 8 Then
        Hoja.Cells(2, 1).CopyFromRecordset rst
    Else
  
        arrData = rst.GetRows
  
        iRec = UBound(arrData, 2) + 1
          
        For iCol = 0 To rst.Fields.Count - 1
            For iRow = 0 To iRec - 1
  
                If IsDate(arrData(iCol, iRow)) Then
                    arrData(iCol, iRow) = Format(arrData(iCol, iRow))
  
                ElseIf IsArray(arrData(iCol, iRow)) Then
                    arrData(iCol, iRow) = "Array Field"
                End If
            Next iRow
        Next iCol
              
        ' -- Traspasa los datos a la hoja de Excel
        Hoja.Cells(2, 1).Resize(iRec, rst.Fields.Count).Value = GetData(arrData)
    End If
  
    Excel.Selection.CurrentRegion.Columns.AutoFit
    Excel.Selection.CurrentRegion.Rows.AutoFit
  
    ' -- Cierra el recordset y la base de datos y los objetos ADO

      
   
    ' -- guardar el libro
    Libro.saveAs sOutputPathXLS
    Libro.Close
    ' -- Elimina las referencias Xls
    Set Hoja = Nothing
    Set Libro = Nothing
    Excel.quit
    Set Excel = Nothing
      
    Exportar_ADO_Excel = True
    Me.Enabled = True
    Exit Function
errSub:
    MsgBox Err.Description, vbCritical, "Error"
    Exportar_ADO_Excel = False
   
End Function
  
Private Function GetData(vValue As Variant) As Variant
    Dim x As Long, y As Long, xMax As Long, yMax As Long, T As Variant
      
    xMax = UBound(vValue, 2): yMax = UBound(vValue, 1)
      
    ReDim T(xMax, yMax)
    For x = 0 To xMax
        For y = 0 To yMax
            T(x, y) = vValue(y, x)
        Next y
    Next x
      
    GetData = T
End Function
