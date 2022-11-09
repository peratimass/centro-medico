VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmRegistroSunat 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "LIBROS ELECTRONICOS"
   ClientHeight    =   2565
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5850
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   5850
   ShowInTaskbar   =   0   'False
   Begin VitekeySoft.ChameleonBtn cmdGenrarLibro 
      Height          =   615
      Left            =   480
      TabIndex        =   4
      Top             =   1800
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   1085
      BTYPE           =   3
      TX              =   "GENERAR LIBRO ELECTRONICO"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmRegistroSunat.frx":0000
      PICN            =   "FrmRegistroSunat.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   645
      TabIndex        =   0
      Top             =   240
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
   End
   Begin VitekeySoft.ChameleonBtn cmdcerrar 
      Height          =   375
      Left            =   5280
      TabIndex        =   5
      Top             =   120
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   14215660
      BCOLO           =   14215660
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmRegistroSunat.frx":336A
      PICN            =   "FrmRegistroSunat.frx":3386
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lbloperacion 
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      Caption         =   "OPERACION REALIZADA CON EXITO !!"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   3795
   End
   Begin VB.Label lblNombreArchivo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " "
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   2400
      TabIndex        =   2
      Top             =   1320
      Width           =   3285
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOMBRE ARCHIVO :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   1995
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   2565
      Left            =   0
      Top             =   0
      Width           =   5850
   End
End
Attribute VB_Name = "FrmRegistroSunat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede
Private Sub cmdGenrarLibro_Click()
Dim strPle As String
Dim carpeta As String


If FrmRegistroCompras.Procedencia = nuevo Then
    carpeta = Trim(Trim(FrmRegistroCompras.HfdPersona.TextMatrix(FrmRegistroCompras.HfdPersona.Row, 1)) & "-" & Trim(FrmRegistroCompras.HfdPersona.TextMatrix(FrmRegistroCompras.HfdPersona.Row, 3)))
    strPle = App.Path & "\ple\" & carpeta
    If VerificarFichero(strPle) = False Then
       Call MkDir(App.Path & "\ple\" & carpeta)
    End If
 
    
    Call libro_compras(FrmRegistroCompras.HfdPersona.TextMatrix(FrmRegistroCompras.HfdPersona.Row, 1), FrmRegistroCompras.HfdPersona.TextMatrix(FrmRegistroCompras.HfdPersona.Row, 3), FrmRegistroCompras.HfdPersona.TextMatrix(FrmRegistroCompras.HfdPersona.Row, 0))
    Procedencia = Neutro
    Exit Sub
End If

If FrmRegistroVentas.Procedencia = nuevo Then
    carpeta = Trim(Trim(FrmRegistroVentas.HfdPersona.TextMatrix(FrmRegistroVentas.HfdPersona.Row, 1)) & "-" & Trim(FrmRegistroVentas.HfdPersona.TextMatrix(FrmRegistroVentas.HfdPersona.Row, 3)))
    strPle = App.Path & "\ple\" & carpeta
    If VerificarFichero(strPle) = False Then
       Call MkDir(App.Path & "\ple\" & carpeta)
    End If
    Call libro_ventas(FrmRegistroVentas.HfdPersona.TextMatrix(FrmRegistroVentas.HfdPersona.Row, 1), FrmRegistroVentas.HfdPersona.TextMatrix(FrmRegistroVentas.HfdPersona.Row, 3), FrmRegistroVentas.HfdPersona.TextMatrix(FrmRegistroVentas.HfdPersona.Row, 0))
    Procedencia = Neutro
    Exit Sub
End If


End Sub
Private Sub libro_ventas(ByVal id_mes As String, ByVal id_anio As String, ByVal ruc As String)
Dim Nombre As String
Dim Archivo As String
     
carpeta = Trim(id_mes & "-" & id_anio)

SqlDatos = "SELECT * FROM movimiento_venta WHERE ruc='" & ruc & "' AND month(fecha_emision)='" & id_mes & "' AND id_anio='" & id_anio & "' AND (id_doc='0001' or id_doc='0003' or id_doc='0007')  ORDER BY fecha_emision,id_doc,serie,numero ASC"
Call ConfiguraRst(SqlDatos)
'demo:     LE2053151604520130100080100001111.txt
'X = "140100001111"
Archivo = "LE" & ruc & id_anio & id_mes & "00" & "140100001111"
                                                 
'archivo = Trim("LE" & d_ruc & FrmRegistroCompras.HfdPersona.TextMatrix(FrmRegistroCompras.HfdPersona.Row, 3) & FrmRegistroCompras.HfdPersona.TextMatrix(FrmRegistroCompras.HfdPersona.Row, 1) & "0008010001OIM1")
ruta = Trim(App.Path & "\ple\" & carpeta & "\" & Archivo & ".txt")
Open ruta For Append As #1
Dim periodo As Date
periodo = Format(CVDate("01" & "-" & id_mes & "-" & id_anio), "YYYY-mm-dd")
Me.ProgressBar1.Min = 0
Me.ProgressBar1.Max = rst.RecordCount - 1
For i = 0 To rst.RecordCount - 1
    Me.ProgressBar1.Value = i
    
    campo1 = rst("id_anio") & rst("id_mes") & "00"
    'campo1 = "20130200"
    If DateDiff("m", rst("fecha_emision"), periodo) = 0 Then
        campo27 = "1"
    End If
    If rst("anulado") = "si" Then
        campo27 = "2"
    End If
    If DateDiff("m", rst("fecha_emision"), periodo) > 0 Then
        campo27 = "8"
    End If
    
    
    campo2 = formato_item(i + 1, 8)
    Campo3 = Format(rst("fecha_emision"), "dd/mm/YYYY")
    
    If rst("id_doc") = "0014" Then
        Campo4 = Format(rst("fecha_cancelacion"), "dd/mm/YYYY")
    Else
        Campo4 = ""
    End If
    
    campo5 = formato_item(rst("id_doc"), 2)
    
    
    If rst("id_doc") = "0001" Or rst("id_doc") = "0003" Or rst("id_doc") = "0004" Or rst("id_doc") = "0007" Or rst("id_doc") = "0008" Then
        Campo6 = formato_item(rst("serie"), 4)
     Else
        Campo6 = rst("serie")
    End If
    
    campo7 = rst("numero")
    campo8 = "0"
        
     campo9 = "0"
    
    If rst("id_doc") = "0000" Or rst("id_doc") = "0003" Or rst("id_doc") = "0005" Or rst("id_doc") = "0006" Or rst("id_doc") = "0007" Or rst("id_doc") = "0008" Or rst("id_doc") = "0011" Or rst("id_doc") = "0012" Or rst("id_doc") = "0013" Or rst("id_doc") = "0014" Or rst("id_doc") = "0015" Or rst("id_doc") = "0016" Or rst("id_doc") = "0018" Or rst("id_doc") = "0019" Or rst("id_doc") = "0022" Or rst("id_doc") = "0023" Or rst("id_doc") = "0026" Or rst("id_doc") = "0028" Or rst("id_doc") = "0030" Or rst("id_doc") = "0034" Or rst("id_doc") = "0035" Or rst("id_doc") = "0036" Or rst("id_doc") = "0037" Or rst("id_doc") = "0055" Or rst("id_doc") = "0056" Or rst("id_doc") = "0087" Or rst("id_doc") = "0088" Then
        If Len(rst("id_cliente")) = 8 Then
            campo9 = 1
            GoTo saltardoc
        End If
        If Len(rst("id_cliente")) = 11 Then
            campo9 = 6
            GoTo saltardoc
        End If
        campo9 = "0"
    Else
        If Len(rst("id_cliente")) = 8 Then
            campo9 = 1
            GoTo saltardoc
        End If
        If Len(rst("id_cliente")) = 11 Then
            campo9 = 6
            GoTo saltardoc
        End If
    End If
    
saltardoc:
   If rst("id_doc") = "0000" Or rst("id_doc") = "0003" Or rst("id_doc") = "0005" Or rst("id_doc") = "0006" Or rst("id_doc") = "0007" Or rst("id_doc") = "0008" Or rst("id_doc") = "0011" Or rst("id_doc") = "0012" Or rst("id_doc") = "0013" Or rst("id_doc") = "0014" Or rst("id_doc") = "0015" Or rst("id_doc") = "0016" Or rst("id_doc") = "0018" Or rst("id_doc") = "0019" Or rst("id_doc") = "0022" Or rst("id_doc") = "0023" Or rst("id_doc") = "0026" Or rst("id_doc") = "0028" Or rst("id_doc") = "0030" Or rst("id_doc") = "0034" Or rst("id_doc") = "0035" Or rst("id_doc") = "0036" Or rst("id_doc") = "0037" Or rst("id_doc") = "0055" Or rst("id_doc") = "0056" Or rst("id_doc") = "0087" Or rst("id_doc") = "0088" Then
        If Len(rst("id_cliente")) > 0 Then
            campo10 = rst("id_cliente")
        Else
            campo10 = "-"
        End If
    Else
        If Len(rst("id_cliente")) > 0 Then
            campo10 = rst("id_cliente")
        Else
            campo10 = "-"
            'MsgBox "CAMPO DOCUMENTO IDENTIDAD OBLIGATORIO", vbInformation, KEY_EMPRESA
            'Exit Sub
        End If
        
    End If
   If rst("id_doc") = "0001" And rst("id_cliente") = "" Then
    
    campo10 = X
    Else
    X = rst("id_cliente")
   End If
   
   If rst("ncliente") <> "" Then
        Campo11 = Mid(rst("ncliente"), 1, 58)
    Else
        Campo11 = "-"
    End If
    campo12 = "0.00"
    
    
    Campo13 = rst("valor_venta")
    Campo14 = rst("exonerado")
    If rst("exonerado") > 0 Then
        campo15 = rst("total")
    Else
        campo15 = 0#
    End If
    
    campo16 = 0#
    campo17 = rst("igv")
    If rst("id_doc") = "0040" Then
        campo18 = rst("valor_venta")
        campo19 = rst("igv") 'modificar y colocar el ivap
    Else
        campo18 = "0.00"
        campo19 = "0.00"
    End If
    
    
    campo20 = "0.00"
    campo21 = Format(rst("total"), "###0.00")
    Campo22 = Format(rst("tc"), "##0.000")
    
    If rst("id_doc") = "0007" Or rst("id_doc") = "0008" Or rst("id_doc") = "0087" Or rst("id_doc") = "0088" Or rst("id_doc") = "0097" Or rst("id_doc") = "0098" Then
       
        If rst("id_doc") = "0007" Then
            campo23 = Format(rst("fecha_fact"), "dd/mm/YYYY")
            GoTo 24
        End If
       If IsNull(rst("fecha_fact")) = True Then
            campo23 = Format(rst("fecha_fact"), "dd/mm/YYYY")
        Else
            campo23 = "01/01/0001"
       End If
    Else
        If rst("id_doc") = "0007" Then
            campo23 = Format(rst("fecha_fact"), "dd/mm/YYYY")
        Else
            campo23 = "01/01/0001"
        End If
        
    End If
    
24:
    If rst("id_doc") = "0007" Or rst("id_doc") = "0008" Or rst("id_doc") = "0087" Or rst("id_doc") = "0088" Or rst("id_doc") = "0097" Or rst("id_doc") = "0098" Then
        If rst("id_doc_fact") <> 0 Then
            campo24 = formato_item(rst("id_doc_fact"), 2)
        Else
            campo24 = "00"
        End If
    Else
        campo24 = "00"
    End If
     
     
     If rst("id_doc") = "0007" Or rst("id_doc") = "0008" Or rst("id_doc") = "0087" Or rst("id_doc") = "0088" Or rst("id_doc") = "0097" Or rst("id_doc") = "0098" Then
        If rst("serie_fact") = "0" Then
            MsgBox "MODIFICAR COMPROBANTE HAY UN VALOR INCORRECTO:" + Chr(13) + Chr(13) + "TIPO DOC  :" + rst("id_doc") + Chr(13) + "SERIE         :" + rst("serie") + Chr(13) + "NUMERO     :" + rst("numero"), vbInformation, KEY_EMPRESA
            Close #1
            Exit Sub
        Else
            If rst("id_doc") = "0007" Or rst("id_doc") = "0008" Then
                campo25 = formato_item(rst("serie_fact"), 4)
            Else
                campo25 = "0001" 'rst("serie_fact")
            End If
            campo26 = "004108" 'rst("numero_fact")
        End If
      Else
        campo25 = "-" 'aqui no hay descripcion si es nulo
        campo26 = "-"
     End If
    
    
    
    
    
Print #1, campo1 & "|" & campo2 & "|" & Campo3 & "|" & Campo4 & "|" & campo5 & "|" & Campo6 & "|" & campo7 & "|" & campo8 & "|" & campo9 & "|" & campo10 & "|" & Campo11 & "|" & campo12 & "|" & Campo13 & "|" & Campo14 & "|" & campo15 & "|" & campo16 & "|" & campo17 & "|" & campo18 & "|" & campo19 & "|" & campo20 & "|" & campo21 & "|" & Campo22 & "|" & campo23 & "|" & campo24 & "|" & campo25 & "|" & campo26 & "|" & campo27 & "|"
rst.MoveNext
Next i



Close #1
Me.lblNombreArchivo.Caption = Archivo
Me.lbloperacion.Visible = True

End Sub

Private Sub libro_compras(ByVal id_mes As String, ByVal id_anio As String, ByVal ruc As String)
Dim Nombre As String
Dim Archivo As String
Dim rura As String
Dim carpeta As String
     carpeta = Trim(id_mes & "-" & id_anio)
     d_periodo = "PERIODO" & Mid(FrmRegistroCompras.HfdPersona.TextMatrix(FrmRegistroCompras.HfdPersona.Row, 2), 18, 15) + Space(1) + FrmRegistroCompras.HfdPersona.TextMatrix(FrmRegistroCompras.HfdPersona.Row, 3)
     d_ruc = ruc
     
SqlDatos = "SELECT * FROM movimiento_compra WHERE ruc='" & ruc & "' AND id_mes='" & id_mes & "' AND id_anio='" & id_anio & "' AND id_doc<>'0089' ORDER BY fecha_emision,id_doc,serie,numero ASC"
Call ConfiguraRst(SqlDatos)
'demo:     LE2053151604520130100080100001111.txt
Archivo = "LE" & ruc & id_anio & id_mes & "00" & "080100001111"
'archivo = Trim("LE" & d_ruc & FrmRegistroCompras.HfdPersona.TextMatrix(FrmRegistroCompras.HfdPersona.Row, 3) & FrmRegistroCompras.HfdPersona.TextMatrix(FrmRegistroCompras.HfdPersona.Row, 1) & "0008010001OIM1")
ruta = Trim(App.Path & "\ple\" & carpeta & "\" & Archivo & ".txt")
Open ruta For Append As #1
Dim periodo As Date
periodo = Format(CVDate("01" & "-" & id_mes & "-" & id_anio), "YYYY-mm-dd")
Me.ProgressBar1.Min = 0
Me.ProgressBar1.Max = rst.RecordCount - 1
For i = 0 To rst.RecordCount - 1
    Me.ProgressBar1.Value = i
    campo1 = rst("id_anio") & rst("id_mes") & "00"
    If DateDiff("m", rst("fecha_emision"), periodo) >= 1 Then
        If DateDiff("m", rst("fecha_emision"), periodo) <= 12 Then
            campo32 = "6"
        End If
        If DateDiff("m", rst("fecha_emision"), periodo) > 12 Then
            campo32 = "7"
        End If
            
        
    Else
        campo32 = "1"
    End If
    campo2 = formato_item(i + 1, 8)
    Campo3 = Format(rst("fecha_emision"), "dd/mm/YYYY")
    If rst("id_doc") = "0014" Then
        'campo4 = Format(rst("fecha_cancelacion"), "dd/mm/YYYY")
        Campo4 = "31/03/2013" 'Format(DateAdd("d", 15, rst("fecha_emision")), "dd/mm/YYYY")
        
    Else
        Campo4 = ""
    End If
    
    campo5 = formato_item(rst("id_doc"), 2)
    If rst("id_doc") = "0001" Or rst("id_doc") = "0003" Or rst("id_doc") = "0004" Or rst("id_doc") = "0007" Or rst("id_doc") = "0008" Then
        Campo6 = formato_item(rst("serie"), 4)
        GoTo saltar6
    End If
    If rst("id_doc") = "0050" Or rst("id_doc") = "0052" Then
        Campo6 = formato_item(rst("serie"), 3)
    Else
        Campo6 = rst("serie")
    End If
    
    
saltar6:
If rst("id_doc") = "0050" Or rst("id_doc") = "0052" Then
    campo7 = rst("anio_dua")
Else
    campo7 = "0"
End If

    
    campo8 = rst("numero")
    campo9 = "0"
    If rst("id_doc") = "0001" Or rst("id_doc") = "0002" Then
        campo10 = "6"
    Else
        campo10 = "0"
    End If
    
    If rst("id_doc") = "0000" Or rst("id_doc") = "0003" Or rst("id_doc") = "0005" Or rst("id_doc") = "0006" Or rst("id_doc") = "0007" Or rst("id_doc") = "0008" Or rst("id_doc") = "0011" Or rst("id_doc") = "0012" Or rst("id_doc") = "0013" Or rst("id_doc") = "0014" Or rst("id_doc") = "0015" Or rst("id_doc") = "0016" Or rst("id_doc") = "0018" Or rst("id_doc") = "0019" Or rst("id_doc") = "0022" Or rst("id_doc") = "0023" Or rst("id_doc") = "0026" Or rst("id_doc") = "0028" Or rst("id_doc") = "0030" Or rst("id_doc") = "0034" Or rst("id_doc") = "0035" Or rst("id_doc") = "0036" Or rst("id_doc") = "0037" Or rst("id_doc") = "0055" Or rst("id_doc") = "0056" Or rst("id_doc") = "0087" Or rst("id_doc") = "0088" Or rst("id_doc") = "0091" Or rst("id_doc") = "0097" Or rst("id_doc") = "0098" Then
        If rst("id_proveedor") <> "" Then
            Campo11 = rst("id_proveedor")
        Else
            Campo11 = "-"
        End If
    Else
        If rst("id_proveedor") <> "" Then
            Campo11 = rst("id_proveedor")
        Else
            MsgBox "INGRESE RUC PARA LA DUA" + Space(1) + rst("serie") + "-" + rst("numero"), vbInformation, KEY_EMPRESA
            Exit Sub
        End If
    End If
    
    
   
    If rst("nproveedor") <> "" Then
        campo12 = Mid(rst("nproveedor"), 1, 58)
    Else
        campo12 = "-"
    End If
    If rst("id_tipo_compra") = "01" Then
        Campo13 = rst("valor_venta")
        Campo14 = rst("igv")
        campo15 = 0#
        campo16 = 0#
        campo17 = 0#
        campo18 = 0#
    End If
    If rst("id_tipo_compra") = "02" Then
        Campo13 = 0#
        Campo14 = 0#
        campo15 = rst("valor_venta")
        campo16 = rst("igv")
        campo17 = 0#
        campo18 = 0#
    End If
    If rst("id_tipo_compra") = "03" Then
        Campo13 = 0#
        Campo14 = 0#
        campo15 = 0#
        campo16 = 0#
        campo17 = rst("valor_venta")
        campo18 = rst("igv")
    End If
        
       
    campo19 = rst("exonerado")
    campo20 = rst("isc")
    campo21 = rst("otros")
    Campo22 = Format(rst("total"), "###0.00")
    campo23 = Format(rst("tc"), "##0.000")
    'If campo5 = "07" Or campo5 = "08" Or campo5 = "87" Or campo5 = "88" Or campo5 = "97" Or campo5 = "98" Then
    If rst("id_doc") = "0007" Or rst("id_doc") = "0008" Or rst("id_doc") = "0087" Or rst("id_doc") = "0088" Or rst("id_doc") = "0097" Or rst("id_doc") = "0098" Then
       If IsNull(rst("fecha_fact")) = True Then
            campo24 = Format(rst("fecha_fact"), "dd/mm/YYYY")
        Else
            campo24 = "01/01/0001"
       End If
    Else
        campo24 = "01/01/0001"
    End If
    If rst("id_doc") = "0007" Or rst("id_doc") = "0008" Or rst("id_doc") = "0087" Or rst("id_doc") = "0088" Or rst("id_doc") = "0097" Or rst("id_doc") = "0098" Then
        If rst("id_doc_fact") <> 0 Then
            campo25 = formato_item(rst("id_doc_fact"), 2)
        Else
            campo25 = "00"
        End If
    Else
        campo25 = "00"
    End If
     If rst("id_doc") = "0007" Or rst("id_doc") = "0008" Or rst("id_doc") = "0087" Or rst("id_doc") = "0088" Or rst("id_doc") = "0097" Or rst("id_doc") = "0098" Then
        If rst("serie_fact") = "0" Then
            If rst("id_doc") = "0007" Or rst("id_doc") = "0008" Then
                campo26 = formato_item("1", 4)
            Else
                campo26 = rst("serie_fact")
            End If
            campo27 = "007557" 'rst("numero_fact")
        Else
            If rst("id_doc") = "0007" Or rst("id_doc") = "0008" Then
                campo26 = formato_item(rst("serie_fact"), 4)
            Else
                campo26 = rst("serie_fact")
            End If
            campo27 = rst("numero_fact")
        End If
      Else
        campo26 = "-" 'aqui no hay descripcion si es nulo
        campo27 = "-"
     End If
    
    
    
     If rst("id_doc") = "0091" Or rst("id_doc") = "0097" Or rst("id_doc") = "0098" Then
     If Val(rst("numero_no_domiciliado")) > 0 Then
        Campo28 = rst("numero_no_domiciliado")
       
     Else
        Campo28 = "-"
     End If
     Else
        Campo28 = "-"
     End If
     
     
     
     If rst("numero_detrac") <> "0" Then
        Campo29 = Format(rst("fecha_detrac"), "dd/mm/YYYY")
     Else
        Campo29 = "01/01/0001"
    End If
    
    
    If rst("numero_detrac") <> "0" Then
        Campo30 = rst("numero_detrac")
     Else
        Campo30 = "0"
    End If
    If rst("retencion") > 0 Then
        campo31 = "1"
    Else
        campo31 = "0"
    End If
    
    
    
Print #1, campo1 & "|" & campo2 & "|" & Campo3 & "|" & Campo4 & "|" & campo5 & "|" & Campo6 & "|" & campo7 & "|" & campo8 & "|" & campo9 & "|" & campo10 & "|" & Campo11 & "|" & campo12 & "|" & Campo13 & "|" & Campo14 & "|" & campo15 & "|" & campo16 & "|" & campo17 & "|" & campo18 & "|" & campo19 & "|" & campo20 & "|" & campo21 & "|" & Campo22 & "|" & campo23 & "|" & campo24 & "|" & campo25 & "|" & campo26 & "|" & campo27 & "|"; Campo28 & "|" & Campo29 & "|" & Campo30 & "|" & campo31 & "|" & campo32 & "|"
rst.MoveNext
Next i



Close #1
Me.lblNombreArchivo.Caption = Archivo
Me.lbloperacion.Visible = True

End Sub

Private Sub ChameleonBtn1_Click()

End Sub

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 2000
End Sub
