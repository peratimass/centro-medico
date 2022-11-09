VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MsComCtl.ocx"
Begin VB.Form FrmRegistroSunatDiario 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6120
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   6120
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "GENERAR LIBRO ELECTRONICO MAYOR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2160
      Width           =   5535
   End
   Begin VB.CommandButton cmdGenrarLibro 
      Caption         =   "GENERAR LIBRO ELECTRONICO DIARIO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Width           =   5535
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   165
      TabIndex        =   1
      Top             =   240
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NOMBRE ARCHIVO :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   4
      Top             =   1080
      Width           =   1470
   End
   Begin VB.Label lblNombreArchivo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " "
      Height          =   195
      Left            =   1800
      TabIndex        =   3
      Top             =   1080
      Width           =   45
   End
   Begin VB.Label lbloperacion 
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      Caption         =   "OPERACION REALIZADA CON EXITO !!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   1200
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00808080&
      Height          =   2505
      Left            =   120
      Top             =   120
      Width           =   5850
   End
End
Attribute VB_Name = "FrmRegistroSunatDiario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede
Private Sub cmdGenrarLibro_Click()
    

        Call libro_diario(FrmRegistroDiario.HfdPersona.TextMatrix(FrmRegistroDiario.HfdPersona.Row, 1), FrmRegistroDiario.HfdPersona.TextMatrix(FrmRegistroDiario.HfdPersona.Row, 3), FrmRegistroDiario.HfdPersona.TextMatrix(FrmRegistroDiario.HfdPersona.Row, 0))
        FrmRegistroDiario.Procedencia = Neutro
        Exit Sub
        
    
End Sub
Private Sub libro_diario(ByVal id_mes As String, ByVal id_anio As String, ByVal ruc As String)
Dim Nombre As String
Dim Archivo As String
     
     
SqlDatos = "SELECT * FROM registro_diario_detalle WHERE ruc='" & ruc & "' AND id_mes='" & id_mes & "' AND id_anio='" & id_anio & "'  ORDER BY fecha ASC,id_detalle ASC"
Call ConfiguraRst(SqlDatos)
'demo:     LERRRRRRRRRRRAAAAMM0005010000OIM1.TXT
'X = "140100001111"
Archivo = "LE" & ruc & id_anio & id_mes & "00" & "05010000" & "1" & "1" & "1" & "1"
                                                 
'archivo = Trim("LE" & d_ruc & FrmRegistroCompras.HfdPersona.TextMatrix(FrmRegistroCompras.HfdPersona.Row, 3) & FrmRegistroCompras.HfdPersona.TextMatrix(FrmRegistroCompras.HfdPersona.Row, 1) & "0008010001OIM1")
ruta = Trim(App.Path & "\ple\" & Archivo & ".txt")
Open ruta For Append As #1
Dim periodo As Date
periodo = Format(CVDate("01" & "/" & id_mes & "/" & id_anio), "YYYY-mm-dd")
Me.ProgressBar1.Min = 0
Me.ProgressBar1.Max = rst.RecordCount - 1
For i = 0 To rst.RecordCount - 1
    Me.ProgressBar1.Value = i
    periodo = Format(CVDate("01" & "-" & rst("id_mes") & "-" & rst("id_anio")), "YYYY-mm-dd")
    
      campo1 = rst("id_anio") & rst("id_mes") & "00"
    campo9 = "1"
    If DateDiff("m", rst("fecha"), periodo) = 0 Then
        campo9 = "1"
        
    End If
    
    If DateDiff("m", rst("fecha"), periodo) > 0 Then
        campo9 = "8"
        
    End If
    
    
    'strCadena = "SELECT count(*) FROM registro_diario_detalle WHERE ruc='" & ruc & "' AND id_mes='" & id_mes & "' AND id_anio='" & id_anio & "' AND id_cuenta='" & Trim(rst("id_cuenta")) & "'"
    'Call ConfiguraRstT(strCadena)
    'If rstT(0) > 1 Then
      
     '   campo2 = formato_item(rst("num_correlativo") + 1, 2)
    'Else
        campo2 = rst("num_correlativo")
    'End If
    Campo3 = "01"
    Campo4 = rst("id_cuenta")
    campo5 = Format(rst("fecha"), "dd/mm/YYYY")
    Campo6 = Mid(rst("glosa"), 1, 99)
    campo7 = rst("debe")
    campo8 = rst("haber")
    
    
Print #1, campo1 & "|" & campo2 & "|" & Campo3 & "|" & Campo4 & "|" & campo5 & "|" & Campo6 & "|" & campo7 & "|" & campo8 & "|" & campo9 & "|"
rst.MoveNext
Next i



Close #1
Me.lblNombreArchivo.Caption = Archivo
Me.lbloperacion.Visible = True

End Sub
Private Sub libro_mayor(ByVal id_mes As String, ByVal id_anio As String, ByVal ruc As String)
Dim Nombre As String
Dim Archivo As String

SqlDatos = "SELECT DISTINCT id_cuenta FROM registro_diario_detalle WHERE  ruc='" & ruc & "' AND id_mes='" & id_mes & "' AND id_anio='" & id_anio & "'  ORDER BY fecha ASC,id_detalle ASC"
Call ConfiguraRst(SqlDatos)
'LERRRRRRRRRRRAAAAMM0006010000OIM1.TXT
Archivo = "LE" & ruc & id_anio & id_mes & "00" & "06010000" & "1" & "1" & "1" & "1"
ruta = Trim(App.Path & "\ple\" & Archivo & ".txt")
Open ruta For Append As #1
Dim periodo As Date
periodo = Format(CVDate("01" & "/" & id_mes & "/" & id_anio), "YYYY-mm-dd")

Me.ProgressBar1.Min = 0
Me.ProgressBar1.Max = rst.RecordCount - 1

For i = 0 To rst.RecordCount - 1
    Me.ProgressBar1.Value = i
    strCadena = "SELECT * FROM registro_diario_detalle WHERE id_cuenta='" & rst("id_cuenta") & "' AND ruc='" & ruc & "' AND id_mes='" & id_mes & "' AND id_anio='" & id_anio & "' ORDER BY fecha ASC"
    Call ConfiguraRstT(strCadena)
    If rstT.RecordCount > 0 Then
        rstT.MoveFirst
        For j = 0 To rstT.RecordCount - 1
            periodo = Format(CVDate("01" & "-" & rstT("id_mes") & "-" & rstT("id_anio")), "YYYY-mm-dd")
            campo1 = rstT("id_anio") & rstT("id_mes") & "00"
            campo8 = "1"
            If DateDiff("m", rstT("fecha"), periodo) = 0 Then
                campo8 = "1"
            End If
            campo2 = rstT("num_correlativo")
            Campo3 = rstT("id_cuenta")
            Campo4 = Format(rstT("fecha"), "dd/mm/YYYY")
            campo5 = Mid(rstT("glosa"), 1, 99)
            Campo6 = rstT("debe")
            campo7 = rstT("haber")
            Print #1, campo1 & "|" & campo2 & "|" & Campo3 & "|" & Campo4 & "|" & campo5 & "|" & Campo6 & "|" & campo7 & "|" & campo8 & "|"
            rstT.MoveNext
        Next j
        rst.MoveNext
    End If
Next i
Close #1
Me.lblNombreArchivo.Caption = Archivo
Me.lbloperacion.Visible = True
End Sub

Private Sub libro_compras(ByVal id_mes As String, ByVal id_anio As String, ByVal ruc As String)
Dim Nombre As String
Dim Archivo As String
Dim rura As String
     d_periodo = "PERIODO" & Mid(FrmRegistroCompras.HfdPersona.TextMatrix(FrmRegistroCompras.HfdPersona.Row, 2), 18, 15) + Space(1) + FrmRegistroCompras.HfdPersona.TextMatrix(FrmRegistroCompras.HfdPersona.Row, 3)
     d_ruc = ruc
     
SqlDatos = "SELECT * FROM movimiento_compra WHERE ruc='" & ruc & "' AND id_mes='" & id_mes & "' AND id_anio='" & id_anio & "' AND id_doc<>'0089' ORDER BY fecha_emision,id_doc,serie,numero ASC"
Call ConfiguraRst(SqlDatos)
'demo:     LE2053151604520130100080100001111.txt
Archivo = "LE" & ruc & id_anio & id_mes & "00" & "080100001111"
'archivo = Trim("LE" & d_ruc & FrmRegistroCompras.HfdPersona.TextMatrix(FrmRegistroCompras.HfdPersona.Row, 3) & FrmRegistroCompras.HfdPersona.TextMatrix(FrmRegistroCompras.HfdPersona.Row, 1) & "0008010001OIM1")
ruta = Trim(App.Path & "\ple\" & Archivo & ".txt")
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
            MsgBox "INGRESE SERIE MODIFICAION FACTURA:" + rst("id_doc") + ":" + rst("serie") + "-" + rst("numero"), vbInformation, KEY_EMPRESA
            Exit Sub
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

Private Sub Command1_Click()
        Call libro_mayor(FrmRegistroDiario.HfdPersona.TextMatrix(FrmRegistroDiario.HfdPersona.Row, 1), FrmRegistroDiario.HfdPersona.TextMatrix(FrmRegistroDiario.HfdPersona.Row, 3), FrmRegistroDiario.HfdPersona.TextMatrix(FrmRegistroDiario.HfdPersona.Row, 0))
 Exit Sub
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

