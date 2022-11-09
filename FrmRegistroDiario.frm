VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmRegistroDiario 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8775
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15015
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   15015
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtEmpresa 
      Height          =   285
      Left            =   480
      TabIndex        =   3
      Text            =   "TxtEmpresa"
      Top             =   6360
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox txtRuc 
      Height          =   285
      Left            =   3480
      TabIndex        =   2
      Text            =   "txtRuc"
      Top             =   6360
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox txtAnio 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Buscar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2760
      TabIndex        =   0
      Top             =   720
      Width           =   1335
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   4680
      TabIndex        =   4
      Top             =   840
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   12840
      Top             =   6600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegistroDiario.frx":0000
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegistroDiario.frx":0454
            Key             =   "(Ple)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegistroDiario.frx":37A2
            Key             =   "(Importar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegistroDiario.frx":3E74
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegistroDiario.frx":4194
            Key             =   "(Exportar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegistroDiario.frx":4866
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegistroDiario.frx":4CBA
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegistroDiario.frx":510E
            Key             =   "(RCompras)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmRegistroDiario.frx":6880
            Key             =   "(RVentas)"
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdPersona 
      Height          =   7335
      Left            =   240
      TabIndex        =   5
      Top             =   1200
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   12938
      _Version        =   393216
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   7305
      Left            =   13800
      TabIndex        =   6
      Top             =   1320
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   12885
      BandCount       =   1
      ForeColor       =   -2147483635
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      Orientation     =   1
      EmbossPicture   =   -1  'True
      _CBWidth        =   900
      _CBHeight       =   7305
      _Version        =   "6.0.8169"
      Child1          =   "TlbAcciones"
      MinHeight1      =   840
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   780
         Left            =   30
         TabIndex        =   7
         Top             =   30
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   1376
         ButtonWidth     =   1482
         ButtonHeight    =   1376
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ImgIconos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   15
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "   Nuevo   "
               Key             =   "(Nuevo)"
               Object.ToolTipText     =   "Nuevo"
               ImageKey        =   "(Nuevo)"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ingresar"
               Key             =   "(Modificar)"
               Object.ToolTipText     =   "Modificar"
               ImageKey        =   "(Modificar)"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Importar"
               Key             =   "(Importar)"
               ImageKey        =   "(Importar)"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Exportar"
               Key             =   "(Exportar)"
               ImageKey        =   "(Exportar)"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Exp.Dia"
               Key             =   "(ExportarDia)"
               ImageKey        =   "(Exportar)"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "PLE 3.0"
               Key             =   "(Ple)"
               ImageKey        =   "(Ple)"
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Eliminar"
               Key             =   "(Eliminar)"
               Object.ToolTipText     =   "Eliminar"
               ImageKey        =   "(Eliminar)"
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Salir"
               Key             =   "(Salir)"
               ImageKey        =   "(Salir)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LIBRO DIARIO :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   165
      Left            =   1170
      TabIndex        =   10
      Top             =   200
      Width           =   1245
   End
   Begin VB.Shape Shape1 
      Height          =   8775
      Left            =   0
      Top             =   0
      Width           =   15015
   End
   Begin VB.Label lblEmpresa 
      BackStyle       =   0  'Transparent
      Caption         =   "Registro de Ventas Mensual:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   2640
      TabIndex        =   9
      Top             =   195
      Width           =   7035
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AÑO:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   570
      TabIndex        =   8
      Top             =   720
      Width           =   495
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00DFDFE0&
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   435
      Left            =   240
      Top             =   60
      Width           =   13455
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00DFDFE0&
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   555
      Left            =   240
      Top             =   600
      Width           =   4215
   End
End
Attribute VB_Name = "FrmRegistroDiario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede

Private Sub Command1_Click()
Call actualizar_anio(Me.TxtRuc.Text, Me.txtanio.Text)
End Sub
Public Sub actualizar_anio(ByVal ruc As String, ByVal Anio As String)
strCadena = "SELECT ruc,mes,descripcion AS Periodo,anio, estado  as Estado FROM  dbo.RegistroVentas WHERE ruc='" & Trim(Me.TxtRuc.Text) & "' AND anio LIKE '%" & Trim(Me.txtanio.Text) & "%' ORDER BY anio,mes"
Call llenarGrid(Me.HfdPersona, Me)
End Sub

Private Sub Command2_Click()

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 50
Me.LblEmpresa.Caption = KEY_EMPRESA + Space(2) + "***[" + "RUC:" + Space(2) + KEY_RUC + "]***"
Me.TxtEmpresa.Text = KEY_EMPRESA
Me.TxtRuc.Text = KEY_RUC
Call actualizar


 
End Sub
Public Sub actualizar()
strCadena = "SELECT ruc,mes,R.descripcion,anio, R.debe ,R.haber,razon FROM registro_diario R ORDER BY R.anio,R.mes"
Call llenarGrid(Me.HfdPersona, Me)
  
  
End Sub

Private Sub HfdPersona_SelChange()
If Len(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) = 11 Then
    TlbAcciones.Buttons(KEY_UPDATE).Enabled = True
    TlbAcciones.Buttons(KEY_DELETE).Enabled = True
    TlbAcciones.Buttons(KEY_IMPORTAR).Enabled = True
    TlbAcciones.Buttons(KEY_EXPORTAR).Enabled = True
    TlbAcciones.Buttons("(Ple)").Enabled = True
    TlbAcciones.Buttons("(ExportarDia)").Enabled = True
    
Else
    TlbAcciones.Buttons(KEY_UPDATE).Enabled = False
    TlbAcciones.Buttons(KEY_DELETE).Enabled = False
    TlbAcciones.Buttons("(Ple)").Enabled = False
  TlbAcciones.Buttons("(ExportarDia)").Enabled = False
    TlbAcciones.Buttons(KEY_IMPORTAR).Enabled = True
    TlbAcciones.Buttons(KEY_EXPORTAR).Enabled = True
    
End If
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim afecto As Double, exonerado As Double, igv As Double, Total                 As Double
        
Select Case Button.key
    Case KEY_NEW
      Procedencia = nuevo
      FrmDetalleRegistrodiario.Show
      Exit Sub
    Case KEY_UPDATE
      Procedencia = Modificar
      FrmregistroDiarioList.Show
    Case KEY_DELETE
      If MsgBox(MSGELIMINAR, vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
        If MsgBox("Se Borraran Todos los registros relacionados", vbQuestion + vbYesNo) = vbYes Then
            
            'strCadena = "DELETE FROM RegistroVentasDetalle WHERE Ruc='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) & "' AND mes='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1)) & "' AND anio='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 3)) & "'"
            'CnBd.Execute (strCadena)
            strCadena = "DELETE FROM registro_diario WHERE ruc='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) & "' AND mes='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1)) & "' AND anio='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 3)) & "'"
            CnBd.Execute (strCadena)
             
            strCadena = "DELETE FROM registro_diario_detalle WHERE ruc='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) & "' AND id_mes='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1)) & "' AND id_anio='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 3)) & "'"
            CnBd.Execute (strCadena)
             
            Call actualizar
        End If
      End If
    Case KEY_IMPORTAR
       ' FrmImportar.Show
        Call ImportarVentas(Trim(Me.TxtRuc.Text))
    Case "(Ple)"
        Procedencia = nuevo
        FrmRegistroSunatDiario.Show
    Case KEY_EXPORTAR
     
    Dim bol_ini As String
    Dim bol_fin As String
    Dim i As Double
    Dim fecha As Date
    strCadena = "DELETE FROM registroventassunat WHERE dni_save='" & KEY_USUARIO & "' AND ruc='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) & "' "
    CnBd.Execute (strCadena)
     
    
    
strCadena = "SELECT * FROM movimiento_venta WHERE ruc='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) & " ' AND id_mes='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1)) & "' AND id_anio='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 3)) & "' AND id_doc<>'0003' AND id_doc<>'0012'AND id_doc<>'0109'  ORDER BY fecha_emision,id_doc,serie,numero ASC"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
rst.MoveFirst
fecha = rst("fecha_emision")
For i = 0 To rst.RecordCount - 1
        If Len(Trim(rst("id_cliente"))) = 11 Then
            tdoc = 6
        Else
            tdoc = 0
        End If
               
        strCadena = "INSERT INTO registroventassunat(ruc,fecha,mes,anio,id_doc,serie,numero,tdoc,ruccliente,NombreCliente,afecto,exonerado,igv,total,retencion,tc,fecha_F,doc_codF,serieF,numeroF,dni_save,anulado)VALUES " & _
        "('" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) & "','" & Format(rst("fecha_emision"), "YYYY/mm/dd") & "','" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1)) & "','" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 3)) & "','" & Trim(formato_item(Val(rst("id_doc")), 2)) & "'," & _
        "'" & rst("serie") & "','" & rst("numero") & "','" & tdoc & "','" & rst("id_cliente") & "','" & Trim(Replace(rst("ncliente"), "'", "")) & "','" & rst("valor_venta") & "','" & rst("exonerado") & "','" & rst("igv") & "'," & _
        "'" & rst("total") & "','" & rst("retencion") & "','" & rst("tc") & "','" & Format(rst("fecha_fact"), "YYYY-mm-dd") & "','" & rst("id_doc_fact") & "','" & rst("serie_fact") & "','" & rst("numero_fact") & "','" & KEY_USUARIO & "','" & rst("anulado") & "' )"
        CnBd.Execute (strCadena)
         
        rst.MoveNext
Next i
End If
strCadena = "SELECT DISTINCT serie FROM movimiento_venta WHERE ruc='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) & "' AND id_mes='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1)) & "' AND id_anio='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 3)) & "' AND id_doc='0003' AND anulado='no' ORDER BY serie ASC"
Call ConfiguraRstT(strCadena)

strCadena = "SELECT DISTINCT fecha_emision FROM movimiento_venta WHERE ruc='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) & "' AND id_mes='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1)) & "' AND id_anio='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 3)) & "' AND id_doc='0003' AND anulado='no' ORDER BY fecha_emision ASC"
Call ConfiguraRst(strCadena)

If rst.RecordCount > 0 And rstT.RecordCount > 0 Then
    rst.MoveFirst
    rstT.MoveFirst
For i = 0 To rst.RecordCount - 1
    rstT.MoveFirst
    For j = 0 To rstT.RecordCount - 1
    strCadena = "SELECT sum(valor_venta),sum(exonerado),sum(igv),sum(total) FROM movimiento_venta WHERE fecha_emision='" & Format(rst("fecha_emision"), "YYYY-mm-dd") & "' AND serie='" & rstT("serie") & "' AND ruc='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) & "' AND id_mes='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1)) & "' AND id_anio='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 3)) & "' AND id_doc='0003'"
    Call ConfiguraTemporal(strCadena)
    If rstTemporal.RecordCount > 0 Then
            If IsNull(rstTemporal(0)) = True Then
                afecto = 0
            Else
                afecto = rstTemporal(0)
            End If
            If IsNull(rstTemporal(1)) = True Then
                exonerado = 0
            Else
                exonerado = rstTemporal(1)
            End If
            If IsNull(rstTemporal(2)) = True Then
                igv = 0
            Else
                igv = rstTemporal(2)
            End If
            If IsNull(rstTemporal(3)) = True Then
                Total = 0
            Else
                Total = rstTemporal(3)
            End If
        strCadena = "SELECT serie,numero FROM movimiento_venta WHERE fecha_emision='" & Format(rst("fecha_emision"), "YYYY-mm-dd") & "' AND serie='" & rstT("serie") & "' AND ruc='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) & "' AND id_mes='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1)) & "' AND id_anio='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 3)) & "' AND id_doc='0003'  ORDER BY numero ASC"
        Call ConfiguraTemporal(strCadena)
        If rstTemporal.RecordCount > 0 Then
            rstTemporal.MoveFirst
            bol_ini = rstTemporal("numero")
            rstTemporal.MoveLast
            bol_fin = rstTemporal("numero")
        Else
            GoTo avanzar
        End If
        
        If rstTemporal.RecordCount = 1 Then
            boleta = rstTemporal("numero")
        Else
        boleta = Right(bol_ini, 7) + "-" + Right(bol_fin, 7)
        End If
        
        strCadena = "INSERT INTO registroventassunat(ruc,fecha,mes,anio,id_doc,serie,numero,tdoc,ruccliente,NombreCliente,afecto,exonerado,igv,total,tc,dni_save)VALUES " & _
        "('" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) & "','" & Format(rst("fecha_emision"), "YYYY-mm-dd") & "','" & Trim(Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1))) & "','" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 3)) & "','03'," & _
        "'" & rstT("serie") & "','" & Trim(boleta) & "','0','','','" & afecto & "','" & exonerado & "','" & igv & "','" & Total & "','" & GetCambio(rst("fecha_emision")) & "','" & KEY_USUARIO & "')"
        CnBd.Execute (strCadena)
         
    End If
avanzar:
    rstT.MoveNext
Next j
    rst.MoveNext
Next i
End If

strCadena = "SELECT DISTINCT serie FROM movimiento_venta WHERE ruc='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) & "' AND id_mes='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1)) & "' AND id_anio='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 3)) & "' AND id_doc='0012' ORDER BY serie ASC"
Call ConfiguraRstT(strCadena)

strCadena = "SELECT DISTINCT fecha_emision FROM movimiento_venta WHERE ruc='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) & "' AND id_mes='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1)) & "' AND id_anio='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 3)) & "' AND id_doc='0012' ORDER BY fecha_emision ASC"
Call ConfiguraRst(strCadena)

If rst.RecordCount > 0 And rstT.RecordCount > 0 Then
    rst.MoveFirst
    rstT.MoveFirst
For i = 0 To rst.RecordCount - 1
    rstT.MoveFirst
    For j = 0 To rstT.RecordCount - 1
    strCadena = "SELECT sum(valor_venta),sum(exonerado),sum(igv),sum(total) FROM movimiento_venta WHERE fecha_emision='" & Format(rst("fecha_emision"), "YYYY-mm-dd") & "' AND serie='" & rstT("serie") & "' AND ruc='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) & "' AND id_mes='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1)) & "' AND id_anio='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 3)) & "' AND id_doc='0012'"
    Call ConfiguraTemporal(strCadena)
    If rstTemporal.RecordCount > 0 Then
            If IsNull(rstTemporal(0)) = True Then
                afecto = 0
            Else
                afecto = rstTemporal(0)
            End If
            If IsNull(rstTemporal(1)) = True Then
                exonerado = 0
            Else
                exonerado = rstTemporal(1)
            End If
            If IsNull(rstTemporal(2)) = True Then
                igv = 0
            Else
                igv = rstTemporal(2)
            End If
            If IsNull(rstTemporal(3)) = True Then
                Total = 0
            Else
                Total = rstTemporal(3)
            End If
        strCadena = "SELECT serie,numero FROM movimiento_venta WHERE fecha_emision='" & Format(rst("fecha_emision"), "YYYY-mm-dd") & "' AND serie='" & rstT("serie") & "' AND ruc='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) & "' AND id_mes='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1)) & "' AND id_anio='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 3)) & "' AND id_doc='0012'  ORDER BY numero ASC"
        Call ConfiguraTemporal(strCadena)
        If rstTemporal.RecordCount > 0 Then
            rstTemporal.MoveFirst
            bol_ini = rstTemporal("numero")
            rstTemporal.MoveLast
            bol_fin = rstTemporal("numero")
        Else
            GoTo avanzar1
        End If
        
        If rstTemporal.RecordCount = 1 Then
            boleta = rstTemporal("numero")
        Else
        boleta = Right(bol_ini, 7) + "-" + Right(bol_fin, 7)
        End If
        
        strCadena = "INSERT INTO registroventassunat(ruc,fecha,mes,anio,id_doc,serie,numero,tdoc,ruccliente,NombreCliente,afecto,exonerado,igv,total,tc,dni_save,anulado)VALUES " & _
        "('" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) & "','" & Format(rst("fecha_emision"), "YYYY-mm-dd") & "','" & Trim(Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1))) & "','" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 3)) & "','03'," & _
        "'" & rstT("serie") & "','" & Trim(boleta) & "','0','','','" & afecto & "','" & exonerado & "','" & igv & "','" & Total & "','" & GetCambio(rst("fecha_emision")) & "','" & KEY_USUARIO & "','" & rst("anulado") & "' )"
        CnBd.Execute (strCadena)
         
    End If
avanzar1:
    rstT.MoveNext
Next j
    rst.MoveNext
Next i
End If


 Procedencia = nuevo
 frmTempExportExcel.Show
 'FrmRegistroVentas.Procedencia = Neutro
    Exit Sub
strCadena = "SELECT fecha, doc_cod, serie, numero, RucCliente, NombreCliente, afecto, exonerado, igv, total,retencion,anulado " & _
"FROM dbo.RegistroVentasSunat WHERE Ruc='" & Trim(TxtRuc.Text) & "' AND mes='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1)) & "' AND anio='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 3)) & "' ORDER BY fecha,doc_cod,serie,numero ASC"
Call ConfiguraRst(strCadena)
Ans = ShowMultiReport(rst, "RptRegistroVentas4", , App.Path + "\Reportes\")

    Case "(ExportarDia)"
           strCadena = "SELECT fecha,id_doc,serie,numero,tdoc,ruccliente,NombreCliente,afecto,exonerado,igv,total,retencion,otros,tc FROM registroventassunat WHERE mes='" & Trim(Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 1))) & "' AND anio='" & Trim(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 3)) & "' AND ruc='" & KEY_RUC & "' AND (id_doc='01' OR id_doc='03') AND anulado='no'"
           Call ConfiguraRst(strCadena)
           Ans = ShowMultiReport(rst, "RptRegistroVentasdia", , App.Path + "\Reportes\")
    Case "(Salir)"
      Unload Me
  End Select
End Sub
Public Sub ImportarVentas(ByVal ruc As String)
Dim rstRemoto As New ADODB.Record
  Set rstT = New ADODB.Recordset
  rstT.CursorLocation = adUseClient
  strCadena = "SELECT * FROM RegistroVentasDetalle WHERE mes='01'OR mes='02' AND anio='2013' WHERE ruc='20104050337' ORDER BY  codigounico ASC"
  rstT.Open strCadena, cnbd1, adOpenKeyset, adLockOptimistic
  
If rstT.RecordCount > 0 Then
    rstT.MoveFirst
    For i = 0 To rstT.RecordCount - 1
        strCadena = "P_insert_venta('" & formato_item(rstT("doc_cod"), 4) & "','00001','" & formato_item(rstT("idformapago"), 2) & "','" & rstT("moneda") & "','no'," & _
            "'" & formato_item(Trim(rstT("serie")), 3) & "','" & formato_item(Trim(rstT("numero")), 6) & "','" & rstT("RucCliente") & "','" & rstT("NombreCliente") & "','" & rstT("afecto") & "','" & rstT("igv") & "','" & rstT("exonerado") & "','" & rstT("total") & "','0'," & _
            "'" & rstT("total") & "','0','" & Format(rstT("fecha"), "YYYY-mm-dd") & "','" & Format(rstT("fecha"), "YYYY-mm-dd") & "','00001','" & KEY_USUARIO & "','" & rstT("tc") & "','no','" & rstT("mes") & "','" & rstT("anio") & "','" & ruc & "')"
            CnBd.Execute (strCadena)
             
            
            
            id_venta = LastRegistroRUC("movimiento_venta", "id_venta")
            If rstT("anulado") = "V" Then
                strCadena = "UPDATE movimiento_venta SET anulado='si',ncliente='A N U L A D O',id_cliente='',afecto='0',exonerado='0',igv='0',total='0',saldo='0' WHERE id_venta='" & id_venta & "'"
                CnBd.Execute (strCadena)
                 
            End If
            rstT.MoveNext
            DoEvents
    Next i
End If
End Sub

Private Sub llenarGrid(ByVal Grilla As MSHFlexGrid, ByVal Formulario As Form)
On Error GoTo salir
Dim Acumulado As Double
 Call ConfiguraRst(strCadena)
  If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub

End If
  
      N = 1
      Grilla.Clear
      Grilla.Rows = 0
      ' Me.HfdGrilla.Rows = rst.RecordCount - 2
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 1200
           Grilla.ColWidth(1) = 0
           Grilla.ColWidth(2) = 3000
           Grilla.ColWidth(3) = 700
           Grilla.ColWidth(4) = 2500
           Grilla.ColWidth(5) = 1500
           Grilla.ColWidth(6) = 4000
        Next
          
        cabecera = "RUC" & vbTab & "MES" & vbTab & "PERIODO" & vbTab & "AÑO" & vbTab & "DEBE" & vbTab & "HABER" & vbTab & "RAZON"
        Grilla.AddItem cabecera
         For k = 0 To 6
             Grilla.col = k
             Grilla.Row = 0
             Grilla.CellBackColor = &HDFDFE0
        Next k

        rst.MoveFirst
        Acumulado = 0
        For i = 0 To rst.RecordCount - 1
            strCadena = "SELECT SUM(debe),sum(haber) FROM registro_diario_detalle WHERE id_anio='" & rst("anio") & "' AND id_mes='" & Trim(rst("mes")) & "' AND ruc='" & Trim(rst("ruc")) & "'  "
            Call ConfiguraTemporal(strCadena)
            If IsNull(rstTemporal(0)) = True And IsNull(rstTemporal(1)) = True Then
                debes = 0
                habers = 0
            Else
                debes = rstTemporal(0)
                habers = rstTemporal(1)
            End If
            Fila = rst("ruc") & vbTab & rst(1) & vbTab & rst("descripcion") & vbTab & rst(3) & vbTab & Format(debes, "#,##0.00") & vbTab & Format(debes, "#,##0.00") & vbTab & rst("razon")
            Grilla.AddItem Fila
               
                    
               
            Fila = ""
            rst.MoveNext
             
        Next i
  Formulario.TlbAcciones.Buttons(KEY_UPDATE).Enabled = False
  Formulario.TlbAcciones.Buttons(KEY_DELETE).Enabled = False
  Formulario.TlbAcciones.Buttons(KEY_IMPORTAR).Enabled = False
  Formulario.TlbAcciones.Buttons(KEY_EXPORTAR).Enabled = False
  Formulario.TlbAcciones.Buttons("(Ple)").Enabled = False
  Formulario.TlbAcciones.Buttons("(ExportarDia)").Enabled = False
    
  Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"

End Sub

Private Sub txtAnio_Change()
If (Len(Me.txtanio.Text) = 4) Then
    Me.Command1.Enabled = True
Else
    Me.Command1.Enabled = False
End If
End Sub

Private Sub txtAnio_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call actualizar_anio(Me.TxtRuc.Text, Me.txtanio.Text)
End If
End Sub






