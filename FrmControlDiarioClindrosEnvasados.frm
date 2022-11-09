VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmControlDiarioClindrosEnvasados 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7335
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13950
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   13950
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox tXTgLOSA 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   6720
      Width           =   10215
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   3360
      Top             =   2670
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
            Picture         =   "FrmControlDiarioClindrosEnvasados.frx":0000
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmControlDiarioClindrosEnvasados.frx":0454
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmControlDiarioClindrosEnvasados.frx":0774
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmControlDiarioClindrosEnvasados.frx":0BC8
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmControlDiarioClindrosEnvasados.frx":101C
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmControlDiarioClindrosEnvasados.frx":133C
            Key             =   "(Cancelar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmControlDiarioClindrosEnvasados.frx":165C
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmControlDiarioClindrosEnvasados.frx":197C
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmControlDiarioClindrosEnvasados.frx":1C9C
            Key             =   "(Buscar)"
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgDetalle 
      Height          =   6135
      Left            =   240
      TabIndex        =   1
      Top             =   390
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   10821
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
      Height          =   5745
      Left            =   12885
      TabIndex        =   2
      Top             =   360
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   10134
      BandCount       =   1
      ForeColor       =   8388608
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      Orientation     =   1
      EmbossPicture   =   -1  'True
      _CBWidth        =   900
      _CBHeight       =   5745
      _Version        =   "6.0.8169"
      Caption1        =   "Acciones"
      Child1          =   "TlbAcciones"
      MinHeight1      =   840
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   7020
         Left            =   30
         TabIndex        =   3
         Top             =   345
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   12383
         ButtonWidth     =   1535
         ButtonHeight    =   1376
         Style           =   1
         ImageList       =   "ImgIconos"
         DisabledImageList=   "ImgIconos"
         HotImageList    =   "ImgIconos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   9
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
               Caption         =   "Detalles"
               Key             =   "(Modificar)"
               Object.ToolTipText     =   "Modificar"
               ImageKey        =   "(Modificar)"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Declarar"
               Key             =   "(Declarar)"
               ImageKey        =   "(Aceptar)"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Eliminar"
               Key             =   "(Eliminar)"
               Object.ToolTipText     =   "Eliminar"
               ImageKey        =   "(Eliminar)"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Salir"
               Key             =   "(Salir)"
               ImageKey        =   "(Salir)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   7335
      Left            =   0
      Top             =   0
      Width           =   13950
   End
   Begin VB.Label LblEmpresa 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CONTROL DIARIO DE CILINDROS ENSADOS"
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
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   3975
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   3  'Dot
      Height          =   615
      Left            =   480
      Top             =   6600
      Width           =   12375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GLOSA :"
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
      Left            =   720
      TabIndex        =   4
      Top             =   6720
      Width           =   795
   End
End
Attribute VB_Name = "FrmControlDiarioClindrosEnvasados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede
Private Sub Form_Activate()
CenterForm Me
Me.Top = 500
End Sub
Public Sub actualizar()

Call llenarGrid(Me.HfgDetalle, Me)
End Sub

Private Sub HfgMarcas_Click()
If HfgMarcas.Row > 0 Then
    TlbAcciones.Buttons(KEY_UPDATE).Enabled = True
    TlbAcciones.Buttons(KEY_DELETE).Enabled = True
  End If
End Sub

Private Sub Form_Load()
Call actualizar
End Sub

Private Sub HfgDetalle_SelChange()
If Val(Me.HfgDetalle.TextMatrix(Me.HfgDetalle.Row, 0)) > 0 Then
    strCadena = "SELECT * FROM SolicitudViaticos WHERE id_Solicitud='" & Val(Me.HfgDetalle.TextMatrix(Me.HfgDetalle.Row, 0)) & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        Me.tXTgLOSA.Text = rst("glosa")
    End If
    Set rst = Nothing
End If
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Key
    Case KEY_NEW
      Procedencia = Nuevo
      FrmSolicitudViaticosDet.Show
     Case KEY_UPDATE
        
       ' strCadena = "SELECT * FROM movimiento_caja WHERE tipo_trans='E' AND  "
        
         strCadena = "SELECT     dbo.Comprobantes.doc_abrev, dbo.movimiento_caja.serie, dbo.movimiento_caja.numero, dbo.movimiento_caja.cPersona, " & _
        "dbo.movimiento_caja.descripcion_per, dbo.Persona.sDireccionCliente1, dbo.Persona.Per_Ruc, dbo.movimiento_caja.fecha_valor," & _
        "dbo.movimiento_caja.cambio , dbo.movimiento_caja.glosa, dbo.centro_costos.descripcion, dbo.movimiento_caja.Monto,dbo.movimiento_caja.monto_letras " & _
        "FROM dbo.movimiento_caja INNER JOIN dbo.Comprobantes ON dbo.movimiento_caja.doc_cod = dbo.Comprobantes.doc_cod INNER JOIN " & _
        "dbo.centro_costos ON dbo.movimiento_caja.id_costo = dbo.centro_costos.id_costo INNER JOIN " & _
        "dbo.Persona ON dbo.movimiento_caja.cPersona = dbo.Persona.cPersona WHERE codigo='" & idRecibo & "'"
        Call ConfiguraRst(strCadena)
        Ans = ShowMultiReport(rst, "RptReciboCaja", , App.Path + "\Reportes\")
    Case "(Declarar)"
        FrmSolicitudViaticosDeclarar.Show
   Case KEY_DELETE
      If MsgBox(MSGELIMINAR + Chr(13) + "Se Eliminaran los cheques Relacionados a esta Chequera", vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
        
    End If
    Case KEY_EXIT
        Unload Me
  End Select
End Sub
Private Sub llenarGrid(ByVal Grilla As MSHFlexGrid, ByVal formulario As Form)
'On Error GoTo salir
Dim tTotal As Double
Dim rstB As New ADODB.Recordset
strCadena = "SELECT dbo.SolicitudViaticos.id_Solicitud, dbo.SolicitudViaticos.fechaInicio, dbo.SolicitudViaticos.fechaFinal, dbo.Persona.NombrePersona, " & _
" dbo.SolicitudViaticos.idOrden, dbo.SolicitudViaticos.monto, dbo.SolicitudViaticos.montodeclarado as declarado, dbo.SolicitudViaticos.vuelto," & _
" dbo.SolicitudViaticos.estado FROM dbo.SolicitudViaticos INNER JOIN  dbo.Persona ON dbo.SolicitudViaticos.cPersona = dbo.Persona.cPersona WHERE SolicitudViaticos.Ruc='" & KEY_RUC & "' ORDER BY SolicitudViaticos.estado "
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub

End If
   n = 1
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
      ' Me.HfdGrilla.Rows = rst.RecordCount - 2
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 1100
           Grilla.ColWidth(1) = 2100
           Grilla.ColWidth(2) = 3200
           Grilla.ColWidth(3) = 2400
           Grilla.ColWidth(4) = 1100
           Grilla.ColWidth(5) = 1100
           Grilla.ColWidth(6) = 1100
           
        Next
        cabecera = "SOLICITUD" & vbTab & "PERIODO" & vbTab & "EMPLEADO" & vbTab & "ORDEN COMPRA" & vbTab & " MONTO" & vbTab & "DECLARADO" & vbTab & "VUELTO"
        Grilla.AddItem cabecera
         For k = 0 To 6
                                Grilla.Col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            If rst("idOrden") > 0 Then
                strCadena = "SELECT dbo.Comprobantes.doc_abrev, dbo.OrdenCompra.serie, dbo.OrdenCompra.numero FROM dbo.OrdenCompra INNER JOIN " & _
                " dbo.Comprobantes ON dbo.OrdenCompra.doc_cod = dbo.Comprobantes.doc_cod WHERE idOrden='" & rst("idOrden") & "'"
                Call ConfiguraTemporal(strCadena)
                If rstTemporal.RecordCount > 0 Then
                    orden = rstTemporal("doc_abrev") + ":" + rstTemporal("serie") + ":" + rstTemporal("numero")
                Else
                    orden = "-------"
                End If
                Set rstTemporal = Nothing
            Else
                orden = "-------"
            End If
          
            strCadena = "SELECT sum(total) FROM SolicitudViaticosComprobanteDetalle WHERE id_solicitud='" & Val(rst("id_solicitud")) & "'"
            Call ConfiguraTemporal(strCadena)
            If IsNull(rstTemporal(0)) = True Then
                mdeclarado = 0
            Else
                mdeclarado = rstTemporal(0)
            End If
            Set rstTemporal = Nothing
            
            fila = fila & rst("id_Solicitud") & vbTab & Str(rst("fechaInicio")) + " - " + Str(rst("fechaFinal")) & vbTab & rst("NombrePersona") & vbTab & orden & vbTab & Format(rst("Monto"), "#,##0.00") & vbTab & Format(mdeclarado, "#,##0.00") & vbTab & Format(rst("vuelto"), "#,##0.00")
            If (fila = "") Then
                x = 1
            End If
          Grilla.AddItem fila
          If rst("estado") = "P" Then
                        For k = 0 To 6
                                Grilla.Col = k
                                Grilla.Row = i + 1
                                Grilla.CellBackColor = &H8080FF
                            Next k
          End If
          
        fila = ""
        rst.MoveNext
             
        Next i
    
    
    
    
 ' Exit Sub
'salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"


End Sub











