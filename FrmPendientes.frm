VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmPendientes 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9120
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13935
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   13935
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtNumero 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   5350
      Width           =   2535
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfGuardado 
      Height          =   4695
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   8281
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      GridColor       =   0
      FocusRect       =   0
      GridLines       =   2
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   10440
      Top             =   5160
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
            Picture         =   "FrmPendientes.frx":0000
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPendientes.frx":0454
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPendientes.frx":0774
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPendientes.frx":0BC8
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPendientes.frx":101C
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPendientes.frx":133C
            Key             =   "(Cancelar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPendientes.frx":165C
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPendientes.frx":197C
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmPendientes.frx":1C9C
            Key             =   "(Buscar)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   3225
      Left            =   12765
      TabIndex        =   2
      Top             =   600
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   5689
      BandCount       =   1
      ForeColor       =   8388608
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      Orientation     =   1
      EmbossShadow    =   -2147483628
      EmbossPicture   =   -1  'True
      _CBWidth        =   900
      _CBHeight       =   3225
      _Version        =   "6.7.9782"
      Caption1        =   "Acciones"
      Child1          =   "TlbAcciones"
      MinHeight1      =   840
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   780
         Left            =   30
         TabIndex        =   3
         Top             =   345
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   1376
         ButtonWidth     =   1349
         ButtonHeight    =   1376
         Style           =   1
         ImageList       =   "ImgIconos"
         DisabledImageList=   "ImgIconos"
         HotImageList    =   "ImgIconos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Cancelar"
               Key             =   "(Cancelar)"
               Object.ToolTipText     =   "Nuevo"
               ImageKey        =   "(Aceptar)"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Anular"
               Key             =   "(Anular)"
               Object.ToolTipText     =   "Anular"
               ImageKey        =   "(Modificar)"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Salir"
               Key             =   "(Salir)"
               ImageKey        =   "(Salir)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgDetalle 
      Height          =   3135
      Left            =   240
      TabIndex        =   6
      Top             =   5880
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   5530
      _Version        =   393216
      BackColor       =   16777215
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      GridColor       =   0
      FocusRect       =   0
      GridLines       =   2
      SelectionMode   =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Buscar:"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   570
      TabIndex        =   4
      Top             =   5280
      Width           =   975
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      Height          =   495
      Left            =   240
      Top             =   5280
      Width           =   12375
   End
   Begin VB.Shape Shape1 
      Height          =   9120
      Left            =   0
      Top             =   0
      Width           =   13935
   End
   Begin VB.Label LblEmpresa 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Documentos Pendientes."
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   270
      TabIndex        =   1
      Top             =   120
      Width           =   3075
   End
End
Attribute VB_Name = "FrmPendientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede
Private Sub Form_Activate()

  Me.HfGuardado.SetFocus
End Sub


Sub llenarGridME(ByVal Grilla As MSHFlexGrid, ByVal Formulario As Form)
On Error GoTo salir
  Call ConfiguraRst(strCadena)
  Grilla.Clear
  Grilla.rows = rst.RecordCount + 1
  Set Grilla.Recordset = rst
  Grilla.ColWidth(0) = 0
  Grilla.ColWidth(1) = 0
  Grilla.ColWidth(2) = 0
  Grilla.ColWidth(3) = 0
  Grilla.ColWidth(4) = 2650
  Grilla.ColWidth(5) = 3000
  Grilla.ColWidth(6) = 1200
  Grilla.ColWidth(7) = 1200
  Grilla.ColWidth(8) = 1000
  Grilla.ColWidth(9) = 2500
  
  'Call DarFormato(Grilla, 6)
  'Call DarFormato(Grilla, 7)
  'Call DarFormato(Grilla, 8)
  
   Formulario.TlbAcciones.Buttons(KEY_CANCELAR).Enabled = False
    Formulario.TlbAcciones.Buttons(KEY_ANULAR).Enabled = False
  Exit Sub
salir: MsgBox "Ocurrio un Error Intentelo Nuevamente", vbInformation, "Mensaje para el usuario"
End Sub
Sub llenarGridDetalle(ByVal Grilla As MSHFlexGrid, ByVal Formulario As Form)
On Error GoTo salir


  Call ConfiguraRst(strCadena)
  Grilla.Clear
  Grilla.rows = rst.RecordCount
  Set Grilla.Recordset = rst
  Grilla.ColWidth(0) = 700
  Grilla.ColWidth(1) = 4300
  Grilla.ColWidth(2) = 700
  Grilla.ColWidth(3) = 1300
  Grilla.Refresh
   Me.HfGuardado.SetFocus
  Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If (KeyCode = 27) Then
    Procedencia = Neutro
    Unload Me
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 50
  Call llenar
End Sub
Public Sub llenar()
Dim StrOperador As String

If FrmReporteRecaudacionDiaria.Procedencia = buscar Then
    If FrmReporteRecaudacionDiaria.chkOperador.Value = 1 Then
         StrOperador = Replace(FrmReporteRecaudacionDiaria.DtcOperador.BoundText, "'", "''")
    End If
    strCadena = "SELECT DocumentoVenta.id_documentoventa,DocumentoVenta.cDocumentoVenta,DocumentoVenta.sSerie,DocumentoVenta.doc_cod,(Comprobantes.doc_abrev+':'+ DocumentoVenta.sSerie +'-'+ DocumentoVenta.cDocumentoVenta) as COMPROBANTE, DocumentoVenta.Persona,DocumentoVenta.nTotalVenta as TOTAL ," & _
"DocumentoVenta.monto_pagado as PAGADO, DocumentoVenta.montoenvase AS ENVASE,DocumentoVenta.Observacion as DETALLE FROM DocumentoVenta INNER JOIN Comprobantes ON DocumentoVenta.doc_cod = Comprobantes.doc_cod " & _
"WHERE envase='no' AND Anulado<>'V'"

Else
strCadena = "SELECT DocumentoVenta.id_documentoventa,DocumentoVenta.cDocumentoVenta,DocumentoVenta.sSerie,DocumentoVenta.doc_cod,(Comprobantes.doc_abrev+':'+ DocumentoVenta.sSerie +'-'+ DocumentoVenta.cDocumentoVenta) as COMPROBANTE, DocumentoVenta.Persona,DocumentoVenta.nTotalVenta as TOTAL ," & _
"DocumentoVenta.monto_pagado as PAGADO, DocumentoVenta.montoenvase AS ENVASE,DocumentoVenta.Observacion as DETALLE FROM DocumentoVenta INNER JOIN Comprobantes ON DocumentoVenta.doc_cod = Comprobantes.doc_cod " & _
"WHERE  envase='no' AND Anulado<>'V'"
End If
Call llenarGridME(Me.HfGuardado, Me)

End Sub


Private Sub HfGuardado_Click()
If Me.HfGuardado.rows > 0 Then
    Me.TlbAcciones.Buttons(KEY_CANCELAR).Enabled = True
    Me.TlbAcciones.Buttons(KEY_ANULAR).Enabled = True
End If
End Sub



Private Sub HfGuardado_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 And FrmVentas.Procedencia = buscar Then
    Me.HfGuardado.col = 0
    strCadena = "UPDATE temporal_ventas SET guardado='F' WHERE cDocumentoVenta='" & Trim(Me.HfGuardado.Text) & "' AND id_usuario='" & Trim(KEY_USUARIO) & "'"
    Call EjecutaRST(strCadena)
    Set RstEjecuta = Nothing
    strCadena = "DELETE FROM Temporal_Venta_Guardado WHERE id_guardado='" & Trim(Me.HfGuardado.Text) & "' and id_usuario='" & Trim(KEY_USUARIO) & "'"
    Call EjecutaRST(strCadena)
    Set RstEjecuta = Nothing
    
     FrmVentas.TxtCodCliente.Text = "00004"
     FrmVentas.txtdireccion.Text = "TARAPOTO"
     
    
    Me.HfGuardado.col = 1
    FrmVentas.TxtCliente.Text = Me.HfGuardado.Text
    Unload Me
End If
End Sub


Private Sub HfGuardado_SelChange()
strCadena = "SELECT   *  FROM DocumentoVenta WHERE (DocumentoVenta.id_documentoventa='" & Trim(Me.HfGuardado.TextMatrix(Me.HfGuardado.Row, 0)) & "' AND DocumentoVenta.cDocumentoVenta='" & Trim(Me.HfGuardado.TextMatrix(Me.HfGuardado.Row, 1)) & "' " & _
" AND DocumentoVenta.doc_cod='" & Trim(Me.HfGuardado.TextMatrix(Me.HfGuardado.Row, 3)) & "' AND DocumentoVenta.sSerie='" & Trim(Me.HfGuardado.TextMatrix(Me.HfGuardado.Row, 2)) & "')"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    If Len(rst("Observacion")) > 0 Then
       '  Me.TxtDetalle.Text = rst("Observacion")
    Else
       ' Me.TxtDetalle.Text = ""
    End If
End If
Set rst = Nothing
strCadena = "SELECT   Detalle_DocumentoVenta.cProducto as Codigo, Producto.DescripcionProducto as Descripcion, Unidad.sAbreviatura as UND, Detalle_DocumentoVenta.Precio, " & _
"Detalle_DocumentoVenta.cantidad as Cantidad , Detalle_DocumentoVenta.total as Total  FROM DocumentoVenta INNER JOIN Detalle_DocumentoVenta ON " & _
"DocumentoVenta.id_documentoventa = Detalle_DocumentoVenta.id_documentoventa AND DocumentoVenta.cDocumentoVenta = Detalle_DocumentoVenta.cDocumentoVenta " & _
"AND DocumentoVenta.doc_cod = Detalle_DocumentoVenta.doc_cod AND    DocumentoVenta.Alm_cod = Detalle_DocumentoVenta.Alm_Cod AND " & _
"DocumentoVenta.sSerie = Detalle_DocumentoVenta.sSerie INNER JOIN Producto ON Detalle_DocumentoVenta.cProducto = Producto.cProducto INNER JOIN " & _
"Unidad ON Producto.cUnidad = Unidad.cUnidad WHERE (DocumentoVenta.id_documentoventa='" & Trim(Me.HfGuardado.TextMatrix(Me.HfGuardado.Row, 0)) & "' AND DocumentoVenta.cDocumentoVenta='" & Trim(Me.HfGuardado.TextMatrix(Me.HfGuardado.Row, 1)) & "' " & _
" AND DocumentoVenta.doc_cod='" & Trim(Me.HfGuardado.TextMatrix(Me.HfGuardado.Row, 3)) & "' AND DocumentoVenta.sSerie='" & Trim(Me.HfGuardado.TextMatrix(Me.HfGuardado.Row, 2)) & "')"

  Call llenarGridDetalle(Me.HfgDetalle, Me)
   
End Sub


Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.key
    Case KEY_CANCELAR
        Procedencia = Modificar
        FrmSeguridad.Show
        
    Case KEY_ANULAR
        If MsgBox(MSGANULAR, vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
                 Procedencia = anular
                 FrmSeguridad.Show
                Exit Sub
        End If
    Case KEY_EXIT
        Procedencia = Neutro
        Unload Me
  End Select
End Sub

Private Sub TxtNumero_Change()

           ' Me.HfGuardado.Col = 1
           ' Numero = Me.HfGuardado.Text
            
            
            
            
strCadena = "SELECT DocumentoVenta.id_documentoventa,DocumentoVenta.cDocumentoVenta,DocumentoVenta.sSerie,DocumentoVenta.doc_cod,(Comprobantes.doc_abrev+':'+ DocumentoVenta.sSerie +'-'+ DocumentoVenta.cDocumentoVenta) as COMPROBANTE, DocumentoVenta.Persona,DocumentoVenta.nTotalVenta as TOTAL ," & _
"DocumentoVenta.monto_pagado as PAGADO, DocumentoVenta.montoenvase AS ENVASE,DocumentoVenta.Observacion as DETALLE FROM DocumentoVenta INNER JOIN Comprobantes ON DocumentoVenta.doc_cod = Comprobantes.doc_cod " & _
"WHERE  envase='no' AND Anulado<>'V' AND DocumentoVenta.cDocumentoVenta LIKE '%" & Trim(Me.txtNumero.Text) & "%' "
Call llenarGridME(Me.HfGuardado, Me)
End Sub
