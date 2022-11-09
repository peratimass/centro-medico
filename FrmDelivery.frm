VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmDelivery 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5925
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12270
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   12270
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   3360
      Top             =   2880
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
            Picture         =   "FrmDelivery.frx":0000
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDelivery.frx":0454
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDelivery.frx":0774
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDelivery.frx":0BC8
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDelivery.frx":101C
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDelivery.frx":133C
            Key             =   "(Cancelar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDelivery.frx":165C
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDelivery.frx":197C
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDelivery.frx":1C9C
            Key             =   "(Buscar)"
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgLinea 
      Height          =   5295
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   9340
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
      Height          =   2865
      Left            =   11205
      TabIndex        =   2
      Top             =   450
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   5054
      BandCount       =   1
      ForeColor       =   8388608
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      Orientation     =   1
      EmbossPicture   =   -1  'True
      _CBWidth        =   900
      _CBHeight       =   2865
      _Version        =   "6.0.8169"
      Caption1        =   "Acciones"
      Child1          =   "TlbAcciones"
      MinHeight1      =   840
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   2340
         Left            =   30
         TabIndex        =   3
         Top             =   420
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   4128
         ButtonWidth     =   1217
         ButtonHeight    =   1376
         Style           =   1
         ImageList       =   "ImgIconos"
         DisabledImageList=   "ImgIconos"
         HotImageList    =   "ImgIconos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Finalizar"
               Key             =   "(Aceptar)"
               Object.ToolTipText     =   "Eliminar"
               ImageKey        =   "(Aceptar)"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Salir"
               Key             =   "(Salir)"
               ImageKey        =   "(Salir)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      DrawMode        =   1  'Blackness
      FillColor       =   &H00FFFFFF&
      Height          =   5925
      Left            =   0
      Top             =   0
      Width           =   12270
   End
   Begin VB.Label LblEmpresa 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DELIVERY'S PENDIENTES"
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
      Left            =   330
      TabIndex        =   0
      Top             =   120
      Width           =   1995
   End
End
Attribute VB_Name = "FrmDelivery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede
Public Sub Form_Activate()

End Sub
Public Sub LlenarDelivery(ByVal Grilla As MSHFlexGrid)
On Error GoTo SALIR
Dim tTotal As Double
'strCadena = "SELECT DocumentoVenta.idVenta,  DocumentoVenta.dEmisionVenta, (Comprobantes.doc_abrev +':'+ " & _
"DocumentoVenta.sSerie +'-'+ DocumentoVenta.cDocumentoVenta) AS Numero, DocumentoVenta.Persona, DocumentoVenta.nTotalVenta,DocumentoVenta.monto_pagado " & _
"FROM DocumentoVenta INNER JOIN Comprobantes ON DocumentoVenta.doc_cod = Comprobantes.doc_cod WHERE DocumentoVenta.Ruc='" & KEY_RUC & "' AND delivery='V' AND dEmisionVenta='" & KEY_FECHA & "' "

strCadena = "SELECT V.id_venta,V.fecha_emision,CONCAT(C.doc_abrev,':',V.serie,'-',V.numero)as numero,ncliente,total,monto_pago,monto_vuelto FROM movimiento_venta V,comprobantes C WHERE V.id_doc=C.id_doc AND V.ruc='" & KEY_RUC & "' AND id_delivery='si' AND anulado='no' AND id_vendedor='" & KEY_USUARIO & "'"
Call ConfiguraRst(strCadena)

If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub

End If
   Grilla.Clear
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1200
           Grilla.ColWidth(2) = 2200
           Grilla.ColWidth(3) = 2900
           Grilla.ColWidth(4) = 1300
           Grilla.ColWidth(5) = 1300
           Grilla.ColWidth(6) = 1300
         Next
        cabecera = "CODIGO" & vbTab & "EMISION" & vbTab & "DOCUMENTO" & vbTab & "CLIENTE" & vbTab & "TOTAL" & vbTab & "PAGADO" & vbTab & "VUELTO"
        Grilla.AddItem cabecera
         For k = 0 To 6
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        
        For i = 0 To rst.RecordCount - 1
            Fila = Fila & rst("id_venta") & vbTab & rst("fecha_emision") & vbTab & rst("numero") & vbTab & rst("ncliente") & vbTab & Format(rst("total"), "#,##0.00") & vbTab & Format(rst("monto_pago"), "#,##0.00") & vbTab & Format(rst("monto_vuelto"), "#,##0.00")
            Grilla.AddItem Fila
            Fila = ""
            rst.MoveNext
        Next i
     
      
  Exit Sub
SALIR: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    FrmVentas.Nuevo
    Unload Me
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 500
Call LlenarDelivery(Me.HfgLinea)
End Sub

Private Sub HfgLinea_Click()
If HfgLinea.Row > 0 Then
      
    TlbAcciones.Buttons(KEY_OK).Enabled = True
  End If
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
 Select Case Button.Key
    
    Case KEY_OK
        Procedencia = Modificar
        FrmSeguridad.Show
       
     
    Case KEY_EXIT
        FrmVentas.Nuevo
        Unload Me
        
  End Select
End Sub

Private Sub TxtLinea_Change()
'StrCadena = "SELECT id_producto as Codigo,descripcion as Descripcion FROM Producto_Recomendado WHERE descripcion LIKE '%" & Trim(Me.TxtLinea.Text) & "%'ORDER BY sDescripcion ASC"
 ' Call llenarGridME(Me.HfgUnidad, Me)

End Sub

