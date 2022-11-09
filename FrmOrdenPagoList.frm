VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmOrdenPagoList 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   6975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12015
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   12015
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdBuscar 
      Caption         =   "Buscar"
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   240
      Width           =   1695
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   7560
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenPagoList.frx":0000
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenPagoList.frx":031C
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenPagoList.frx":077C
            Key             =   "(Imprimir)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenPagoList.frx":0809
            Key             =   "(Inicio)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenPagoList.frx":0C69
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenPagoList.frx":0F85
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenPagoList.frx":13E5
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenPagoList.frx":1701
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenPagoList.frx":1B61
            Key             =   "(Red)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenPagoList.frx":1FC1
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenPagoList.frx":28A1
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenPagoList.frx":2BBD
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmOrdenPagoList.frx":2ED9
            Key             =   "(Cancelar)"
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfDetalle 
      Height          =   5895
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   10398
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
   Begin MSComCtl2.DTPicker DtpInicio 
      Height          =   315
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarForeColor=   -2147483635
      CalendarTitleBackColor=   16777215
      CalendarTitleForeColor=   -2147483635
      Format          =   121176065
      CurrentDate     =   37091
   End
   Begin MSComCtl2.DTPicker DtpFinal 
      Height          =   315
      Left            =   2280
      TabIndex        =   3
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   556
      _Version        =   393216
      CalendarBackColor=   16777215
      CalendarForeColor=   -2147483635
      CalendarTitleBackColor=   16777215
      CalendarTitleForeColor=   -2147483635
      Format          =   121176065
      CurrentDate     =   37091
   End
   Begin ComCtl3.CoolBar ClbAcciones1 
      Height          =   4185
      Left            =   10920
      TabIndex        =   4
      Top             =   240
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   7382
      BandCount       =   1
      ForeColor       =   8388608
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      Orientation     =   1
      EmbossPicture   =   -1  'True
      _CBWidth        =   900
      _CBHeight       =   4185
      _Version        =   "6.0.8169"
      Caption1        =   "Acciones"
      Child1          =   "TlbAcciones"
      MinHeight1      =   840
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   780
         Left            =   30
         TabIndex        =   5
         Top             =   420
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   1376
         ButtonWidth     =   1588
         ButtonHeight    =   1376
         Style           =   1
         ImageList       =   "ImgIconos"
         DisabledImageList=   "ImgIconos"
         HotImageList    =   "ImgIconos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "  &Nuevo   "
               Key             =   "(Nuevo)"
               Object.ToolTipText     =   "Nuevo"
               ImageKey        =   "(Nuevo)"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Anular"
               Key             =   "(Anular)"
               ImageKey        =   "(Modificar)"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Eliminar"
               Key             =   "(Eliminar)"
               Object.ToolTipText     =   "Eliminar"
               ImageKey        =   "(Eliminar)"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Salir"
               Key             =   "(Salir)"
               ImageKey        =   "(Salir)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   2
      Height          =   6975
      Left            =   0
      Top             =   0
      Width           =   12015
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "AL"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   2040
      TabIndex        =   6
      Top             =   360
      Width           =   195
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   240
      Top             =   120
      Width           =   10575
   End
End
Attribute VB_Name = "FrmOrdenPagoList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
CenterForm Me
Me.Top = 50
Call llenarGrid(Me.HfDetalle, Me)
End Sub
Private Sub llenarGrid(ByVal Grilla As MSHFlexGrid, ByVal Formulario As Form)
On Error GoTo salir
Dim tTotal As Double
tTotal = 0


strCadena = "SELECT     OrdenPago.idorden, OrdenPago.fecha, Comprobantes.doc_abrev, OrdenPago.doc_cod, OrdenPago.serie," & _
" OrdenPago.numero, OrdenPago.empresa , OrdenPago.num_cheque, OrdenPago.Monto, OrdenPago.glosa " & _
"FROM         OrdenPago INNER JOIN Comprobantes ON OrdenPago.doc_cod = Comprobantes.doc_cod " & _
" WHERE  OrdenPago.doc_cod='0096'  ORDER BY  OrdenPago.idorden DESC"
Call ConfiguraRst(strCadena)
 
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub

End If
  
  N = 1
  
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
      ' Me.HfdGrilla.Rows = rst.RecordCount - 2
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 600
           Grilla.ColWidth(1) = 1000
           Grilla.ColWidth(2) = 1300
           Grilla.ColWidth(3) = 600
           Grilla.ColWidth(4) = 800
           Grilla.ColWidth(5) = 3000
           Grilla.ColWidth(6) = 1200
           Grilla.ColWidth(7) = 1200
           Grilla.ColWidth(8) = 3500
           Grilla.ColWidth(9) = 0
          Next
          
          
         cabecera = "ITEM" & vbTab & "EMISION" & vbTab & "COMPROBANTE" & vbTab & "SERIE" & vbTab & "NUMERO" & vbTab & "ENTIDAD" & vbTab & "CHEQUE" & vbTab & "MONTO" & vbTab & "GLOSA"
         Grilla.AddItem cabecera
         For k = 0 To 9
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 1 To rst.RecordCount
             Fila = Fila & str(rst.RecordCount - i) & vbTab & rst("fecha") & vbTab & rst("doc_abrev") & vbTab & rst("serie") & vbTab & rst("numero") & vbTab & rst("empresa") & vbTab & rst("num_cheque") & vbTab & Format(rst("monto"), "#,##0.00") & vbTab & rst("glosa")
            If (Fila = "") Then
                x = 1
            End If
            
          Grilla.AddItem Fila
           ' If (Trim(rst("anulado")) = "si") Then
            '                For k = 0 To 8
             '                   Grilla.Col = k
              '                  Grilla.Row = i
               '                 Grilla.CellBackColor = &H8080FF
                '            Next k
        'End If
        Fila = ""
        rst.MoveNext
        Next i
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
'On Error GoTo Error
  Select Case Button.key
    Case KEY_NEW
        FrmreciboIngresos.Show
    Case KEY_DELETE
       ' If MsgBox("Esta Seguro de Eliminar Este Comprobante", vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
       ' StrCadena = "DELETE FROM movimiento_caja  WHERE (serie='" & Trim(Me.TxtSerie.Text) & "' AND numero='" & Trim(Me.TxtNumeroDoc.Text) & "' AND doc_cod='" & Trim(Me.DtcTipoDoc.BoundText) & "')"
        'CnBd.Execute (strCadena)
        'Me.TlbAcciones.Buttons(KEY_DELETE).Enabled = False
        'Me.TlbAcciones.Buttons(KEY_ANULAR).Enabled = False
       'End If
    Case KEY_ANULAR
    'If MsgBox("Esta Seguro de Anular Este Comprobante", vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
     '   StrCadena = "UPDATE movimiento_caja SET anulado='si' WHERE (serie='" & Trim(Me.TxtSerie.Text) & "' AND numero='" & Trim(Me.TxtNumeroDoc.Text) & "' AND doc_cod='" & Trim(Me.DtcTipoDoc.BoundText) & "')"
      '  CnBd.Execute (strCadena)
       '  doc = Me.DtcTipoDoc.Text + ":" + Me.TxtSerie.Text + "-" + Me.TxtNumeroDoc.Text
        ' StrCadena = "DELETE mis_cuentas_det WHERE documento='" & Trim(doc) & "'"
        ' CnBd.Execute (strCadena)
        'Me.lblAnulado.Visible = True
        'Me.TlbAcciones.Buttons(KEY_ANULAR).Enabled = False
       'End If
    Case KEY_EXIT
        Unload Me
'Error:
 ' MsgBox Err.Number & " " & Err.Description, vbCritical, MSGERROR
  
  Exit Sub
End Select
End Sub


