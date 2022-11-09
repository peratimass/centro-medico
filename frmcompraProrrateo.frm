VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmcompraProrrateo 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   19005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   19005
   ShowInTaskbar   =   0   'False
   Begin VitekeySoft.ChameleonBtn cmdprocesar 
      Height          =   735
      Left            =   15840
      TabIndex        =   1
      Top             =   6360
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   1296
      BTYPE           =   5
      TX              =   "ACTUALIZAR Y SALIR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   0
      MICON           =   "frmcompraProrrateo.frx":0000
      PICN            =   "frmcompraProrrateo.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdDetalle 
      Height          =   4935
      Left            =   240
      TabIndex        =   0
      Top             =   1320
      Width           =   18495
      _ExtentX        =   32623
      _ExtentY        =   8705
      _Version        =   393216
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      ForeColorSel    =   16777215
      BackColorBkg    =   16777215
      GridColor       =   -2147483635
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
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "COD.UNICO:"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label lblidcompra 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      Caption         =   "F0001"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   1680
      TabIndex        =   10
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label lblfecha 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      Caption         =   "F0001"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   1680
      TabIndex        =   9
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label lblrazonsocial 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      Caption         =   "F0001"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   5400
      TabIndex        =   8
      Top             =   840
      Width           =   5535
   End
   Begin VB.Label lblruc 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      Caption         =   "F0001"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   5400
      TabIndex        =   7
      Top             =   480
      Width           =   2655
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PROVEEDOR :"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4125
      TabIndex        =   6
      Top             =   480
      Width           =   1170
   End
   Begin VB.Label lblnumero 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      Caption         =   "F0001"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   6120
      TabIndex        =   5
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lblserie 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      Caption         =   "F0001"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   5400
      TabIndex        =   4
      Top             =   120
      Width           =   615
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "COMPROBANTE :"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3840
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "FECHA EMISION :"
      BeginProperty Font 
         Name            =   "Cambria"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   18600
      Picture         =   "frmcompraProrrateo.frx":3664
      Top             =   120
      Width           =   240
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   7290
      Left            =   0
      Top             =   0
      Width           =   19005
   End
End
Attribute VB_Name = "frmcompraProrrateo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub llenarGrid_Comprobante(ByVal Grilla As MSHFlexGrid, ByVal id_compra As Double)
 On Error GoTo salir
Dim Total As Double, SUBTOTAL As Double, igv As Single, tpercepcion As Single
Dim prorrateo_imp As String
Dim prorrateo_gas As String


strCadena = "SELECT * FROM movimiento_compra WHERE id_compra='" & id_compra & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   prorrateo_imp = rst("prorrateo_importacion")
   prorrateo_gas = rst("prorrateo_gastos")
Else
  prorrateo_imp = "no"
  prorrateo_gas = "no"
End If

strCadena = "SELECT * FROM view_detalle_compra WHERE id_compra='" & id_compra & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    
   
  Grilla.Rows = 0
  ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            Grilla.ColWidth(0) = 0
            Grilla.ColWidth(1) = 700
            Grilla.ColWidth(2) = 3500
            Grilla.ColWidth(3) = 500
            Grilla.ColWidth(4) = 500
            Grilla.ColWidth(5) = 850
            Grilla.ColWidth(6) = 850
            Grilla.ColWidth(7) = 800
            Grilla.ColWidth(8) = 600
            Grilla.ColWidth(9) = 900
            Grilla.ColWidth(10) = 900
            Grilla.ColWidth(11) = 900
            Grilla.ColWidth(12) = 900
            Grilla.ColWidth(13) = 900
            Grilla.ColWidth(14) = 1000
            Grilla.ColWidth(15) = 1000
            Grilla.ColWidth(16) = 1200
            Grilla.ColWidth(17) = 1200
            Grilla.ColWidth(18) = 1200
    Next
  
 
             
             Fila = "IDTEMPORAL" & vbTab & "CODIGO" & vbTab & "P R O D U C T O" & vbTab & "UND" & vbTab & "CANT" & vbTab & "P.UNIT" & vbTab & "V.NETO" & vbTab & "OTROS" & vbTab & "T.DSC" & vbTab & "ISC" & vbTab & "IVAP" & vbTab & "P.VENTA" & vbTab & "P.COSTO" & vbTab & "INC [NETO]" & vbTab & "INC [%]" & vbTab & "G.VINC" & vbTab & "VALOR VENTA" & vbTab & "IGV" & vbTab & "TOTAL"
             Grilla.AddItem Fila
             For k = 0 To 18
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
             Fila = ""
             cantidad = 0
             unit = 0
             neto = 0
             desS = 0
             desP = 0
             otros = 0
             Tdes = 0
             isc = 0
             igv = 0
             ivap = 0
             pventa = 0
             valor_venta = 0
             in_exonerado = 0
             in_incremento_neto = 0
             in_incremento_neto_gasto = 0
             in_acumulado = 0
             in_retencion = 0
        For i = 0 To rst.RecordCount - 1

             If prorrateo_imp = "si" Then
                pventa = pventa + rst("precio_venta") + rst("incremento_neto")
                in_venta = rst("precio_venta") + rst("incremento_neto")
             Else
                pventa = pventa + rst("precio_venta")
                in_venta = rst("precio_venta")
             End If
             
             If prorrateo_gas = "si" Then
                pventa = pventa + rst("incremento_neto_gasto")
                in_venta = rst("precio_venta") + rst("incremento_neto_gasto") + rst("incremento_neto")
             Else
                pventa = pventa
                in_venta = in_venta
             End If
             
             Fila = rst("id_detalle_compra") & vbTab & rst("id_producto") & vbTab & UCase(rst("nombre_prod")) & vbTab & rst("abreviatura") & vbTab & rst("cantidad") & vbTab & Format(rst("c_unitario"), "###0.00") & vbTab & Format(rst("valor_neto"), "###0.00") & vbTab & Format(rst("otros"), "###0.00") & vbTab & Format(rst("total_descuento"), "###0.00") & vbTab & Format(rst("isc"), "###0.00") & vbTab & Format(rst("ivap"), "###0.00") & vbTab & Format(rst("p_venta"), "###0.00") & vbTab & Format(rst("p_costo"), "###0.00") & vbTab & Format(rst("incremento_neto"), "###0.000") & vbTab & Format(rst("incremento"), "###0.000") & vbTab & Format(rst("incremento_neto_gasto"), "###0.000") & vbTab & Format(rst("valor_venta"), "###0.00") & vbTab & Format(rst("igv"), "###0.00") & vbTab & Format(in_venta, "###0.00")
             Grilla.AddItem Fila
              
                                Grilla.col = 18
                                Grilla.Row = i + 1
                                Grilla.CellBackColor = &HC0C0FF
                           
           
             cantidad = cantidad + rst("cantidad")
             unit = unit + rst("c_unitario")
             neto = neto + rst("valor_neto")
             desS = desS + rst("dsto_soles")
             desP = desP + rst("dsto_procentaje")
             otros = otros + rst("otros")
             Tdes = Tdes + rst("total_descuento")
             isc = isc + rst("isc")
             igv = igv + rst("igv")
             in_retencion = in_retencion + rst("retencion")
             ivap = ivap + rst("ivap")
             in_pventa = in_pventa + rst("p_venta")
             in_pcosto = in_pcosto + rst("p_costo")
             in_acumulado = in_acumulado + in_venta
             in_exonerado = in_exonerado + rst("exonerado")
             valor_venta = valor_venta + rst("valor_venta")
             in_incremento = in_incremento + rst("incremento")
             in_incremento_neto = in_incremento_neto + rst("incremento_neto")
             in_incremento_neto_gasto = in_incremento_neto_gasto + rst("incremento_neto_gasto")
             rst.MoveNext
        Next i

         Fila = "" & vbTab & "" & vbTab & " [      :::::::::::::::: T  O  T  A  L  E  S ::::::::::::::::: ]" & vbTab & "" & vbTab & Format(cantidad, "###0.00") & vbTab & Format(unit, "###0.00") & vbTab & Format(neto, "###0.00") & vbTab & Format(otros, "###0.00") & vbTab & Format(Tdes, "###0.00") & vbTab & Format(isc, "###0.00") & vbTab & Format(ivap, "###0.00") & vbTab & Format(in_pventa, "###0.00") & vbTab & Format(in_pcosto, "###0.00") & vbTab & Format(in_incremento_neto, "###0.000") & vbTab & Format(in_incremento, "###0.000") & vbTab & Format(in_incremento_neto_gasto, "###0.000") & vbTab & Format(valor_venta, "###0.00") & vbTab & Format(igv, "###0.00") & vbTab & Format(in_acumulado, "###0.00")
         Grilla.AddItem Fila
                        For k = 0 To 18
                                Grilla.col = k
                                Grilla.Row = i + 1
                                Grilla.CellBackColor = &HC0C0FF
                            Next k
                            
End If
Me.LblCantidad.Caption = Trim(rst.RecordCount)


  Exit Sub

salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"

End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 150
 



End Sub
