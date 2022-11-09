VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmProductoTransformaciones 
   BorderStyle     =   0  'None
   Caption         =   "Transformaciones"
   ClientHeight    =   5715
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14295
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   14295
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Equivalencia"
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   240
      TabIndex        =   28
      Top             =   3600
      Width           =   11535
      Begin VB.TextBox txtCantidad 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1665
         TabIndex        =   35
         Top             =   840
         Width           =   855
      End
      Begin VB.TextBox txtCantA 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         TabIndex        =   30
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtCantB 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4440
         TabIndex        =   29
         Top             =   360
         Width           =   975
      End
      Begin VitekeySoft.ChameleonBtn cmdNuevo 
         Height          =   855
         Left            =   8160
         TabIndex        =   37
         Top             =   480
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1508
         BTYPE           =   5
         TX              =   "NUEVO"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmProductoTransformaciones.frx":0000
         PICN            =   "FrmProductoTransformaciones.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdProcesar 
         Height          =   855
         Left            =   9240
         TabIndex        =   38
         Top             =   480
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1508
         BTYPE           =   5
         TX              =   "PROCESAR"
         ENAB            =   0   'False
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmProductoTransformaciones.frx":046E
         PICN            =   "FrmProductoTransformaciones.frx":048A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VitekeySoft.ChameleonBtn cmdSalir 
         Height          =   855
         Left            =   10320
         TabIndex        =   39
         Top             =   480
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   1508
         BTYPE           =   5
         TX              =   "SALIR"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   -1  'True
         BCOL            =   16777215
         BCOLO           =   16777215
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "FrmProductoTransformaciones.frx":3AD2
         PICN            =   "FrmProductoTransformaciones.frx":3AEE
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   2
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CANTIDAD"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   525
         TabIndex        =   36
         Top             =   840
         Width           =   795
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         Height          =   495
         Left            =   360
         Top             =   720
         Width           =   5055
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "="
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   2775
         TabIndex        =   34
         Top             =   360
         Width           =   135
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PRODUCTO ""B"""
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   3135
         TabIndex        =   33
         Top             =   360
         Width           =   1125
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PRODUCTO ""A"""
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   375
         TabIndex        =   32
         Top             =   360
         Width           =   1125
      End
   End
   Begin VB.TextBox TxtVentaB 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   10680
      TabIndex        =   19
      Top             =   2760
      Width           =   855
   End
   Begin VB.TextBox txtCostoB 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   9790
      Locked          =   -1  'True
      TabIndex        =   18
      Top             =   2760
      Width           =   855
   End
   Begin VB.TextBox txtDescripcionB 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3525
      TabIndex        =   17
      Top             =   2760
      Width           =   4335
   End
   Begin VB.TextBox TxtCodigobarraB 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1920
      TabIndex        =   16
      Top             =   2760
      Width           =   1575
   End
   Begin VB.TextBox txtcodigo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1920
      TabIndex        =   4
      Top             =   1560
      Width           =   1575
   End
   Begin VB.TextBox TxtProducto 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3525
      TabIndex        =   3
      Top             =   1560
      Width           =   4455
   End
   Begin VB.TextBox TxtTranformacion 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   9840
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   360
      Width           =   1695
   End
   Begin VB.TextBox txtCosto 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9790
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1560
      Width           =   855
   End
   Begin VB.TextBox TxtPVenta 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10680
      TabIndex        =   0
      Top             =   1560
      Width           =   855
   End
   Begin MSDataListLib.DataCombo DtcAlmacen 
      Height          =   315
      Left            =   1230
      TabIndex        =   5
      Top             =   480
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   8388608
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TRANSFORMACION :"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   7335
      TabIndex        =   40
      Top             =   360
      Width           =   2025
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUCTO ""B"""
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   495
      TabIndex        =   31
      Top             =   2760
      Width           =   1125
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "STOCK"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   8040
      TabIndex        =   27
      Top             =   2475
      Width           =   525
   End
   Begin VB.Label LblStockB 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   8040
      TabIndex        =   26
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label lblUndB 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   8920
      TabIndex        =   25
      Top             =   2760
      Width           =   855
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UND"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   8985
      TabIndex        =   24
      Top             =   2475
      Width           =   345
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P.COSTO"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   9795
      TabIndex        =   23
      Top             =   2475
      Width           =   705
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DESCRIPCION PRODUCTO"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   3630
      TabIndex        =   22
      Top             =   2475
      Width           =   1935
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   2040
      TabIndex        =   21
      Top             =   2475
      Width           =   645
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P.VENTA"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   10740
      TabIndex        =   20
      Top             =   2475
      Width           =   645
   End
   Begin VB.Label Label19 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P.VENTA"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   10725
      TabIndex        =   15
      Top             =   1275
      Width           =   675
   End
   Begin VB.Label Label18 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   2040
      TabIndex        =   14
      Top             =   1275
      Width           =   645
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DESCRIPCION PRODUCTO"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   3615
      TabIndex        =   13
      Top             =   1275
      Width           =   1965
   End
   Begin VB.Label Label16 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P.COSTO"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   9810
      TabIndex        =   12
      Top             =   1275
      Width           =   675
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "UND"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   8970
      TabIndex        =   11
      Top             =   1275
      Width           =   375
   End
   Begin VB.Label lblUnidad 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   8925
      TabIndex        =   10
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SUCURSAL:"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   345
      TabIndex        =   9
      Top             =   480
      Width           =   765
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUCTO ""A"""
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   465
      TabIndex        =   8
      Top             =   1560
      Width           =   1185
   End
   Begin VB.Label lblStock 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   8040
      TabIndex        =   7
      Top             =   1560
      Width           =   855
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "STOCK"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   8055
      TabIndex        =   6
      Top             =   1275
      Width           =   495
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      Height          =   855
      Left            =   240
      Top             =   1200
      Width           =   11535
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0FF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      BorderStyle     =   6  'Inside Solid
      Height          =   855
      Left            =   240
      Top             =   2400
      Width           =   11535
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00DFDFE0&
      BorderColor     =   &H00E0E0E0&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   5715
      Left            =   0
      Top             =   0
      Width           =   14295
   End
End
Attribute VB_Name = "FrmProductoTransformaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Procedencia As EnumProcede
Public cod_producto As String
Dim strTransformacion As String
Dim costo_total As Single
Dim stock_total As Single
Public prodB As Boolean
Public prodA As Boolean

Private Sub ClbAcciones_HeightChanged(ByVal NewHeight As Single)

End Sub

Private Sub cmdNuevo_Click()

    Me.txtcodigo.Text = ""
  Me.Txtproducto.Text = ""
  Me.TxtCodigobarraB.Text = ""
  Me.txtDescripcionB.Text = ""
  Me.lblStock.Caption = ""
  Me.LblStockB.Caption = ""
  Me.txtcosto.Text = ""
  Me.txtCostoB.Text = ""
  Me.TxtPVenta.Text = ""
  Me.TxtVentaB.Text = ""
  Me.txtCantA.Text = ""
  Me.txtCantB.Text = ""
  Me.txtCantidad.Text = ""

End Sub

Private Sub cmdProcesar_Click()
Dim in_costo_acum As Double

strCadena = "SELECT * FROM producto WHERE id_producto='" & Trim(Me.txtcodigo.Text) & "' and   ruc='" & KEY_RUC & "' "
Call ConfiguraRstA(strCadena)
If rstA.RecordCount > 0 Then
   rstA.MoveFirst
   in_costo_acum = 0
   
   
   For i = 0 To rstA.RecordCount - 1
        
        strCadena = "SELECT numero,serie FROM movimiento_compra WHERE id_doc='0090' and id_alm='" & Me.DtcAlmacen.BoundText & "' and ruc='" & KEY_RUC & "' ORDER BY numero DESC LIMIT 1"
        Call ConfiguraRstK(strCadena)
       If rstK.RecordCount > 0 Then
              in_numero = Format(Val(rstK("numero")) + 1, "00000000")
              in_serie = rstK("serie")
       Else
             
             strCadena = "SELECT * FROM almacen_comprobante WHERE ruc='" & KEY_RUC & "' and id_alm='" & Me.DtcAlmacen.BoundText & "' and id_doc='0090' LIMIT 1"
             Call ConfiguraRstlocal(strCadena)
             If rstLocal.RecordCount > 0 Then
               in_serie = rstLocal("serie")
               in_numero = formato_item(rstLocal("numero"), 8)
             Else
                MsgBox "CREE EL COMPROBANTE SALIDA A ALMACEN" + Chr(13) + "PARA ESTA SUCURSAL", vbInformation
                Exit Sub
             End If
             
        End If
        
        in_cta_compra = KEY_CTA_COMPRA_SOLES
        in_costo = get_costo_ultimo(rstA("id_producto"), KEY_FECHA)
        in_total = in_costo * Val(Me.txtCantA.Text) * Val(Me.txtCantidad.Text)
        in_costo_acum = in_costo_acum + in_total
        
        strCadena = "call P_insert_compra_ultimate('0090','" & Me.DtcAlmacen.BoundText & "','" & Format(KEY_FECHA, "YYYY-mm-dd") & "','" & Format(KEY_FECHA, "YYYY-mm-dd") & "','02'," & _
        "'03','--','00001','" & formato_item(Month(KEY_FECHA), 2) & "','" & Year(KEY_FECHA) & "','" & in_serie & "'," & _
        "'" & in_numero & "','6','" & KEY_RUC & "','" & KEY_EMPRESA & "','" & KEY_CAMBIO_VENTA & "'," & _
        "'0','" & in_valor_venta & "','" & in_igv & "','0','0','0','0','0','0','" & in_total & "','0'," & _
        " '" & KEY_USUARIO & "','OBSERVACION','01','" & get_periodo_actual(KEY_FECHA) & "','" & in_cta_compra & "','" & KEY_USUARIO & "','0','0','0','0','" & KEY_RUC & "')"
        Call ConfiguraRstP(strCadena)
        
        id_compra = rstP(0)
                
        strCadena = "UPDATE movimiento_compra SET fecha_kardex='" & Format(KEY_FECHA, "YYYY-mm-dd") & "' WHERE id_compra='" & id_compra & "'"
        Call ConfiguraRstP(strCadena)
        
                in_total = in_costo * Val(Me.txtCantA.Text)
                If KEY_CON_IGV = "si" Then
                    in_valor_venta = in_total
                    in_igv = 0
                Else
                    in_valor_venta = in_total
                    in_igv = 0
                End If
                
                strCadena = "INSERT INTO movimiento_compra_detalle(id_compra,id_producto,cantidad,c_unitario,dsto_soles,dsto_procentaje,total_descuento,valor_neto,isc,igv,retencion,otros,percepcion, " & _
                "valor_venta,exonerado,total,p_venta,p_costo,id_alm,detalle,incremento_fs,ruc) VALUES ('" & id_compra & "','" & Trim(rstA("id_producto")) & "','" & Val(Me.txtCantA.Text) * Val(Me.txtCantidad.Text) & "','" & Val(in_costo) & "'," & _
                "'0','0','0','" & in_valor_venta & "','" & Val(in_igv) & "','0', " & _
                "'0','0','0','" & in_valor_venta & "','0','" & Val(in_costo) * Val(Me.txtCantA.Text) * Val(Me.txtCantidad.Text) & "','" & Val(in_costo) & "','" & Val(in_costo) & "','" & Me.DtcAlmacen.BoundText & "','" & Me.Txtproducto.Text & "','0','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
                
                
               
               strCadena = "call put_kardex_stock_vitekey_v1('02','" & Format(KEY_FECHA, "YYYY-mm-dd") & "','" & Val(id_compra) & "','0090','" & in_serie & "','" & in_numero & "','" & KEY_RUC & "','" & Trim(rstA("id_producto")) & "','" & Val(Me.txtCantA.Text) * Val(Me.txtCantidad.Text) & "','" & Val(in_costo) & "','" & Me.DtcAlmacen.BoundText & "','" & KEY_USUARIO & "','no','" & KEY_RUC & "')"
                CnBd.Execute (strCadena)
                Call put_actualizar_kardex_update(rstA("id_producto"), Me.DtcAlmacen.BoundText)
                rstA.MoveNext
                DoEvents
                
   Next i
   
   
            
                
                in_total = in_costo_acum / (Val(Me.txtCantB.Text) * Val(Me.txtCantidad.Text))
                
                
                
                If KEY_CON_IGV = "si" Then
                    in_valor_venta = in_total / (1 + KEY_IGV)
                    in_igv = in_total - in_valor_venta
                Else
                    in_valor_venta = in_total
                    in_igv = 0
                End If
            
            
            strCadena = "call P_insert_compra_ultimate('0089','" & Me.DtcAlmacen.BoundText & "','" & Format(KEY_FECHA, "YYYY-mm-dd") & "','" & Format(KEY_FECHA, "YYYY-mm-dd") & "','02'," & _
            "'03','--','00001','" & formato_item(Month(KEY_FECHA), 2) & "','" & Year(KEY_FECHA) & "','" & in_serie & "'," & _
            "'" & in_numero & "','6','" & KEY_RUC & "','" & KEY_EMPRESA & "','" & KEY_CAMBIO_VENTA & "'," & _
            "'0','" & Val(in_valor_venta) & "','" & Val(in_igv) & "','0','0','0','0','0','0','" & Val(in_total) & "','" & Val(in_total) & "'," & _
            " '" & KEY_USUARIO & "','OBSERVACION','01','" & get_periodo_actual(KEY_FECHA) & "','" & in_cta_compra & "','" & KEY_USUARIO & "','0','0','0','0','" & KEY_RUC & "')"
            Call ConfiguraRstP(strCadena)
            id_compra = rstP(0)
            
            strCadena = "UPDATE movimiento_compra SET fecha_kardex='" & Format(KEY_FECHA, "YYYY-mm-dd") & "' WHERE id_compra='" & id_compra & "'"
            Call ConfiguraRstP(strCadena)
        
        
           strCadena = "INSERT INTO movimiento_compra_detalle(id_compra,id_producto,cantidad,c_unitario,dsto_soles,dsto_procentaje,total_descuento,valor_neto,isc,igv,retencion,otros,percepcion, " & _
           "valor_venta,exonerado,total,p_venta,p_costo,id_alm,detalle,incremento_fs,ruc) VALUES ('" & id_compra & "','" & Trim(Me.TxtCodigobarraB.Text) & "','" & Val(Me.txtCantidad.Text) * Val(Me.txtCantB.Text) & "','" & Val(in_total) & "'," & _
           "'0','0','0','" & Val(in_total) & "','0','0', " & _
           "'0','0','0','" & Val(in_total) & "','0','" & Val(in_total) & "','" & Val(get_precio_venta_now(Me.TxtCodigobarraB.Text)) & "','" & Val(in_total) & "','" & Me.DtcAlmacen.BoundText & "','" & Trim(Me.txtDescripcionB.Text) & "','0','" & KEY_RUC & "')"
           CnBd.Execute (strCadena)
           
           
           
           
           strCadena = "call put_kardex_stock_vitekey_v1('02','" & Format(KEY_FECHA, "YYYY-mm-dd") & "','" & Val(id_compra) & "','0089','" & in_serie & "','" & in_numero & "','" & KEY_RUC & "','" & Trim(Me.TxtCodigobarraB.Text) & "','" & Val(Me.txtCantidad.Text) * Val(Me.txtCantB.Text) & "','" & Val(in_total) & "','" & Me.DtcAlmacen.BoundText & "','" & KEY_USUARIO & "','no','" & KEY_RUC & "')"
           CnBd.Execute (strCadena)
           
           
           
           
           MsgBox "TRANSFORMACION REALIZADA EXITOSAMENTE", vbInformation
   
   
End If

End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Unload Me
End If
End Sub
Public Sub Resalta(ByVal Texto As TextBox)
Texto.SelStart = 0
Texto.SelLength = Len(Trim(Texto))
Texto.Text = Texto.SelText
Texto.SetFocus
End Sub
Private Sub Form_Load()
CenterForm Me
prodB = False
Me.Top = 500
  Call actualizar
  

  
  
  
  
End Sub

Sub actualizar()
  strCadena = "SELECT id_alm as Codigo, descripcion as Descripcion FROM almacen WHERE ruc='" & KEY_RUC & "'  ORDER BY id_alm ASC"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcAlmacen)
  Me.DtcAlmacen.BoundText = KEY_ALM
  
  strTransformacion = formato_item(ConsultaUltimoRegistro("producto_transformacion", "id_transformacion", "ruc", KEY_RUC), 6)
  Me.TxtTranformacion.Text = strTransformacion
  Me.txtcodigo.Text = ""
  Me.Txtproducto.Text = ""
  Me.TxtCodigobarraB.Text = ""
  Me.txtDescripcionB.Text = ""
  Me.lblStock.Caption = ""
  Me.LblStockB.Caption = ""
  Me.txtcosto.Text = ""
  Me.txtCostoB.Text = ""
  Me.TxtPVenta.Text = ""
  Me.TxtVentaB.Text = ""
  Me.txtCantA.Text = ""
  Me.txtCantB.Text = ""
  Me.txtCantidad.Text = ""
End Sub

Private Sub txtCantB_Change()

If (Val(Me.txtCantA.Text) > 0 And Val(Me.txtCantB.Text) > 0) Then
    
Else
    
End If


End Sub

Private Sub txtcantidad_Change()
If Val(Me.txtCantidad.Text) > 0 Then
    
    Me.cmdProcesar.Enabled = True
End If
End Sub

Private Sub txtcodigo_KeyPress(KeyAscii As Integer)
Dim Criterio As String
If KeyAscii = 13 Then
  If Len(Me.txtcodigo.Text) > 0 Then
  
    Me.txtcodigo.Text = formato_item(Me.txtcodigo.Text, 5)
    Criterio = " A.id_producto LIKE '%" & Trim(Me.txtcodigo.Text) & "%'"
  
   
  
  
  ' If KEY_BARRAS = "si" Then
  '      strCadena = "SELECT A.id_producto,P.nombre_prod,U.abreviatura,A.stock,P.precio_compra,P.precio_venta FROM almacen_producto A,producto P,unidad U,producto_barras B WHERE A.id_producto=P.id_producto AND P.id_unidad=U.id_und AND A.ruc='" & KEY_RUC & "'" & _
   '     " AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND A.id_alm='" & KEY_ALM & "' AND A.id_producto=B.id_producto AND B.ruc='" & KEY_RUC & "' AND P.id_producto=B.id_producto AND  " & Criterio & "ORDER BY nombre_prod LIMIT 0,27"
 
    ' Else
        strCadena = "SELECT A.id_producto,P.nombre_prod,U.abreviatura,A.stock,P.precio_compra,P.precio_venta FROM almacen_producto A,producto P,unidad U WHERE A.id_producto=P.id_producto AND P.id_unidad=U.id_und AND A.ruc='" & KEY_RUC & "'" & _
        " AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND A.id_alm='" & KEY_ALM & "' AND  " & Criterio & ""
    ' End If
    
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        cod_producto = rst("id_producto")
        Me.Txtproducto.Text = rst("nombre_prod")
        Me.lblStock.Caption = rst("stock")
        Me.txtcosto.Text = rst("precio_compra")
        Me.TxtPVenta.Text = rst("precio_venta")
        Set rst = Nothing
        
    Else
         Procedencia = transformaciones
         prodA = True
         FrmProducto.Show
    End If
    Else
        Procedencia = transformaciones
         prodA = True
        FrmProducto.Show
    End If
End If
End Sub





Private Sub TxtCodigobarraB_KeyPress(KeyAscii As Integer)
Dim Criterio As String
If KeyAscii = 13 Then
  If Len(Me.TxtCodigobarraB.Text) > 0 Then
  
  
  
  
  'If KEY_BARRAS = "no" Then
    Me.TxtCodigobarraB.Text = formato_item(Me.TxtCodigobarraB.Text, 5)
    Criterio = " A.id_producto LIKE '%" & Trim(Me.TxtCodigobarraB.Text) & "%'"
  'Else
   ' Criterio = "B.cod_barra= '" & Trim(Me.TxtCodigobarraB.Text) & "'"
  'End If
  
  ' If KEY_BARRAS = "si" Then
   '     strCadena = "SELECT A.id_producto,P.nombre_prod,U.abreviatura,A.stock,P.precio_compra,P.precio_venta FROM almacen_producto A,producto P,unidad U,producto_barras B WHERE A.id_producto=P.id_producto AND P.id_unidad=U.id_und AND A.ruc='" & KEY_RUC & "'" & _
   '     " AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND A.id_alm='" & KEY_ALM & "' AND A.id_producto=B.id_producto AND B.ruc='" & KEY_RUC & "' AND P.id_producto=B.id_producto AND  " & Criterio & "ORDER BY nombre_prod LIMIT 0,27"
 
    ' Else
        strCadena = "SELECT A.id_producto,P.nombre_prod,U.abreviatura,A.stock,P.precio_compra,P.precio_venta FROM almacen_producto A,producto P,unidad U WHERE A.id_producto=P.id_producto AND P.id_unidad=U.id_und AND A.ruc='" & KEY_RUC & "'" & _
        " AND P.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND A.id_alm='" & KEY_ALM & "' AND  " & Criterio & ""
    ' End If
    
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        cod_producto = rst("id_producto")
        Me.txtDescripcionB.Text = rst("nombre_prod")
        Me.LblStockB.Caption = rst("stock")
        Me.txtCostoB.Text = rst("precio_compra")
        Me.TxtVentaB.Text = rst("precio_venta")
        Set rst = Nothing
        
    Else
         Procedencia = transformaciones
         prodB = True
         FrmProducto.Show
    End If
    Else
        Procedencia = transformaciones
         prodB = True
        FrmProducto.Show
    End If
End If
End Sub
