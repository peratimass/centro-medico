VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmAlmacenes 
   BorderStyle     =   0  'None
   Caption         =   "Almacenes-Comprobantes"
   ClientHeight    =   8520
   ClientLeft      =   2640
   ClientTop       =   1320
   ClientWidth     =   15390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   15390
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtid_producto 
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
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   13680
      MaxLength       =   80
      TabIndex        =   22
      Top             =   2840
      Width           =   1335
   End
   Begin VitekeySoft.ChameleonBtn cmdnuevo 
      Height          =   800
      Left            =   13680
      TabIndex        =   13
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1402
      BTYPE           =   5
      TX              =   "Nuevo"
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmAlmacen.frx":0000
      PICN            =   "FrmAlmacen.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.TextBox TxtSerie 
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
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   5850
      MaxLength       =   80
      TabIndex        =   3
      Top             =   7800
      Width           =   975
   End
   Begin VB.TextBox TxtNumero 
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
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   6840
      MaxLength       =   80
      TabIndex        =   2
      Top             =   7800
      Width           =   1335
   End
   Begin VB.TextBox txtbuscar 
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
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   12120
      MaxLength       =   80
      TabIndex        =   0
      Top             =   4395
      Width           =   1455
   End
   Begin MSDataListLib.DataCombo DtcTipoEntidad 
      Height          =   330
      Left            =   7920
      TabIndex        =   1
      Top             =   4395
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      ForeColor       =   8388608
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgAlmacen 
      Height          =   3975
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   7011
      _Version        =   393216
      ForeColor       =   8388608
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      ForeColorFixed  =   12582912
      BackColorBkg    =   16777215
      GridColor       =   8388608
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgComprobante 
      Height          =   2535
      Left            =   120
      TabIndex        =   5
      Top             =   4800
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   4471
      _Version        =   393216
      ForeColor       =   8388608
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      ForeColorFixed  =   12582912
      BackColorBkg    =   16777215
      GridColor       =   8388608
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
   Begin MSDataListLib.DataCombo DtcComprobante 
      Height          =   330
      Left            =   330
      TabIndex        =   6
      Top             =   7800
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   8388608
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   7560
      Top             =   2790
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
            Picture         =   "FrmAlmacen.frx":046E
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAlmacen.frx":08C2
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAlmacen.frx":0BE2
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAlmacen.frx":1036
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAlmacen.frx":148A
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAlmacen.frx":17AA
            Key             =   "(Cancelar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAlmacen.frx":1ACA
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAlmacen.frx":1DEA
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAlmacen.frx":210A
            Key             =   "(Buscar)"
         EndProperty
      EndProperty
   End
   Begin VitekeySoft.ChameleonBtn cmdmodificar 
      Height          =   800
      Left            =   13680
      TabIndex        =   14
      Top             =   1100
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1402
      BTYPE           =   5
      TX              =   "Modificar"
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmAlmacen.frx":242A
      PICN            =   "FrmAlmacen.frx":2446
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdeliminar 
      Height          =   795
      Left            =   13680
      TabIndex        =   15
      Top             =   1950
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1402
      BTYPE           =   5
      TX              =   "Eliminar"
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmAlmacen.frx":4A7F
      PICN            =   "FrmAlmacen.frx":4A9B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdsalir 
      Height          =   795
      Left            =   13680
      TabIndex        =   16
      Top             =   4200
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1402
      BTYPE           =   5
      TX              =   "Salir"
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmAlmacen.frx":6EE5
      PICN            =   "FrmAlmacen.frx":6F01
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdupdate 
      Height          =   615
      Left            =   13680
      TabIndex        =   17
      Top             =   6000
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1085
      BTYPE           =   5
      TX              =   "Modificar"
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmAlmacen.frx":72F1
      PICN            =   "FrmAlmacen.frx":730D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmddelete 
      Height          =   615
      Left            =   13680
      TabIndex        =   18
      Top             =   6720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1085
      BTYPE           =   5
      TX              =   "Eliminar"
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmAlmacen.frx":9946
      PICN            =   "FrmAlmacen.frx":9962
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdAgregar 
      Height          =   405
      Left            =   8400
      TabIndex        =   19
      Top             =   7800
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   714
      BTYPE           =   5
      TX              =   "AGREGAR"
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
      COLTYPE         =   2
      FOCUSR          =   -1  'True
      BCOL            =   16777215
      BCOLO           =   16777215
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmAlmacen.frx":BDAC
      PICN            =   "FrmAlmacen.frx":BDC8
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdKardex 
      Height          =   795
      Left            =   13680
      TabIndex        =   20
      Top             =   3405
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1402
      BTYPE           =   5
      TX              =   "ACTUALIZAR STOCK"
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
      FCOL            =   8388608
      FCOLO           =   8388608
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmAlmacen.frx":F09E
      PICN            =   "FrmAlmacen.frx":F0BA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.ProgressBar progress 
      Height          =   195
      Left            =   13680
      TabIndex        =   21
      Top             =   3180
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Label LblRazonSocial 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "COMPROBANTE "
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
      Left            =   330
      TabIndex        =   12
      Top             =   7515
      Width           =   1245
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SERIE"
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
      Left            =   5850
      TabIndex        =   11
      Top             =   7515
      Width           =   465
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NUMERO"
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
      Left            =   6810
      TabIndex        =   10
      Top             =   7515
      Width           =   735
   End
   Begin VB.Label LblAlamacen 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Height          =   315
      Left            =   2595
      TabIndex        =   9
      Top             =   4395
      Width           =   4365
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "COMPROBANTES  X SUCURSAL :"
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
      Left            =   150
      TabIndex        =   8
      Top             =   4440
      Width           =   2385
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AREA :"
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
      Left            =   7200
      TabIndex        =   7
      Top             =   4440
      Width           =   495
   End
   Begin VB.Shape ShpProducto 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Height          =   855
      Left            =   120
      Top             =   7440
      Width           =   13455
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   8520
      Left            =   0
      Top             =   0
      Width           =   15390
   End
End
Attribute VB_Name = "FrmAlmacenes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cod_almacen As String
Dim cod_comprobante As String
Dim strCodDet_alm As String
Public Procedencia As EnumProcede



Private Sub cmdagregar_Click()
 Call Save
End Sub

Private Sub cmddelete_Click()
If MsgBox(MSGELIMINAR, vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
        strCadena = "DELETE FROM almacen_comprobante WHERE id_alm_com='" & Val(Me.HfgComprobante.TextMatrix(Me.HfgComprobante.Row, 0)) & "'"
        CnBd.Execute (strCadena)
        Call actualizacomp(Me.HfgComprobante, Me.HfgAlmacen.TextMatrix(Me.HfgAlmacen.Row, 0))
      End If
End Sub

Private Sub cmdestructura_Click()

End Sub

Private Sub cmdEliminar_Click()
Procedencia = Eliminar
Call disabled_form(Me)
frmsegurity.Show

End Sub

Private Sub cmdkardex_Click()
On Error GoTo salir

If MsgBox("Esta seguro de Actualizar los Stock's" + Chr(13) + "No cierre esta pantalla hasta que Finalice.", vbQuestion + vbYesNo, KEY_VENDEDOR) = vbYes Then
    
  If Trim(Me.txtid_producto.Text) <> "" Then
    strCadena = "SELECT id_producto,stock,id_alm FROM almacen_producto WHERE id_producto='" & Trim(Me.txtid_producto.Text) & "' and  id_alm='" & Me.HfgAlmacen.TextMatrix(Me.HfgAlmacen.Row, 0) & "' and  ruc='" & KEY_RUC & "'"
  Else
    strCadena = "SELECT id_producto,stock,id_alm FROM almacen_producto WHERE  id_alm='" & Me.HfgAlmacen.TextMatrix(Me.HfgAlmacen.Row, 0) & "' and  ruc='" & KEY_RUC & "'"
  End If
   
   Call ConfiguraRst(strCadena)
   If rst.RecordCount > 0 Then
      rst.MoveFirst
      
     
   
      Me.progress.Min = 0
      Me.progress.Max = rst.RecordCount
      For i = 0 To rst.RecordCount - 1
           strCadena = "call ADM_saldo_producto_kardex('" & rst("id_producto") & "','" & rst("id_alm") & "','" & rst("stock") & "','" & KEY_RUC & "')"
           Call ConfiguraRstL(strCadena)
           
           If rstL(0) <> "-" Then
                MsgBox rstL(0)
           End If
           
           rst.MoveNext
           On Error GoTo N
           Me.progress.Value = i
           DoEvents
      Next i
   End If
End If

Exit Sub
salir:
N:

End Sub

Private Sub cmdModificar_Click()
     Procedencia = modificar
      FrmDetalleAlmacen.Show
End Sub

Private Sub cmdNuevo_Click()
 Procedencia = nuevo
      FrmDetalleAlmacen.Show
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdupdate_Click()
      Procedencia = modificar
      frmDetalleComprobanteAlmacen.Show
      
End Sub

Private Sub CoolBar2_HeightChanged(ByVal NewHeight As Single)

End Sub

Private Sub DtcComprobante_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtSerie.SetFocus
End If
End Sub

Private Sub DtcTipoEntidad_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
      strCadena = "SELECT * FROM view_almacen WHERE id_tipoentidad='" & Trim(Me.DtcTipoEntidad.BoundText) & "' and  ruc='" & KEY_RUC & "' ORDER BY tipoentidad ASC"
      Call listar_almacenes(Me.HfgAlmacen)
    Exit Sub
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 50

strCadena = "SELECT id_doc as Codigo  ,CONCAT(id_doc,' - ',doc_des) as Descripcion FROM comprobantes ORDER BY id_doc ASC "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcComprobante)

strCadena = "SELECT id_tipoentidad as Codigo,descripcion as Descripcion FROM  almacen_tipo ORDER BY descripcion  "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcTipoEntidad)
Call Actualizar_Alm
Call desactivar_controles


End Sub
Public Sub activar_controles()
Me.cmdmodificar.Enabled = True
Me.cmdeliminar.Enabled = True
End Sub
Public Sub desactivar_controles()
Me.cmdmodificar.Enabled = False
Me.cmdeliminar.Enabled = False
End Sub
Public Sub Actualizar_Alm()
If KEY_CARGO = "00001" Then
     strCadena = "SELECT * FROM view_almacen WHERE id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "' ORDER BY descripcion ASC"
Else
    strCadena = "SELECT * FROM view_almacen WHERE ruc='" & KEY_RUC & "' ORDER BY descripcion"
End If
Call listar_almacenes(Me.HfgAlmacen)
End Sub



Private Sub HfgAlmacen_SelChange()
If Val(Me.HfgAlmacen.TextMatrix(Me.HfgAlmacen.Row, 0)) > 0 Then
   Me.LblAlamacen.Caption = Me.HfgAlmacen.TextMatrix(Me.HfgAlmacen.Row, 1)
   Call actualizacomp(Me.HfgComprobante, Me.HfgAlmacen.TextMatrix(Me.HfgAlmacen.Row, 0))
   Call activar_controles
   
Else
    Call desactivar_controles
    Me.HfgComprobante.Rows = 0
    
  End If
End Sub





Private Sub HfgComprobante_SelChange()
If Val(Me.HfgComprobante.TextMatrix(Me.HfgComprobante.Row, 0)) > 0 Then
    If KEY_CARGO = KEY_ADMIN Or KEY_CARGO = KEY_SUPER Or KEY_CARGO = "00001" Or KEY_CARGO = "00023" Then
        Me.cmddelete.Enabled = True
        Me.cmdupdate.Enabled = True
        
    Else
        Me.cmddelete.Enabled = False
        Me.cmdupdate.Enabled = False
    End If
Else
    Me.cmddelete.Enabled = False
        Me.cmdupdate.Enabled = False
End If
End Sub

Private Sub medubicacion_Click()
frmalmaceneestructura.Show
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
 Select Case Button.key
    Case KEY_NEW
     
    Case KEY_UPDATE
     
    Case KEY_DELETE
          
            
  Case KEY_EXIT
        Unload Me
  End Select
End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo error
  Select Case Button.key
    Case KEY_SAVE
      Call Save
    Case KEY_DELETE
        If MsgBox(MSGELIMINAR, vbQuestion + vbYesNo, MSGVERIFICACION) = vbYes Then
        Me.HfgComprobante.col = 0
        cod_comprobante = Trim(Me.HfgComprobante.Text)
        strCadena = "DELETE FROM Det_alm_com WHERE cod_det_alm='" & Trim(Me.HfgComprobante.Text) & "'"
        Call EjecutaRST(strCadena)
        Set RstEjecuta = Nothing
        Call actualizacomp(Me.HfgAlmacen, Me.HfgAlmacen.TextMatrix(Me.HfgAlmacen.Row, 0))
      End If
  End Select
  Exit Sub
error:
  MsgBox Err.Number & " " & Err.Description, vbCritical, MSGERROR
  Exit Sub
End Sub

Private Sub TxtBuscar_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT * FROM view_almacen WHERE descripcion LIKE '%" & Trim(Me.txtbuscar.Text) & "%' and  ruc='" & KEY_RUC & "' ORDER BY tipoentidad ASC"
      Call listar_almacenes(Me.HfgAlmacen)
End If
End Sub

Private Sub txtNumero_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtNumero.Text = FormatosCeros(Me.TxtNumero.Text, 6)
    Me.cmdAgregar.SetFocus
End If
End Sub

Private Sub txtSerie_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtSerie.Text = FormatosCeros(Me.TxtSerie.Text, 3)
    Me.TxtNumero.SetFocus
End If
End Sub
Private Sub Save()

If Me.TxtSerie.Text <> "" And Me.TxtNumero.Text <> "" And Me.DtcComprobante.BoundText <> "" Then
    If ValidarSerie = False Then
       strCadena = "INSERT INTO almacen_comprobante(ruc,id_alm,id_doc,serie,numero)VALUES " & _
        "('" & KEY_RUC & "','" & Trim(HfgAlmacen.TextMatrix(Me.HfgAlmacen.Row, 0)) & "','" & DtcComprobante.BoundText & "','" & Trim(TxtSerie.Text) & "','" & FormatosCeros(TxtNumero.Text, 6) & "')"
        CnBd.Execute (strCadena)
        Call actualizacomp(Me.HfgComprobante, Me.HfgAlmacen.TextMatrix(Me.HfgAlmacen.Row, 0))
        Me.TxtSerie.Text = ""
        Me.TxtNumero.Text = ""
        Me.TxtSerie.SetFocus
    Else
        MsgBox "SERIE ASIGNADA A OTRA SUCURSAL, REGISTRE OTRA SERIE  ", vbCritical, KEY_EMPRESA
        Call Resalta(Me.TxtSerie)
        Exit Sub
        Exit Sub
     End If
Else
    MsgBox "Documento Ya registardo para dicho almacen", vbCritical, KEY_EMPRESA
End If
End Sub
Public Sub actualizacomp(ByVal Grilla As MSHFlexGrid, ByVal idalmacen As String)
On Error GoTo salir
strCadena = "SELECT id_alm_com,A.id_doc,C.doc_des,A.serie,A.numero,A.igv,M.descripcion,A.defecto,A.id_alm,A.numero_caracteres,A.electronico,A.online FROM almacen_comprobante A,moneda M,comprobantes C WHERE A.id_doc=C.id_doc AND A.id_moneda=M.id_moneda AND A.ruc='" & KEY_RUC & "' AND A.id_alm='" & idalmacen & "' ORDER BY A.id_doc"
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
      
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1000
           Grilla.ColWidth(2) = 4500
           Grilla.ColWidth(3) = 1000
           Grilla.ColWidth(4) = 1200
           Grilla.ColWidth(5) = 1200
           Grilla.ColWidth(6) = 1500
           Grilla.ColWidth(7) = 1400
           Grilla.ColWidth(8) = 1300
        Next
         cabecera = "CODIGO" & vbTab & "SUNAT" & vbTab & "COMPROBANTE" & vbTab & "SERIE" & vbTab & "NUMERO" & vbTab & "IGV" & vbTab & "MONEDA" & vbTab & "ELECTRONICO" & vbTab & "ON-LINE"
         Grilla.AddItem cabecera
         For k = 1 To 8
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
          Next k
                            
        rst.MoveFirst
        For i = 1 To rst.RecordCount
            If rst("electronico") = "si" Then
                in_electronico = "            X     "
            Else
                in_electronico = "   "
            End If
            
            If rst("online") = "si" Then
                in_online = "            X     "
            Else
                in_online = "   "
            End If
            
             Fila = rst("id_alm_com") & vbTab & rst("id_doc") & vbTab & rst("doc_des") & vbTab & rst("serie") & vbTab & Format(rst("numero"), rst("numero_caracteres")) & vbTab & UCase(rst("igv")) & vbTab & UCase(rst("descripcion")) & vbTab & in_electronico & vbTab & in_online
             Grilla.AddItem Fila
             If rst("defecto") = "si" Then
        
                            For k = 1 To 8
                                Grilla.col = k
                                Grilla.Row = i
                                Grilla.CellBackColor = &H8080FF
                            Next k
        
             End If
            
        Fila = ""
        rst.MoveNext
        Next i
        
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub
Public Sub listar_almacenes(ByVal Grilla As MSHFlexGrid)
'On Error GoTo SALIR
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    
    Exit Sub

End If
 N = 1
   
   Grilla.Rows = 0
      
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 800
           Grilla.ColWidth(1) = 3000
           Grilla.ColWidth(2) = 4000
           Grilla.ColWidth(3) = 1350
           Grilla.ColWidth(4) = 2000
           Grilla.ColWidth(5) = 2000
           
        Next
         cabecera = "CODIGO" & vbTab & "AREA " & vbTab & "DESCRIPCION" & vbTab & "NUM" & vbTab & "DIRECCION" & vbTab & "HORARIO"
         Grilla.AddItem cabecera
         For k = 0 To 5
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
          Next k
                            
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
             Fila = rst("id_alm") & vbTab & rst("sucursal") & vbTab & rst("descripcion") & vbTab & rst("abreviatura") & vbTab & rst("direccion") & vbTab & "[" & Format(rst("hora_inicio"), "HH:mm") & " - " & Format(rst("hora_fin"), "HH:mm") & " ]"
             Grilla.AddItem Fila
             If rst("defecto") = "si" Then
        
                            For k = 0 To 5
                                Grilla.col = k
                                Grilla.Row = i + 1
                                Grilla.CellBackColor = &H8080FF
                            Next k
        
             End If
            
      
        rst.MoveNext
        Next i
        
Exit Sub
'SALIR: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub


Private Function ValidarSerie() As Boolean
    strCadena = "SELECT * FROM almacen_comprobante WHERE ( id_doc='" & Me.DtcComprobante.BoundText & "'AND serie='" & Me.TxtSerie.Text & "' AND ruc='" & Trim(KEY_RUC) & "') "
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        ValidarSerie = True
    Else
        ValidarSerie = False
    End If
    Set rst = Nothing
End Function




