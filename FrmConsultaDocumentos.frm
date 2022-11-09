VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmBusquedaDocumentos 
   BorderStyle     =   0  'None
   Caption         =   "Busqueda de Documentos"
   ClientHeight    =   8865
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14865
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   14865
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "CERRAR"
      Height          =   375
      Left            =   12840
      TabIndex        =   26
      Top             =   240
      Width           =   1815
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   2355
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      ForeColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "PRODUCTO"
      TabPicture(0)   =   "FrmConsultaDocumentos.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Shape1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "DtpPfin"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "TxtCodigoProducto"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "TxtDescripcion"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "DtpPinicio"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdBuscar_producto"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtUnidad"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "TxtCodigoInterno"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "CLIENTE/PROVEEDOR"
      TabPicture(1)   =   "FrmConsultaDocumentos.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Txtdni"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "txtCliente"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cmdBuscarCliente"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "DtpInicioCliente"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "DtpFinCliente"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label6"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label2"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Shape2"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "FECHA EMISION"
      TabPicture(2)   =   "FrmConsultaDocumentos.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "CmdBuscarFecha"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "DtpIniFecha"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "DtpFinFecha"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label7"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label3"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Shape3"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "TIPO COMPROBANTE"
      TabPicture(3)   =   "FrmConsultaDocumentos.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "CmdBuscarComprobante"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "DtcComprobante"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "DtpInicioComprobante"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "DtpFinComprobante"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Label8"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Label4"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "Shape4"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).ControlCount=   7
      Begin VB.TextBox TxtCodigoInterno 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   2160
         TabIndex        =   32
         Top             =   960
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CommandButton CmdBuscarComprobante 
         Caption         =   "BUSCAR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   -64920
         TabIndex        =   27
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton CmdBuscarFecha 
         Caption         =   "BUSCAR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   -69240
         TabIndex        =   22
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox Txtdni 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   -72840
         TabIndex        =   17
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox txtCliente 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   -70800
         TabIndex        =   16
         Top             =   720
         Width           =   5415
      End
      Begin VB.CommandButton cmdBuscarCliente 
         Caption         =   "BUSCAR"
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
         Height          =   280
         Left            =   -61680
         TabIndex        =   14
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtUnidad 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   4200
         TabIndex        =   12
         Top             =   720
         Width           =   735
      End
      Begin VB.CommandButton cmdBuscar_producto 
         Caption         =   "BUSCAR"
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
         Height          =   280
         Left            =   13320
         TabIndex        =   11
         Top             =   720
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker DtpPinicio 
         Height          =   300
         Left            =   9720
         TabIndex        =   8
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   393216
         Format          =   62717953
         CurrentDate     =   41118
      End
      Begin VB.TextBox TxtDescripcion 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   5040
         TabIndex        =   4
         Top             =   720
         Width           =   4575
      End
      Begin VB.TextBox TxtCodigoProducto 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   2160
         TabIndex        =   3
         Top             =   720
         Width           =   1935
      End
      Begin MSComCtl2.DTPicker DtpPfin 
         Height          =   300
         Left            =   11640
         TabIndex        =   10
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   393216
         Format          =   62717953
         CurrentDate     =   41118
      End
      Begin MSComCtl2.DTPicker DtpInicioCliente 
         Height          =   300
         Left            =   -65280
         TabIndex        =   15
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   393216
         Format          =   62717953
         CurrentDate     =   41118
      End
      Begin MSComCtl2.DTPicker DtpFinCliente 
         Height          =   300
         Left            =   -63360
         TabIndex        =   18
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   393216
         Format          =   62717953
         CurrentDate     =   41118
      End
      Begin MSComCtl2.DTPicker DtpIniFecha 
         Height          =   300
         Left            =   -72840
         TabIndex        =   23
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   393216
         Format          =   62717953
         CurrentDate     =   41118
      End
      Begin MSComCtl2.DTPicker DtpFinFecha 
         Height          =   300
         Left            =   -70920
         TabIndex        =   24
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   393216
         Format          =   62717953
         CurrentDate     =   41118
      End
      Begin MSDataListLib.DataCombo DtcComprobante 
         Height          =   315
         Left            =   -72840
         TabIndex        =   28
         Top             =   720
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ForeColor       =   8388608
         Text            =   "DataCombo1"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker DtpInicioComprobante 
         Height          =   300
         Left            =   -68520
         TabIndex        =   29
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   393216
         Format          =   62717953
         CurrentDate     =   41118
      End
      Begin MSComCtl2.DTPicker DtpFinComprobante 
         Height          =   300
         Left            =   -66600
         TabIndex        =   30
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         _Version        =   393216
         Format          =   62717953
         CurrentDate     =   41118
      End
      Begin VB.Label Label8 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "AL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -66885
         TabIndex        =   31
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label7 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "AL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -71205
         TabIndex        =   25
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "AL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -63645
         TabIndex        =   19
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "AL"
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
         Height          =   300
         Left            =   11360
         TabIndex        =   9
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "COMPROBANTE :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -74760
         TabIndex        =   7
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "INTERVALO DE FECHAS"
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
         Height          =   300
         Left            =   -74760
         TabIndex        =   6
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "RUC /  DNI :"
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
         Height          =   300
         Left            =   -74760
         TabIndex        =   5
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "CODIGO PRODUCTO :"
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
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   1635
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   6  'Inside Solid
         Height          =   615
         Left            =   120
         Top             =   600
         Width           =   14415
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         BorderStyle     =   6  'Inside Solid
         Height          =   615
         Left            =   -74880
         Top             =   600
         Width           =   14415
      End
      Begin VB.Shape Shape4 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         Height          =   615
         Left            =   -74880
         Top             =   600
         Width           =   14295
      End
      Begin VB.Shape Shape3 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00808080&
         BorderStyle     =   6  'Inside Solid
         Height          =   615
         Left            =   -74880
         Top             =   600
         Width           =   14295
      End
   End
   Begin MSDataListLib.DataCombo DtcAlmacen 
      Height          =   360
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   8388608
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo DtcMivimientos 
      Height          =   360
      Left            =   5760
      TabIndex        =   13
      Top             =   240
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   8388608
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdGrilla 
      Height          =   3375
      Left            =   120
      TabIndex        =   20
      Top             =   2400
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   5953
      _Version        =   393216
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      BackColorBkg    =   16777215
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdGrillaDetalle 
      Height          =   2895
      Left            =   120
      TabIndex        =   21
      Top             =   5880
      Width           =   14655
      _ExtentX        =   25850
      _ExtentY        =   5106
      _Version        =   393216
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      BackColorBkg    =   16777215
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
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00DFDFE0&
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   660
      Left            =   120
      Top             =   120
      Width           =   14655
   End
End
Attribute VB_Name = "FrmBusquedaDocumentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede

Private Sub cmdBuscar_producto_Click()
Dim stcodigop As String

    
    strCadena = "SELECT id_producto FROM producto WHERE id_producto='" & Trim(Me.TxtCodigoInterno.Text) & "' AND ruc='" & KEY_RUC & "' "

Call ConfiguraRst(strCadena)

If rst.RecordCount > 0 And Me.txtCodigoProducto.Text <> "" Then
    stcodigop = rst(0)
Else
   Set rst = Nothing
   Procedencia = buscar
   FrmProducto.Show
   Exit Sub
End If
Call llenarGrid(Me.HfdGrilla, stcodigop)
End Sub
Public Sub llenarGrid(ByVal Grilla As MSHFlexGrid, Optional id_producto As String, Optional id_persona As String)
'On Error GoTo SALIR
If Me.DtcMivimientos.BoundText = "00001" Then
    If Me.SSTab1.Tab = 0 Then
        strCadena = "SELECT DISTINCT D.id_producto,M.id_venta,M.fecha_emision,CONCAT(C.doc_abrev,':',M.serie,'-',M.numero) as comprobante,M.ncliente,M.total,P.nombre_completo FROM movimiento_venta M,movimiento_venta_detalle D,comprobantes C,persona P  WHERE M.id_doc=C.id_doc AND M.ruc='" & KEY_RUC & "' AND P.dni=M.id_vendedor AND M.fecha_emision>='" & Format(CVDate(Me.DtpPinicio.Value), "YYYY-mm-dd") & "' AND M.fecha_emision<='" & Format(CVDate(Me.DtpPfin), "YYYY-mm-dd") & "' AND M.id_venta=D.id_venta AND D.ruc='" & KEY_RUC & "' AND D.id_producto ='" & id_producto & "'  ORDER BY M.id_venta ASC  "
    End If
    If Me.SSTab1.Tab = 1 Then
        strCadena = "SELECT M.id_venta,M.fecha_emision,CONCAT(C.doc_abrev,':',M.serie,'-',M.numero) as comprobante,M.ncliente,M.total,P.nombre_completo FROM movimiento_venta M,comprobantes C,persona P  WHERE M.id_doc=C.id_doc AND M.ruc='" & KEY_RUC & "' AND P.dni=M.id_vendedor AND M.fecha_emision>='" & Format(CVDate(Me.DtpInicioCliente.Value), "YYYY-mm-dd") & "' AND M.fecha_emision<='" & Format(CVDate(Me.DtpFinCliente.Value), "YYYY-mm-dd") & "' AND M.id_cliente='" & id_persona & "'  ORDER BY M.id_venta ASC  "
    End If
    If Me.SSTab1.Tab = 2 Then
        strCadena = "SELECT M.id_venta,M.fecha_emision,CONCAT(C.doc_abrev,':',M.serie,'-',M.numero) as comprobante,M.ncliente,M.total,P.nombre_completo FROM movimiento_venta M,comprobantes C,persona P  WHERE M.id_doc=C.id_doc AND M.ruc='" & KEY_RUC & "' AND P.dni=M.id_vendedor AND M.fecha_emision>='" & Format(Me.DtpIniFecha.Value, "YYYY-mm-dd") & "' AND M.fecha_emision<='" & Format(Me.DtpFinFecha.Value, "YYYY-mm-dd") & "'   ORDER BY M.id_venta ASC  "
    End If
    If Me.SSTab1.Tab = 3 Then
        strCadena = "SELECT M.id_venta,M.fecha_emision,CONCAT(C.doc_abrev,':',M.serie,'-',M.numero) as comprobante,M.ncliente,M.total,P.nombre_completo FROM movimiento_venta M,comprobantes C,persona P  WHERE M.id_doc=C.id_doc AND M.ruc='" & KEY_RUC & "' AND P.dni=M.id_vendedor AND M.fecha_emision>='" & Format(CVDate(Me.DtpInicioComprobante.Value), "YYYY-mm-dd") & "' AND M.fecha_emision<='" & Format(CVDate(Me.DtpFinComprobante.Value), "YYYY-mm-dd") & "' AND M.id_doc='" & Trim(Me.dtcComprobante.BoundText) & "' ORDER BY M.id_venta ASC  "
    End If
Else
    If Me.SSTab1.Tab = 0 Then
        strCadena = "SELECT DISTINCT D.id_producto,M.id_compra as id_venta,M.fecha_emision,CONCAT(C.doc_abrev,':',M.serie,'-',M.numero) as comprobante,M.nproveedor as ncliente,M.total,P.nombre_completo FROM movimiento_compra M,movimiento_compra_detalle D,comprobantes C,persona P  WHERE M.id_doc=C.id_doc AND M.ruc='" & KEY_RUC & "' AND P.dni=M.dni_save AND M.fecha_emision>='" & Format(CVDate(Me.DtpPinicio.Value), "YYYY-mm-dd") & "' AND M.fecha_emision<='" & Format(CVDate(Me.DtpPfin), "YYYY-mm-dd") & "' AND M.id_compra=D.id_compra AND D.ruc='" & KEY_RUC & "' AND D.id_producto ='" & id_producto & "'  ORDER BY M.id_compra ASC  "
    End If
    If Me.SSTab1.Tab = 1 Then
        strCadena = "SELECT M.id_compra as id_venta,M.fecha_emision,CONCAT(C.doc_abrev,':',M.serie,'-',M.numero) as comprobante,M.nproveedor as ncliente,M.total,P.nombre_completo FROM movimiento_compra M,comprobantes C,persona P  WHERE M.id_doc=C.id_doc AND M.ruc='" & KEY_RUC & "' AND P.dni=M.dni_save AND M.fecha_emision>='" & Format(CVDate(Me.DtpInicioCliente.Value), "YYYY-mm-dd") & "' AND M.fecha_emision<='" & Format(CVDate(Me.DtpFinCliente.Value), "YYYY-mm-dd") & "' AND M.id_proveedor='" & id_persona & "'  ORDER BY M.id_compra ASC  "
    End If
    If Me.SSTab1.Tab = 2 Then
        strCadena = "SELECT M.id_compra as id_venta,M.fecha_emision,CONCAT(C.doc_abrev,':',M.serie,'-',M.numero) as comprobante,M.nproveedor as ncliente,M.total,P.nombre_completo FROM movimiento_compra M,comprobantes C,persona P  WHERE M.id_doc=C.id_doc AND M.ruc='" & KEY_RUC & "' AND P.dni=M.dni_save AND M.fecha_emision>='" & Format(CVDate(Me.DtpIniFecha.Value), "YYYY-mm-dd") & "' AND M.fecha_emision<='" & Format(CVDate(Me.DtpFinFecha.Value), "YYYY-mm-dd") & "'   ORDER BY M.id_compra ASC  "
    End If
    If Me.SSTab1.Tab = 3 Then
        strCadena = "SELECT M.id_compra as id_venta,M.fecha_emision,CONCAT(C.doc_abrev,':',M.serie,'-',M.numero) as comprobante,M.nproveedor as ncliente,M.total,P.nombre_completo FROM movimiento_venta M,comprobantes C,persona P  WHERE M.id_doc=C.id_doc AND M.ruc='" & KEY_RUC & "' AND P.dni=M.dni_save AND M.fecha_emision>='" & Format(CVDate(Me.DtpInicioComprobante.Value), "YYYY-mm-dd") & "' AND M.fecha_emision<='" & Format(CVDate(Me.DtpFinComprobante.Value), "YYYY-mm-dd") & "' AND M.id_doc='" & Trim(Me.dtcComprobante.BoundText) & "' ORDER BY M.id_compra ASC  "
    End If

End If

Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub
End If

 
'   Grilla.Clear
   Grilla.Rows = rst.RecordCount - 1
   Grilla.Rows = 0
     
    
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1200
           Grilla.ColWidth(2) = 2500
           Grilla.ColWidth(3) = 4500
           Grilla.ColWidth(4) = 2000
           Grilla.ColWidth(5) = 3500
       Next
        cabecera = "IDMOVIMIENTO" & vbTab & "FECHA" & vbTab & "COMPROBANTE " & vbTab & "CLIENTE/PROVEEDOR " & vbTab & "MONTO" & vbTab & "ATENDIDO POR"
        Grilla.AddItem cabecera
         For k = 0 To 5
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        
        For i = 0 To rst.RecordCount - 1
            Fila = rst("id_venta") & vbTab & rst("fecha_emision") & vbTab & rst("comprobante") & vbTab & rst("ncliente") & vbTab & Format(rst("total"), "###0.00") & vbTab & rst("nombre_completo")
            Grilla.AddItem Fila
            rst.MoveNext
        Next i
    
     
  Grilla.Row = 1
  Grilla.col = 0
  Grilla.ColSel = 1
  Grilla.RowSel = 1

Exit Sub
'SALIR: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
'Set rst = Nothing
End Sub
Sub llenarGrid_Comprobante(ByVal Grilla As MSHFlexGrid, ByVal idVenta As Double)
On Error GoTo SALIR
Dim tTotal As Double
Dim texonerado As Double
Dim tafecto As Double
If Me.DtcMivimientos.BoundText = "00001" Then
    strCadena = "SELECT D.id_detalle_venta,D.id_producto,P.nombre_prod,U.abreviatura,D.cantidad,D.precio,D.total,P.id_igv FROM movimiento_venta_detalle D,producto P,unidad U WHERE D.id_producto=P.id_producto AND D.id_venta='" & idVenta & "' AND P.id_unidad=U.id_und AND U.id_usu='" & KEY_RUC & "' AND P.ruc='" & KEY_RUC & "' AND D.ruc='" & KEY_RUC & "'"
Else
    strCadena = "SELECT D.id_detalle_compra as id_detalle_venta,D.id_producto,P.nombre_prod,U.abreviatura,D.cantidad,D.c_unitario as precio,D.total,P.id_igv FROM movimiento_compra_detalle D,producto P,unidad U WHERE D.id_producto=P.id_producto AND D.id_compra='" & idVenta & "' AND P.id_unidad=U.id_und AND U.id_usu='" & KEY_RUC & "' AND P.ruc='" & KEY_RUC & "' AND D.ruc='" & KEY_RUC & "'"
End If
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
           Grilla.ColWidth(1) = 900
           Grilla.ColWidth(2) = 6600
           Grilla.ColWidth(3) = 800
           Grilla.ColWidth(4) = 1200
           Grilla.ColWidth(5) = 1200
           Grilla.ColWidth(6) = 1300
        Next
        cabecera = "IDDETALLE" & vbTab & "CODIGO" & vbTab & "DESCRIPCION " & vbTab & "UND " & vbTab & "CANTIDAD" & vbTab & "PRECIO" & vbTab & "TOTAL"
        Grilla.AddItem cabecera
         For k = 0 To 6
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        tTotal = 0
        For i = 0 To rst.RecordCount - 1
            Fila = rst("id_detalle_venta") & vbTab & rst("id_producto") & vbTab & rst("nombre_prod") & vbTab & rst("abreviatura") & vbTab & Format(rst("cantidad"), "#,##0.00") & vbTab & Format(rst("precio"), "#,##0.00") & vbTab & Format(rst("total"), "#,##0.00")
            Grilla.AddItem Fila
            If (Trim(rst("id_igv")) = "no") Then
                            texonerado = texonerado + rst("total")
                            For k = 0 To 6
                                Grilla.col = k
                                Grilla.Row = i + 1
                                Grilla.CellBackColor = &HC0FFFF
                            Next k
             Else
                            tafecto = tafecto + rst("total")
             End If
            
            rst.MoveNext
    Next i
  
Exit Sub
SALIR: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing
End Sub

Private Sub cmdBuscarCliente_Click()
Call llenarGrid(Me.HfdGrilla, , Trim(Me.txtDni.Text))
End Sub

Private Sub CmdBuscarComprobante_Click()
Call llenarGrid(Me.HfdGrilla)
End Sub

Private Sub CmdBuscarFecha_Click()
Call llenarGrid(Me.HfdGrilla)
End Sub

Private Sub cmdcerrar_Click()
Unload Me
End Sub

Private Sub Form_Load()
    CenterForm Me
    Me.Top = 100
    Me.DtpPinicio.Value = KEY_FECHA
    Me.DtpPfin.Value = KEY_FECHA
    Me.DtpInicioCliente.Value = KEY_FECHA
    Me.DtpFinCliente.Value = KEY_FECHA
    Me.DtpIniFecha.Value = KEY_FECHA
    Me.DtpIniFecha.Value = KEY_FECHA
    Me.DtpFinFecha.Value = KEY_FECHA
    Me.DtpInicioComprobante.Value = KEY_FECHA
    Me.DtpFinComprobante.Value = KEY_FECHA
  strCadena = "SELECT id_alm as Codigo, descripcion as Descripcion FROM almacen WHERE ruc='" & KEY_RUC & "' ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.dtcalmacen)
  Me.dtcalmacen.BoundText = KEY_ALM
  
  strCadena = "SELECT id_tipomov as Codigo, descripcion as Descripcion FROM tipo_movimiento WHERE id_tipomov='00001' OR id_tipomov='00002' ORDER BY descripcion"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcMivimientos)
  Me.DtcMivimientos.BoundText = "00001"
  
  strCadena = "SELECT DISTINCT A.id_doc as Codigo,C.doc_des as Descripcion FROM almacen_comprobante A,comprobantes C WHERE A.id_doc=C.id_doc AND ruc='" & KEY_RUC & "' AND id_alm='" & Me.dtcalmacen.BoundText & "' AND A.venta='si'"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.dtcComprobante)
End Sub

Private Sub HfdGrilla_SelChange()
If Val(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)) > 0 Then
    Call llenarGrid_Comprobante(Me.HfdGrillaDetalle, Val(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)))
End If
End Sub

Private Sub TxtCodigoProducto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
 If (Len(Me.txtCodigoProducto.Text) = 0) Or Val(Me.txtCodigoProducto.Text) = 0 Then
        Procedencia = Selecionar
        FrmProducto.Show
        Exit Sub
    End If
    
    If KEY_BARRAS = "si" Then
        strCadena = "SELECT B.id_producto,P.nombre_prod,U.abreviatura FROM producto_barras B,producto P,unidad U WHERE B.id_producto=P.id_producto AND B.ruc='" & KEY_RUC & "' " & _
        "AND P.ruc='" & KEY_RUC & "' AND B.cod_barra='" & Trim(Me.txtCodigoProducto.Text) & "' AND P.id_unidad=U.id_und AND U.id_usu='" & KEY_RUC & "'"
    Else
        Me.txtCodigoProducto.Text = FormatosCeros(Me.txtCodigoProducto.Text, 5)
        strCadena = "SELECT A.id_producto, P.nombre_prod,U.abreviatura FROM almacen_producto A,producto P,unidad U WHERE P.id_unidad=U.id_und AND U.id_usu='" & KEY_RUC & "' AND A.id_producto=P.id_producto AND A.id_alm='" & KEY_ALM & "' AND A.ruc='" & KEY_RUC & "' AND P.ruc='" & KEY_RUC & "' AND A.id_producto='" & Trim(Me.txtCodigoProducto.Text) & "'"
    End If
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        codigoP = rst("id_producto")
        Me.TxtCodigoInterno.Text = rst("id_producto")
        Me.txtDescripcion.Text = UCase(rst("nombre_prod"))
        Me.TxtUnidad.Text = UCase(rst("abreviatura"))
        Me.cmdBuscar_producto.Enabled = True
        Me.cmdBuscar_producto.SetFocus
    Else
        Call Resalta(Me.txtCodigoProducto)
        Procedencia = Selecionar
        FrmProducto.Show
    End If
End If

End Sub

Private Sub txtDNI_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
strCadena = "SELECT * FROM persona  WHERE dni='" & Trim(Me.txtDni.Text) & "' "
Call ConfiguraRst(strCadena)

If rst.RecordCount > 0 And Me.txtDni.Text <> "" Then
    Me.txtDni.Text = rst("dni")
    Me.txtcliente.Text = UCase(rst("nombre_completo"))
    Me.cmdBuscarCliente.Enabled = True
    Me.cmdBuscarCliente.SetFocus
Else
   Set rst = Nothing
   Procedencia = buscar
   FrmPersona.Show
   Exit Sub
End If
End If
End Sub
