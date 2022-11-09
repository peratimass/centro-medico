VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmParametrosEmpresa 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   12360
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   6375
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Visible         =   0   'False
      Width           =   10815
      Begin VB.CheckBox chk_pago_yape 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         Caption         =   "YAPE (BCP)"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   250
         Left            =   4560
         TabIndex        =   26
         Top             =   4380
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CheckBox chk_pago_mastercard 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         Caption         =   "MASTERCARD"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   250
         Left            =   4560
         TabIndex        =   25
         Top             =   4080
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CheckBox chk_pago_visa 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         Caption         =   "VISA"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   250
         Left            =   4560
         TabIndex        =   24
         Top             =   3780
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.CheckBox chk_pago_efectivo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H0000C000&
         Caption         =   "CONTADO"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   250
         Left            =   4560
         TabIndex        =   23
         Top             =   3480
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox txtNombreComercial 
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
         Left            =   2640
         TabIndex        =   22
         Top             =   1200
         Width           =   5775
      End
      Begin VB.CheckBox chk_tienda_online 
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "TIENDA ONLINE"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   280
         Left            =   2640
         TabIndex        =   20
         Top             =   3120
         Width           =   4095
      End
      Begin VB.TextBox txtObservacion 
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
         Left            =   2640
         TabIndex        =   19
         Top             =   2640
         Width           =   5775
      End
      Begin VB.TextBox txtTiempoEntrega 
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
         Left            =   4560
         TabIndex        =   17
         Top             =   5040
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VB.TextBox txtMontoMinimo 
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
         Left            =   4560
         TabIndex        =   16
         Top             =   4680
         Visible         =   0   'False
         Width           =   2175
      End
      Begin VitekeySoft.ChameleonBtn cmdcrearpantalla 
         Height          =   495
         Left            =   2640
         TabIndex        =   13
         Top             =   5760
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   873
         BTYPE           =   5
         TX              =   "CREAR EMPRESA"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
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
         MICON           =   "FrmParametrosEmpresa.frx":0000
         PICN            =   "FrmParametrosEmpresa.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSDataListLib.DataCombo DtaModoAcceso 
         Height          =   330
         Left            =   2640
         TabIndex        =   12
         Top             =   2160
         Width           =   5775
         _ExtentX        =   10186
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
      Begin VB.TextBox txtdireccionfiscal 
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
         Left            =   2640
         TabIndex        =   11
         Top             =   1680
         Width           =   5775
      End
      Begin VB.TextBox txtrazonsocial 
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
         Left            =   2640
         TabIndex        =   10
         Top             =   720
         Width           =   5775
      End
      Begin VB.TextBox txtruc 
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
         Left            =   2640
         TabIndex        =   9
         Top             =   240
         Width           =   2775
      End
      Begin VitekeySoft.ChameleonBtn cmdcerrarpantalla 
         Height          =   495
         Left            =   4800
         TabIndex        =   14
         Top             =   5760
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   873
         BTYPE           =   5
         TX              =   "CERRAR PANTALLA"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
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
         MICON           =   "FrmParametrosEmpresa.frx":3664
         PICN            =   "FrmParametrosEmpresa.frx":3680
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lbltiempoentrega 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TIEMPO DE ENTREGA (MIN):"
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
         Left            =   2640
         TabIndex        =   27
         Top             =   5160
         Visible         =   0   'False
         Width           =   1830
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NOMBRE COMERCIAL :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   600
         TabIndex        =   21
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OBSERVACION :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   1020
         TabIndex        =   18
         Top             =   2640
         Width           =   1185
      End
      Begin VB.Label lblMinimo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "MONTO MINIMO (S/.) :"
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
         Left            =   2640
         TabIndex        =   15
         Top             =   4680
         Visible         =   0   'False
         Width           =   1500
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RUBRO :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   1560
         TabIndex        =   8
         Top             =   2160
         Width           =   645
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DIRECCION FISCAL :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   750
         TabIndex        =   7
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RAZON SOCIAL :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   1005
         TabIndex        =   6
         Top             =   840
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RUC :"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   1800
         TabIndex        =   5
         Top             =   360
         Width           =   405
      End
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
            Picture         =   "FrmParametrosEmpresa.frx":67B7
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmParametrosEmpresa.frx":6C0B
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmParametrosEmpresa.frx":6F2B
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmParametrosEmpresa.frx":737F
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmParametrosEmpresa.frx":77D3
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmParametrosEmpresa.frx":7AF3
            Key             =   "(Cancelar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmParametrosEmpresa.frx":7E13
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmParametrosEmpresa.frx":8133
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmParametrosEmpresa.frx":8453
            Key             =   "(Buscar)"
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgMarcas 
      Height          =   4695
      Left            =   240
      TabIndex        =   0
      Top             =   390
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   8281
      _Version        =   393216
      ForeColor       =   8388608
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      ForeColorFixed  =   8388608
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
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   4065
      Left            =   11205
      TabIndex        =   1
      Top             =   360
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   7170
      BandCount       =   1
      ForeColor       =   8388608
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      Orientation     =   1
      EmbossShadow    =   -2147483628
      EmbossPicture   =   -1  'True
      _CBWidth        =   900
      _CBHeight       =   4065
      _Version        =   "6.7.9782"
      Caption1        =   "Acciones"
      Child1          =   "TlbAcciones"
      MinHeight1      =   840
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   4050
         Left            =   30
         TabIndex        =   2
         Top             =   375
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   7144
         ButtonWidth     =   1799
         ButtonHeight    =   1429
         Style           =   1
         ImageList       =   "ImgIconos"
         DisabledImageList=   "ImgIconos"
         HotImageList    =   "ImgIconos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Nuevo"
               Key             =   "(Nuevo)"
               Object.ToolTipText     =   "Modificar"
               ImageKey        =   "(Nuevo)"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Modificar"
               Key             =   "(Modificar)"
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
   Begin VB.Label LblEmpresa 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PARAMETROS EMPRESA"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   285
      TabIndex        =   3
      Top             =   0
      Width           =   2925
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00808080&
      BorderWidth     =   2
      Height          =   6855
      Left            =   0
      Top             =   0
      Width           =   12360
   End
End
Attribute VB_Name = "FrmParametrosEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede
Public Sub llenarGrid()
strCadena = "SELECT * FROM entidad_empresa E,entidad_parametros P,persona U WHERE E.cod_unico=P.cod_unico AND P.cod_unico='" & KEY_RUC & "' AND E.cod_unico=U.dni AND P.cod_unico=U.dni AND E.id_empresa='0'"
Call llenarGridP(Me.HfgMarcas, Me)
End Sub
Private Sub llenarGridP(ByVal Grilla As MSHFlexGrid, ByVal Formulario As Form)
On Error GoTo salir
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
           Grilla.ColWidth(0) = 1200
           Grilla.ColWidth(1) = 3000
           Grilla.ColWidth(2) = 3500
           Grilla.ColWidth(3) = 850
           Grilla.ColWidth(4) = 1000
           Grilla.ColWidth(5) = 1000
           
        Next
         cabecera = "RUC/DNI" & vbTab & "RAZON SOCIAL" & vbTab & "DIRECCION FISCAL" & vbTab & "IGV" & vbTab & "PERCEP" & vbTab & "RET"
         Grilla.AddItem cabecera
         For k = 0 To 5
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
          Next k
                            
        rst.MoveFirst
        For i = 1 To rst.RecordCount
             Fila = Fila & rst("cod_unico") & vbTab & UCase(rst("nombre_completo")) & vbTab & UCase(rst("direccion")) & vbTab & UCase(rst("igv")) & vbTab & UCase(rst("id_percepcion")) & vbTab & UCase(rst("id_retencion"))
             Grilla.AddItem Fila
            
        
        rst.MoveNext
        Next i
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub


Private Sub ChameleonBtn2_Click()

End Sub

Private Sub chk_tienda_online_Click()
If Me.chk_tienda_online.Value = 1 Then
   Me.lblMinimo.Visible = True
   Me.lbltiempoentrega.Visible = True
   Me.txtTiempoEntrega.Visible = True
   Me.txtMontoMinimo.Visible = True
   Me.txtTiempoEntrega.Visible = True
   Me.chk_pago_efectivo.Visible = True
   Me.chk_pago_mastercard.Visible = True
   Me.chk_pago_visa.Visible = True
   Me.chk_pago_yape.Visible = True
   
Else
   Me.lblMinimo.Visible = False
   Me.lbltiempoentrega.Visible = False
   Me.txtTiempoEntrega.Visible = False
   Me.txtMontoMinimo.Visible = False
   Me.txtTiempoEntrega.Visible = False
   Me.chk_pago_efectivo.Visible = False
   Me.chk_pago_mastercard.Visible = False
   Me.chk_pago_visa.Visible = False
   Me.chk_pago_yape.Visible = False
End If


End Sub

Private Sub cmdCerrarpantalla_Click()
Frame1.Visible = False
End Sub

Private Sub cmdcrearpantalla_Click()

If Me.chk_tienda_online.Value = 1 Then
    in_tienda_online = "si"
Else
    in_tienda_online = "no"
End If


'EFECTIVO
If Me.chk_pago_efectivo.Value = 1 Then
    in_pago_efectivo = "si"
Else
    in_pago_efectivo = "no"
End If

'visa
If Me.chk_pago_visa.Value = 1 Then
    in_pago_visa = "si"
Else
    in_pago_visa = "no"
End If

'mastercard
If Me.chk_pago_mastercard.Value = 1 Then
    in_pago_mastercard = "si"
Else
    in_pago_mastercard = "no"
End If

If Me.chk_pago_yape.Value = 1 Then
    in_pago_yape = "si"
Else
    in_pago_yape = "no"
End If







strCadena = "SELECT * FROM persona WHERE dni='" & Trim(Me.txtruc.Text) & "' "
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
   strCadena = "call P_insert_persona('" & Trim(Me.txtruc.Text) & "' " & _
                ",'-', " & _
                "'-' " & _
                ",'-' " & _
                ",'" & UCase(Trim(Me.txtrazonsocial.Text)) & "' " & _
                ",'" & Trim(Me.txtdireccionfiscal.Text) & "' " & _
                ",'-' " & _
                ",'-'" & _
                ",'si' " & _
                ",'no'" & _
                ",'si' " & _
                ",'no' " & _
                ",'no' " & _
                ",'no' " & _
                ",'" & Trim(Me.txtruc.Text) & "')"
                CnBd.Execute (strCadena)
         
         
         strCadena = "INSERT INTO entidad_parametros(`cod_unico`,id_tipo_per,`igv`,`factura`,`barras`,`doc_cod`,`serie`,`automatico`,`cerveceria`,`contabilidad`,`update_precios`,`foto_producto`," & _
         "`fingerprint`,`id_alm`,`id_tipo_letra`,`activacion_permanente`,`instalacion`,`caducidad`,`tramite_documentario`,`caja_independiente`,id_proveedor_servicio,tienda_online,monto_minimo,observacion,tiempo_entrega,nombre_comercial,pago_efectivo,pago_visa,pago_mastercard,pago_yape)VALUES(" & _
         "'" & Trim(Me.txtruc.Text) & "','" & Trim(Me.DtaModoAcceso.BoundText) & "','si','no','no','0003','001','si','no','no','no','no','no','00001','01','si','" & KEY_FECHA & "','" & KEY_FECHA & "','no','si','" & KEY_RUC & "'," & _
         " '" & in_tienda_online & "','" & Val(Me.txtMontoMinimo.Text) & "','" & Trim(Me.txtObservacion.Text) & "','" & Val(Me.txtTiempoEntrega.Text) & "','" & Trim(Me.txtNombreComercial.Text) & "','" & in_pago_efectivo & "','" & in_pago_visa & "','" & in_pago_mastercard & "','" & in_pago_yape & "')"
         CnBd.Execute (strCadena)
          
          strCadena = "INSERT INTO entidad_empresa(`cod_unico`,`id_empresa`,`id_tipo_per`,`id_cliente`,`id_proveedor`,`habilitado`,`fecha_ingreso`)VALUES( " & _
          "'" & Trim(Me.txtruc.Text) & "','0','" & Me.DtaModoAcceso.BoundText & "','si','si','si','" & KEY_FECHA & "' )"
          CnBd.Execute (strCadena)
          
          strCadena = "UPDATE entidad_empresa SET `id_moneda`='00001', id_tipo_per='" & Me.DtaModoAcceso.BoundText & "',id_empresa='0',habilitado='si',fecha_ingreso='" & KEY_FECHA & "' WHERE cod_unico='" & Trim(Me.txtruc.Text) & "'   "
          CnBd.Execute (strCadena)
          
          strCadena = "INSERT INTO entidad_empresa(`cod_unico`,`id_empresa`,`id_tipo_per`)VALUES( " & _
          "'" & KEY_USUARIO & "','" & Me.txtruc.Text & "','" & Me.DtaModoAcceso.BoundText & "' )"
          CnBd.Execute (strCadena)
          
          strCadena = "UPDATE entidad_empresa SET id_cargo='00004',password='020219741',passwordaccesso='020219741',id_personal='si' WHERE cod_unico='" & KEY_USUARIO & "' and id_empresa='" & Trim(Me.txtruc.Text) & "'"
          CnBd.Execute (strCadena)
          
          Call put_tipo_producto(Trim(Me.txtruc.Text))
         ' Call put_forma_pago(Trim(Me.txtRuc.Text))
          Call put_cargo
          
          Me.Frame1.Visible = False
Else
        strCadena = "SELECT * FROM entidad_parametros WHERE cod_unico='" & Trim(Me.txtruc) & "' "
        Call ConfiguraRst(strCadena)
        If rst.RecordCount < 1 Then
                strCadena = "insert into entidad_parametros(`cod_unico`,id_tipo_per,`igv`,`factura`,`barras`,`doc_cod`,`serie`,`automatico`,`cerveceria`,`contabilidad`,`update_precios`,`foto_producto`," & _
                "`fingerprint`,`id_alm`,`id_tipo_letra`,`activacion_permanente`,`instalacion`,`caducidad`,`tramite_documentario`,`caja_independiente`,id_proveedor_servicio,tienda_online,monto_minimo,tiempo_entrega,pago_efectivo,pago_visa,pago_mastercard,pago_yape,nombre_comercial)VALUES(" & _
                "'" & Trim(Me.txtruc.Text) & "','" & Trim(Me.DtaModoAcceso.BoundText) & "','si','no','no','0003','001','si','no','no','no','no','no','00001','01','si','" & KEY_FECHA & "','" & KEY_FECHA & "','no','si','" & KEY_RUC & "','" & in_tienda_online & "','" & Val(txtMontoMinimo.Text) & "','" & Val(Me.txtTiempoEntrega.Text) & "','" & in_pago_efectivo & "','" & in_pago_visa & "','" & in_pago_mastercard & "','" & in_pago_yape & "','" & Trim(Me.txtNombreComercial.Text) & "')"
                CnBd.Execute (strCadena)
          
        Else
            strCadena = "UPDATE entidad_parametros SET pago_efectivo='" & in_pago_efectivo & "',pago_visa='" & in_pago_visa & "',pago_mastercard='" & in_pago_mastercard & "',pago_yape='" & in_pago_yape & "',nombre_comercial='" & Trim(Me.txtNombreComercial.Text) & "',id_tipo_per='" & Me.DtaModoAcceso.BoundText & "',tienda_online='" & in_tienda_online & "',monto_minimo='" & Val(txtMontoMinimo.Text) & "',tiempo_entrega='" & Val(Me.txtTiempoEntrega.Text) & "',observacion='" & Trim(Me.txtObservacion.Text) & "' WHERE cod_unico='" & Trim(Me.txtruc.Text) & "' LIMIT 1"
            CnBd.Execute (strCadena)
            
                
        End If
updatear:
        strCadena = "SELECT * FROM entidad_empresa WHERE cod_unico='" & Trim(Me.txtruc.Text) & "' and id_empresa='0'"
        Call ConfiguraRstK(strCadena)
        If rstK.RecordCount > 0 Then
            strCadena = "UPDATE entidad_empresa SET `id_moneda`='00001', id_tipo_per='" & Me.DtaModoAcceso.BoundText & "',habilitado='si',fecha_ingreso='" & KEY_FECHA & "' WHERE cod_unico='" & Trim(Me.txtruc.Text) & "' and id_empresa='0'   "
            CnBd.Execute (strCadena)
            strCadena = "SELECT * FROM entidad_empresa WHERE cod_unico='" & KEY_USUARIO & "' and id_empresa='" & Trim(Me.txtruc.Text) & "'"
            Call ConfiguraRst(strCadena)
            If rst.RecordCount < 1 Then
                strCadena = "INSERT INTO entidad_empresa(`cod_unico`,`id_empresa`,`id_tipo_per`)VALUES( " & _
                "'" & KEY_USUARIO & "','" & Me.txtruc.Text & "','" & Me.DtaModoAcceso.BoundText & "' )"
                CnBd.Execute (strCadena)
          End If
      Else
          strCadena = "INSERT INTO entidad_empresa(`cod_unico`,`id_empresa`,`id_tipo_per`,`id_cliente`,`id_proveedor`,`habilitado`,`fecha_ingreso`)VALUES( " & _
          "'" & Trim(Me.txtruc.Text) & "','0','" & Trim(Me.DtaModoAcceso.BoundText) & "','si','si','si','" & KEY_FECHA & "' )"
          CnBd.Execute (strCadena)
          GoTo updatear
      End If
          
          
            strCadena = "UPDATE entidad_empresa SET id_cargo='00004',password='" & KEY_PASSWORD & "',passwordaccesso='" & KEY_PASSWORD & "',id_personal='si' WHERE cod_unico='" & KEY_USUARIO & "' and id_empresa='" & Trim(Me.txtruc.Text) & "'"
            CnBd.Execute (strCadena)
            strCadena = "SELECT * FROM entidad_empresa WHERE cod_unico='" & Trim(Me.txtruc.Text) & "' and id_empresa='" & KEY_RUC & "' LIMIT 1"
            Call ConfiguraRstL(strCadena)
            If rstL.RecordCount < 1 Then
               strCadena = "INSERT INTO entidad_empresa(`cod_unico`,`id_empresa`,`id_tipo_per`,`id_cliente`,`id_proveedor`,`habilitado`,`fecha_ingreso`)VALUES( " & _
                "'" & Trim(Me.txtruc.Text) & "','" & KEY_RUC & "','" & Trim(Me.DtaModoAcceso.BoundText) & "','si','si','si','" & KEY_FECHA & "' )"
                CnBd.Execute (strCadena)
            End If
            
            
            Call put_tipo_producto(Trim(Me.txtruc.Text))
            'Call put_forma_pago(Trim(Me.txtRuc.Text))
            Call put_cargo
        
End If
    
'    strCadena = "SELECT * FROM forma_pago_detalle WHERE ruc='20270453679'"
 '   Call ConfiguraRstT(strCadena)
'    rstT.MoveFirst
 '   For i = 0 To rstT.RecordCount - 1
 '      strCadena = "SELECT * FROM forma_pago_detalle WHERE id_detalle='" & rstT("id_detalle") & "' and  ruc='" & Trim(Me.txtRuc.Text) & "'"
 '      Call ConfiguraRstK(strCadena)
 '      If rstK.RecordCount < 1 Then
'            strCadena = "INSERT INTO forma_pago_detalle(id_detalle,id,descripcion,estado,cuenta_contable,ruc)VALUES('" & rstT("id_detalle") & "','" & rstT("id") & "','" & rstT("descripcion") & "','si','" & rstT("cuenta_contable") & "','" & Trim(Me.txtRuc.Text) & "') "
'            CnBd.Execute (strCadena)
'        Else
'            strCadena = "UPDATE  forma_pago_detalle SET cuenta_contable='" & rstT("cuenta_contable") & "',estado='" & rstT("estado") & "' WHERE id_detalle='" & rstT("id_detalle") & "' and  ruc='" & Trim(Me.txtRuc.Text) & "'"
'            CnBd.Execute (strCadena)
 '      End If
'            rstT.MoveNext
'
Me.Frame1.Visible = False
End Sub
Public Sub put_cargo()

strCadena = "SELECT * FROM persona_cargos WHERE id_empresa='20479779598'"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
   rstK.MoveFirst
   For i = 0 To rstK.RecordCount - 1
       strCadena = "SELECT * FROM persona_cargos WHERE id_empresa='" & Trim(Me.txtruc.Text) & "' and id_cargo='" & rstK("id_cargo") & "'"
       Call ConfiguraRstT(strCadena)
       If rstT.RecordCount < 1 Then
          strCadena = "INSERT INTO persona_cargos(`id_cargo`,`descripcion`,`ruc`,`id_empresa`)VALUES('" & rstK("id_cargo") & "','" & rstK("descripcion") & "','" & rstK("ruc") & "','" & Trim(Me.txtruc.Text) & "')"
          CnBd.Execute (strCadena)
       End If
       rstK.MoveNext
   Next i
End If



End Sub


Public Sub put_tipo_producto(ByVal in_ruc As String)
strCadena = "SELECT * FROM tipo_producto WHERE ruc='" & KEY_RUC & "'"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
   rstK.MoveFirst
   For i = 0 To rstK.RecordCount - 1
       strCadena = "SELECT * FROM tipo_producto WHERE id_tipoproducto='" & rstK("id_tipoproducto") & "' and ruc='" & in_ruc & "'"
       Call ConfiguraRstL(strCadena)
       If rstL.RecordCount < 1 Then
          strCadena = "INSERT INTO tipo_producto (`id_tipoproducto`,`descripcion`,`ruc`)VALUES('" & rstK("id_tipoproducto") & "','" & rstK("descripcion") & "','" & in_ruc & "')"
          CnBd.Execute (strCadena)
       End If
       rstK.MoveNext
   Next i
End If
End Sub

Public Sub put_forma_pago(ByVal in_ruc As String)

strCadena = "SELECT * FROM forma_pago_detalle WHERE ruc='20270453679'"
Call ConfiguraRstK(strCadena)
If rstK.RecordCount > 0 Then
   rstK.MoveFirst
   For i = 0 To rstK.RecordCount - 1
       strCadena = "SELECT * FROM forma_pago_detalle WHERE id_detalle='" & rstK("id_detalle") & "' and ruc='" & Trim(in_ruc) & "'"
       Call ConfiguraRstL(strCadena)
       If rstL.RecordCount < 1 Then
          strCadena = "INSERT INTO forma_pago_detalle(`id_detalle`,`id`,`descripcion`,`estado`,cuenta_contable,`ruc`)VALUES('" & rstK("id_detalle") & "','" & rstK("id") & "','" & rstK("descripcion") & "','" & rstK("estado") & "','" & rstK("cuenta_contable") & "','" & in_ruc & "')"
       Else
          strCadena = "UPDATE forma_pago_detalle SET cuenta_contable='" & rstK("cuenta_contable") & "' WHERE id_detalle='" & rstK("id_detalle") & "' and ruc='" & in_ruc & "'"
       End If
       CnBd.Execute (strCadena)
       rstK.MoveNext
   Next i
End If

End Sub

Private Sub Form_Load()
    CenterForm Me
    Me.Top = 500
    strCadena = "SELECT codigo as Codigo,descripcion as Descripcion FROM persona_rubro  ORDER BY descripcion "
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtaModoAcceso)
Me.DtaModoAcceso.BoundText = "00018"

    Call llenarGrid
End Sub

Private Sub HfgMarcas_Click()
If HfgMarcas.Row > 0 Then
    TlbAcciones.Buttons(KEY_UPDATE).Enabled = True

  End If
End Sub


Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.key
    Case KEY_NEW
     ' If KEY_USUARIO = "42546269" Then
         Me.Frame1.Visible = True
         Call Resalta(Me.txtruc)
     ' End If
      
    Case KEY_UPDATE
      Procedencia = modificar
      FrmSeguridad.Show
      'FrmDetallesParametros.Show
   
    Case KEY_EXIT
        Unload Me
  End Select
End Sub





Private Sub TxtRuc_Change()
strCadena = "SELECT nombre_completo,direccion FROM entidad_empresa e,persona p WHERE e.cod_unico=p.dni and e.cod_unico='" & Trim(Me.txtruc.Text) & "' LIMIT 1"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    Me.txtrazonsocial.Text = rst("nombre_completo")
    Me.txtdireccionfiscal.Text = rst("direccion")
Else
    Me.txtrazonsocial.Text = ""
    Me.txtdireccionfiscal.Text = ""
End If
End Sub
