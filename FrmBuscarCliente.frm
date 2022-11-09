VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmBuscarCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar Clientes"
   ClientHeight    =   2595
   ClientLeft      =   585
   ClientTop       =   750
   ClientWidth     =   7275
   Icon            =   "FrmBuscarCliente.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2595
   ScaleWidth      =   7275
   Begin TabDlg.SSTab SstBusca 
      Height          =   2175
      Left            =   127
      TabIndex        =   2
      Top             =   240
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   3836
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      ForeColor       =   12582912
      TabCaption(0)   =   "Nombre y Apellidos"
      TabPicture(0)   =   "FrmBuscarCliente.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "LblNombres"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "LblApellidos"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "TxtNombre"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "TxtApellido"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Dirección"
      TabPicture(1)   =   "FrmBuscarCliente.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "TxtDireccion"
      Tab(1).Control(1)=   "LblDireccion"
      Tab(1).ControlCount=   2
      Begin VB.TextBox TxtDireccion 
         ForeColor       =   &H00C00000&
         Height          =   350
         Left            =   -73440
         MaxLength       =   80
         TabIndex        =   3
         Top             =   840
         Width           =   4095
      End
      Begin VB.TextBox TxtApellido 
         ForeColor       =   &H8000000D&
         Height          =   350
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   1
         Top             =   1440
         Width           =   3975
      End
      Begin VB.TextBox TxtNombre 
         ForeColor       =   &H8000000D&
         Height          =   350
         Left            =   1680
         MaxLength       =   50
         TabIndex        =   0
         Top             =   840
         Width           =   3975
      End
      Begin VB.Label LblDireccion 
         AutoSize        =   -1  'True
         Caption         =   "Dirección :"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   -74760
         TabIndex        =   7
         Top             =   960
         Width           =   765
      End
      Begin VB.Label LblApellidos 
         AutoSize        =   -1  'True
         Caption         =   "Apellidos:"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   675
      End
      Begin VB.Label LblNombres 
         AutoSize        =   -1  'True
         Caption         =   "Nombres:"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   675
      End
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   1890
      Left            =   6240
      TabIndex        =   8
      Top             =   375
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   3334
      BandCount       =   1
      ForeColor       =   -2147483635
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      Orientation     =   1
      EmbossPicture   =   -1  'True
      _CBWidth        =   900
      _CBHeight       =   1890
      _Version        =   "6.0.8169"
      Child1          =   "TlbAcciones"
      MinHeight1      =   840
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   780
         Left            =   30
         TabIndex        =   4
         Top             =   30
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   1376
         ButtonWidth     =   1429
         ButtonHeight    =   1376
         Style           =   1
         ImageList       =   "ImgIconos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "  Buscar   "
               Key             =   "(Buscar)"
               Object.ToolTipText     =   "Buscar"
               ImageKey        =   "(Buscar)"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               ImageKey        =   "(Cancelar)"
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Cancelar"
               Key             =   "(Cancelar)"
               Object.ToolTipText     =   "Cancelar"
               ImageKey        =   "(Cancelar)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   6120
      Top             =   1800
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
            Picture         =   "FrmBuscarCliente.frx":0342
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBuscarCliente.frx":0796
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBuscarCliente.frx":0AB6
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBuscarCliente.frx":0F0A
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBuscarCliente.frx":135E
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBuscarCliente.frx":167E
            Key             =   "(Cancelar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBuscarCliente.frx":199E
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBuscarCliente.frx":1CBE
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmBuscarCliente.frx":1FDE
            Key             =   "(Buscar)"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmBuscarCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
CenterForm Me
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Key
    Case KEY_BROWSER
      Me.Hide
      
    Case KEY_CANCEL
      Unload Me
  End Select
End Sub

