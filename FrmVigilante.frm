VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmVigilante 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control Ocurrencias"
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   420
   ClientWidth     =   15045
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   15045
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdNuevo 
      Caption         =   "Nuevo"
      Height          =   765
      Left            =   12480
      Picture         =   "FrmVigilante.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   2400
      Width           =   735
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8655
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   15266
      _Version        =   393216
      TabHeight       =   520
      ForeColor       =   8388608
      TabCaption(0)   =   "UNIDADES TRANSPORTE"
      TabPicture(0)   =   "FrmVigilante.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Shape1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label7"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Shape2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label10"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label11"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "DtpFin"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "HfOcurrencias"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "SSTab2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "DtPInicio"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Command1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "TxtPersona"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Command4"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "TxtPlaca"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Command5"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "Tab 1"
      TabPicture(1)   =   "FrmVigilante.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "FrmVigilante.frx":047A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.CommandButton Command5 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   13440
         TabIndex        =   36
         Top             =   4560
         Width           =   735
      End
      Begin VB.TextBox TxtPlaca 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   11280
         TabIndex        =   35
         Top             =   4560
         Width           =   2055
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   9480
         TabIndex        =   33
         Top             =   4560
         Width           =   735
      End
      Begin VB.TextBox TxtPersona 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7320
         TabIndex        =   32
         Top             =   4560
         Width           =   2055
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   4560
         TabIndex        =   28
         Top             =   4560
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker DtPInicio 
         Height          =   375
         Left            =   1200
         TabIndex        =   27
         Top             =   4560
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   121176065
         CurrentDate     =   40976
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   3615
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   14205
         _ExtentX        =   25056
         _ExtentY        =   6376
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         ForeColor       =   8388608
         TabCaption(0)   =   "ORDEN COMPRA"
         TabPicture(0)   =   "FrmVigilante.frx":0496
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Label1"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label8"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label9"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "Label3"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "DtCSalidaIngreso"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "TxtSerie"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "lblPlacaTractor"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "TxtNumero"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "Frame1"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "TxtKilometraje"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).ControlCount=   10
         TabCaption(1)   =   "CISTERNAS-DISTRIBUIDORES"
         TabPicture(1)   =   "FrmVigilante.frx":04B2
         Tab(1).ControlEnabled=   0   'False
         Tab(1).ControlCount=   0
         Begin VB.TextBox TxtKilometraje 
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
            Height          =   315
            Left            =   1560
            TabIndex        =   24
            Top             =   1440
            Width           =   1815
         End
         Begin VB.Frame Frame1 
            Caption         =   "PERSONAL ACOMPAÑANTE"
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
            Height          =   2775
            Left            =   6360
            TabIndex        =   16
            Top             =   720
            Width           =   6735
            Begin VB.CommandButton CmdQuitarPersonal 
               BackColor       =   &H008080FF&
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   5280
               Style           =   1  'Graphical
               TabIndex        =   23
               Top             =   960
               Visible         =   0   'False
               Width           =   375
            End
            Begin VB.CommandButton CmdSave 
               Caption         =   "Grabar"
               Height          =   765
               Left            =   5760
               Picture         =   "FrmVigilante.frx":04CE
               Style           =   1  'Graphical
               TabIndex        =   21
               Top             =   1800
               Width           =   735
            End
            Begin VB.CommandButton cmdAgregar 
               Caption         =   "Agregar"
               Height          =   285
               Left            =   5760
               TabIndex        =   20
               Top             =   360
               Width           =   855
            End
            Begin VB.TextBox TxtDescripcion 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   1230
               TabIndex        =   19
               Top             =   360
               Width           =   4335
            End
            Begin VB.TextBox TxtCPersona 
               Appearance      =   0  'Flat
               Height          =   285
               Left            =   120
               TabIndex        =   18
               Top             =   360
               Width           =   1095
            End
            Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfPersonas 
               Height          =   1695
               Left            =   120
               TabIndex        =   17
               Top             =   840
               Width           =   5535
               _ExtentX        =   9763
               _ExtentY        =   2990
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
            Begin MSComctlLib.ImageList ImgIconos 
               Left            =   2880
               Top             =   3030
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
                     Picture         =   "FrmVigilante.frx":0D98
                     Key             =   "(Nuevo)"
                  EndProperty
                  BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVigilante.frx":11EC
                     Key             =   "(Modificar)"
                  EndProperty
                  BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVigilante.frx":150C
                     Key             =   "(Eliminar)"
                  EndProperty
                  BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVigilante.frx":1960
                     Key             =   "(Salir)"
                  EndProperty
                  BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVigilante.frx":1DB4
                     Key             =   "(Aceptar)"
                  EndProperty
                  BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVigilante.frx":20D4
                     Key             =   "(Cancelar)"
                  EndProperty
                  BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVigilante.frx":23F4
                     Key             =   "(Quitar)"
                  EndProperty
                  BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVigilante.frx":2714
                     Key             =   "(Agregar)"
                  EndProperty
                  BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                     Picture         =   "FrmVigilante.frx":2A34
                     Key             =   "(Buscar)"
                  EndProperty
               EndProperty
            End
         End
         Begin VB.TextBox TxtNumero 
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
            Height          =   315
            Left            =   2400
            TabIndex        =   9
            Top             =   480
            Width           =   1335
         End
         Begin VB.Frame lblPlacaTractor 
            Caption         =   "DATOS DE TRANSPORTE"
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
            Height          =   1575
            Left            =   120
            TabIndex        =   5
            Top             =   1920
            Width           =   5415
            Begin VB.Label LblChofer 
               AutoSize        =   -1  'True
               Caption         =   "--"
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
               Left            =   960
               TabIndex        =   12
               Top             =   1320
               Width           =   150
            End
            Begin VB.Label LblTractor 
               AutoSize        =   -1  'True
               Caption         =   "--"
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
               Left            =   960
               TabIndex        =   11
               Top             =   360
               Width           =   150
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Chofer :"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   120
               TabIndex        =   10
               Top             =   1320
               Width           =   555
            End
            Begin VB.Label LblCisterna 
               AutoSize        =   -1  'True
               Caption         =   "--"
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
               Left            =   960
               TabIndex        =   8
               Top             =   840
               Width           =   150
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Cisterna :"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   120
               TabIndex        =   7
               Top             =   840
               Width           =   660
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Tractor:"
               ForeColor       =   &H00800000&
               Height          =   195
               Left            =   120
               TabIndex        =   6
               Top             =   360
               Width           =   555
            End
         End
         Begin VB.TextBox TxtSerie 
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
            Height          =   315
            Left            =   1560
            TabIndex        =   0
            Top             =   480
            Width           =   735
         End
         Begin MSDataListLib.DataCombo DtCSalidaIngreso 
            Height          =   330
            Left            =   1560
            TabIndex        =   13
            Top             =   960
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   582
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            ForeColor       =   4194304
            Text            =   "DataCombo1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "KM"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   3480
            TabIndex        =   25
            Top             =   1440
            Width           =   240
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Lectura Inicial:"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   240
            TabIndex        =   15
            Top             =   1440
            Width           =   1035
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Ingreso/ Salida:"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   120
            TabIndex        =   14
            Top             =   960
            Width           =   1125
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nº Orden Compra:"
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   120
            TabIndex        =   4
            Top             =   480
            Width           =   1290
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfOcurrencias 
         Height          =   3375
         Left            =   240
         TabIndex        =   2
         Top             =   5160
         Width           =   14175
         _ExtentX        =   25003
         _ExtentY        =   5953
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
      Begin MSComCtl2.DTPicker DtpFin 
         Height          =   375
         Left            =   3000
         TabIndex        =   29
         Top             =   4560
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   121176065
         CurrentDate     =   40976
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PLACA:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   10320
         TabIndex        =   34
         Top             =   4680
         Width           =   600
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PERSONA:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   6360
         TabIndex        =   31
         Top             =   4680
         Width           =   855
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         BorderStyle     =   3  'Dot
         Height          =   615
         Left            =   6240
         Top             =   4440
         Width           =   8175
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "a"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   2760
         TabIndex        =   30
         Top             =   4680
         Width           =   90
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FECHA :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   480
         TabIndex        =   26
         Top             =   4680
         Width           =   660
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderStyle     =   3  'Dot
         Height          =   615
         Left            =   240
         Top             =   4440
         Width           =   5895
      End
   End
End
Attribute VB_Name = "FrmVigilante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede
Private Sub llenarOcurrenciasFecha(ByVal Grilla As MSHFlexGrid)
On Error GoTo salir
strCadena = "SELECT     Ocurrencias.idOcurrencia, Ocurrencias.Fecha, Ocurrencias.Hora, SalidaIngreso.descripcion,Ocurrencias.idOrden," & _
" Persona.NombrePersona,Seguridad.usuario FROM Ocurrencias INNER JOIN Persona ON Ocurrencias.cPersona = Persona.cPersona INNER JOIN " & _
" SalidaIngreso ON Ocurrencias.idsalidaIngreso = SalidaIngreso.id INNER JOIN Seguridad ON Ocurrencias.cVigilante = Seguridad.IdUsuario WHERE fecha>='" & CVDate(Me.DtpInicio.Value) & "' and fecha<='" & CVDate(Me.DtpFin.Value) & "' ORDER BY Ocurrencias.idOcurrencia DESC"
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
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 1200
           Grilla.ColWidth(1) = 1200
           Grilla.ColWidth(2) = 1200
           Grilla.ColWidth(3) = 1200
           Grilla.ColWidth(4) = 5000
           Grilla.ColWidth(5) = 2000
       Next
         cabecera = "CODIGO" & vbTab & "FECHA" & vbTab & "HORA" & vbTab & "SAL/ING" & vbTab & "ENTIDAD" & vbTab & "VIJILANTE"
         Grilla.AddItem cabecera
         For k = 0 To 5
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 1 To rst.RecordCount
             If rst("idOrden") > 0 Then
                Entidad = ""
                strCadena = "SELECT MiTransporte.marca, MiTransporte.placa FROM OrdenCompra INNER JOIN MiTransporte ON OrdenCompra.id_Transporte = MiTransporte.id_Transporte WHERE idOrden='" & rst("idOrden") & "'"
                Call ConfiguraTemporal(strCadena)
                If rstTemporal.RecordCount > 0 Then
                    Entidad = rstTemporal("marca") + "-" + rstTemporal("placa")
                 End If
                Set rstTemporal = Nothing
                strCadena = "SELECT MiTransporte.placa FROM OrdenCompra INNER JOIN MiTransporte ON OrdenCompra.id_Cisterna = MiTransporte.id_Transporte WHERE idOrden='" & rst("idOrden") & "'"
                Call ConfiguraTemporal(strCadena)
                If rstTemporal.RecordCount > 0 Then
                    Entidad = Entidad + "-" + rstTemporal("placa")
                End If
                Set rstTemporal = Nothing
            Else
                Entidad = rst("NombrePersona")
             End If
             Fila = Fila & rst("idOcurrencia") & vbTab & rst("Fecha") & vbTab & rst("Hora") & vbTab & rst("descripcion") & vbTab & Entidad & vbTab & rst("usuario")
             Grilla.AddItem Fila
             Fila = ""
        rst.MoveNext
        Next i
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub
Private Sub llenarOcurrenciasPersona(ByVal Grilla As MSHFlexGrid)
On Error GoTo salir
strCadena = "SELECT     Ocurrencias.idOcurrencia, Ocurrencias.Fecha, Ocurrencias.Hora, SalidaIngreso.descripcion,Ocurrencias.idOrden," & _
" Persona.NombrePersona,Seguridad.usuario FROM Ocurrencias INNER JOIN Persona ON Ocurrencias.cPersona = Persona.cPersona INNER JOIN " & _
" SalidaIngreso ON Ocurrencias.idsalidaIngreso = SalidaIngreso.id INNER JOIN Seguridad ON Ocurrencias.cVigilante = Seguridad.IdUsuario WHERE fecha>='" & CVDate(Me.DtpInicio.Value) & "' and fecha<='" & CVDate(Me.DtpFin.Value) & "' ORDER BY Ocurrencias.idOcurrencia DESC"
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
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 1200
           Grilla.ColWidth(1) = 1200
           Grilla.ColWidth(2) = 1200
           Grilla.ColWidth(3) = 1200
           Grilla.ColWidth(4) = 5000
           Grilla.ColWidth(5) = 2000
       Next
         cabecera = "CODIGO" & vbTab & "FECHA" & vbTab & "HORA" & vbTab & "SAL/ING" & vbTab & "ENTIDAD" & vbTab & "VIJILANTE"
         Grilla.AddItem cabecera
         For k = 0 To 5
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 1 To rst.RecordCount
             If rst("idOrden") > 0 Then
                Entidad = ""
                strCadena = "SELECT MiTransporte.marca, MiTransporte.placa FROM OrdenCompra INNER JOIN MiTransporte ON OrdenCompra.id_Transporte = MiTransporte.id_Transporte WHERE idOrden='" & rst("idOrden") & "'"
                Call ConfiguraTemporal(strCadena)
                If rstTemporal.RecordCount > 0 Then
                    Entidad = rstTemporal("marca") + "-" + rstTemporal("placa")
                 End If
                Set rstTemporal = Nothing
                strCadena = "SELECT MiTransporte.placa FROM OrdenCompra INNER JOIN MiTransporte ON OrdenCompra.id_Cisterna = MiTransporte.id_Transporte WHERE idOrden='" & rst("idOrden") & "'"
                Call ConfiguraTemporal(strCadena)
                If rstTemporal.RecordCount > 0 Then
                    Entidad = Entidad + "-" + rstTemporal("placa")
                End If
                Set rstTemporal = Nothing
            Else
                Entidad = rst("NombrePersona")
             End If
             Fila = Fila & rst("idOcurrencia") & vbTab & rst("Fecha") & vbTab & rst("Hora") & vbTab & rst("descripcion") & vbTab & Entidad & vbTab & rst("usuario")
             Grilla.AddItem Fila
             Fila = ""
        rst.MoveNext
        Next i
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub

Private Sub CmdAgregar_Click()
Dim Idocurrencia As Double
If Me.TxtCPersona.Text = "" Or Me.txtdescripcion.Text = "" Then
    MsgBox "Ingrese un Personal registrado", vbInformation, "Mensaje al Usuario"
    Call Resalta(Me.TxtCPersona)
    Exit Sub
Else
    Idocurrencia = IdInsert("Ocurrencias") + 1
    strCadena = "INSERT INTO OcurrenciasPersonal(idOcurrencia,cPersona)VALUES('" & Idocurrencia & "','" & Trim(Me.TxtCPersona.Text) & "')"
    CnBd.Execute (strCadena)
     
    Call llenarGrid(Me.HfPersonas, Idocurrencia)
    Me.TxtCPersona.Text = ""
    Me.txtdescripcion.Text = ""
    Call Resalta(Me.TxtCPersona)
    
End If
End Sub
Private Sub llenarGrid(ByVal Grilla As MSHFlexGrid, ByVal Idocurrencia As Double)
On Error GoTo salir
strCadena = "SELECT     OcurrenciasPersonal.idDetalle, OcurrenciasPersonal.cPersona,Persona.Per_Ruc, Persona.NombrePersona, OcurrenciasPersonal.idOcurrencia " & _
" FROM OcurrenciasPersonal INNER JOIN Persona ON OcurrenciasPersonal.cPersona = Persona.cPersona WHERE idOcurrencia='" & Idocurrencia & "'"
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
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 600
           Grilla.ColWidth(2) = 1200
           Grilla.ColWidth(3) = 3000
           Grilla.ColWidth(4) = 0
       Next
         cabecera = "IDDETALLE" & vbTab & "CODIGO" & vbTab & "DNI/RUC" & vbTab & "PERSONAL" & vbTab & "IDOCURRENCIA"
         Grilla.AddItem cabecera
         For k = 0 To 4
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 1 To rst.RecordCount
             Fila = Fila & rst("idDetalle") & vbTab & rst("cPersona") & vbTab & rst("Per_Ruc") & vbTab & rst("NombrePersona") & vbTab & rst("IdOcurrencia")
             Grilla.AddItem Fila
             Fila = ""
        rst.MoveNext
        Next i
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub

Private Sub cmdnuevo_Click()
Call nuevo
Me.cmdsave.Enabled = False
End Sub

Private Sub CmdQuitarPersonal_Click()
strCadena = "DELETE FROM OcurrenciasPersonal WHERE idDetalle='" & Val(Me.HfPersonas.TextMatrix(Me.HfPersonas.Row, 0)) & "'"
CnBd.Execute (strCadena)
 
Call llenarGrid(Me.HfPersonas, Val(Me.HfPersonas.TextMatrix(Me.HfPersonas.Row, 4)))
End Sub

Private Sub cmdsave_Click()
Call Save
Me.cmdsave.Enabled = False
Me.cmdNuevo.Enabled = True
End Sub

Private Sub Command1_Click()
Call llenarOcurrenciasFecha(Me.HfOcurrencias)
End Sub

Private Sub Command2_Click()

End Sub
Private Sub Save()
Dim hora As String
Dim cPersona As Double
strCadena = "SELECT * FROM OrdenCompra WHERE serie='" & Trim(Me.TxtSerie.Text) & "' AND numero='" & Trim(Me.txtnumero.Text) & "' AND Ruc='" & KEY_RUC & "' AND doc_cod='0110'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    idOrden = rst("idOrden")
    cPersona = rst("cConductor")
Else
    idOrden = 0
End If
Set rst = Nothing
hora = Time$

    strCadena = "INSERT INTO Ocurrencias(idOrden,cPersona,idsalidaIngreso,Fecha,Hora,kilometraje,cVigilante)VALUES('" & idOrden & "','" & cPersona & "','" & Me.DtCSalidaIngreso.BoundText & "'" & _
    ",'" & KEY_FECHA & "','" & hora & "','" & Val(Me.TxtKilometraje.Text) & "','" & KEY_USUARIO & "')"
    CnBd.Execute (strCadena)
     
    Call llenarOcurrencias(Me.HfOcurrencias)
End Sub
Private Sub llenarOcurrencias(ByVal Grilla As MSHFlexGrid)
On Error GoTo salir
strCadena = "SELECT     Ocurrencias.idOcurrencia, Ocurrencias.Fecha, Ocurrencias.Hora, SalidaIngreso.descripcion,Ocurrencias.idOrden," & _
" Persona.NombrePersona,Seguridad.usuario FROM Ocurrencias INNER JOIN Persona ON Ocurrencias.cPersona = Persona.cPersona INNER JOIN " & _
" SalidaIngreso ON Ocurrencias.idsalidaIngreso = SalidaIngreso.id INNER JOIN Seguridad ON Ocurrencias.cVigilante = Seguridad.IdUsuario WHERE fecha='" & KEY_FECHA & "' ORDER BY Ocurrencias.idOcurrencia DESC"
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
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 1200
           Grilla.ColWidth(1) = 1200
           Grilla.ColWidth(2) = 1200
           Grilla.ColWidth(3) = 1200
           Grilla.ColWidth(4) = 5000
           Grilla.ColWidth(5) = 2000
       Next
         cabecera = "CODIGO" & vbTab & "FECHA" & vbTab & "HORA" & vbTab & "SAL/ING" & vbTab & "ENTIDAD" & vbTab & "VIJILANTE"
         Grilla.AddItem cabecera
         For k = 0 To 5
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 1 To rst.RecordCount
             If rst("idOrden") > 0 Then
                Entidad = ""
                strCadena = "SELECT MiTransporte.marca, MiTransporte.placa FROM OrdenCompra INNER JOIN MiTransporte ON OrdenCompra.id_Transporte = MiTransporte.id_Transporte WHERE idOrden='" & rst("idOrden") & "'"
                Call ConfiguraTemporal(strCadena)
                If rstTemporal.RecordCount > 0 Then
                    Entidad = rstTemporal("marca") + "-" + rstTemporal("placa")
                 End If
                Set rstTemporal = Nothing
                strCadena = "SELECT MiTransporte.placa FROM OrdenCompra INNER JOIN MiTransporte ON OrdenCompra.id_Cisterna = MiTransporte.id_Transporte WHERE idOrden='" & rst("idOrden") & "'"
                Call ConfiguraTemporal(strCadena)
                If rstTemporal.RecordCount > 0 Then
                    Entidad = Entidad + "-" + rstTemporal("placa")
                End If
                Set rstTemporal = Nothing
            Else
                Entidad = rst("NombrePersona")
             End If
             Fila = Fila & rst("idOcurrencia") & vbTab & rst("Fecha") & vbTab & rst("Hora") & vbTab & rst("descripcion") & vbTab & Entidad & vbTab & rst("usuario")
             Grilla.AddItem Fila
             Fila = ""
        rst.MoveNext
        Next i
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub

Private Sub Command3_Click()


End Sub

Private Sub nuevo()
    Me.TxtSerie.Text = ""
Me.txtnumero.Text = ""
Me.DtCSalidaIngreso.BoundText = 1
Me.LblTractor.Caption = ""
Me.LblCisterna.Caption = ""
Me.LblChofer.Caption = ""
Me.TxtKilometraje.Text = ""
  Idocurrencia = IdInsert("Ocurrencias")
  strCadena = "SELECT * FROM Ocurrencias WHERE idOcurrencia='" & Idocurrencia & "'"
  Call ConfiguraRst(strCadena)
  'If rst.RecordCount < 1 Then
    Call llenarGrid(Me.HfPersonas, Idocurrencia)
 ' End If
    Set rst = Nothing
    Call Resalta(Me.TxtSerie)
End Sub

Private Sub DtCSalidaIngreso_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.TxtKilometraje)
End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 116 Then
    Call Save
    Call nuevo
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 100
Me.DtpInicio.Value = KEY_FECHA
Me.DtpFin.Value = KEY_FECHA

    strCadena = "SELECT id as Codigo, descripcion as Descripcion FROM SalidaIngreso " & _
  " ORDER BY id ASC"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtCSalidaIngreso)
  Set rst = Nothing
  Idocurrencia = IdInsert("Ocurrencias")
  strCadena = "SELECT * FROM Ocurrencias WHERE idOcurrencia='" & Idocurrencia & "'"
  Call ConfiguraRst(strCadena)
  If rst.RecordCount < 1 Then
    Call llenarGrid(Me.HfPersonas, Idocurrencia)
  End If
  Set rst = Nothing
  Call llenarOcurrencias(Me.HfOcurrencias)
  Call nuevo
  Me.cmdsave.Enabled = False
End Sub

Private Sub Text1_Change()

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtSerie.Text = formato_item(Me.TxtSerie.Text, 4)
    Call Resalta(Me.txtnumero)
End If
End Sub

Private Sub HfPersonas_SelChange()
If Val(Me.HfPersonas.TextMatrix(Me.HfPersonas.Row, 0)) > 0 Then
    Me.CmdQuitarPersonal.Visible = True
End If
End Sub

Private Sub txtCpersona_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Me.TxtCPersona.Text = "" Then
        Procedencia = Selecionar
        FrmPersona.Show
        Exit Sub
    Else
        If Len(Me.TxtCPersona.Text) >= 4 Then
            strCadena = "SELECT * FROM Persona WHERE Per_Ruc='" & Trim(Me.TxtCPersona.Text) & "'"
            Call ConfiguraRst(strCadena)
            If rst.RecordCount > 0 Then
                Me.TxtCPersona.Text = rst("cPersona")
                Me.txtdescripcion.Text = rst("NombrePersona")
                Me.cmdagregar.SetFocus
            Else
               strCadena = "SELECT * FROM Persona WHERE cPersona='" & Trim(Me.TxtCPersona.Text) & "'"
               Call ConfiguraRst(strCadena)
               If rst.RecordCount > 0 Then
                    Me.TxtCPersona.Text = rst("cPersona")
                    Me.txtdescripcion.Text = rst("NombrePersona")
                    Me.cmdagregar.SetFocus
                Else
                Procedencia = Selecionar
                FrmPersona.Show
                Set rst = Nothing
                Exit Sub
               End If
                
            End If
        End If
    End If
Set rst = Nothing
End If
End Sub

Private Sub TxtKilometraje_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Val(Me.TxtKilometraje.Text) > 0 Then
        Call Resalta(Me.TxtCPersona)
    Else
        Call Resalta(Me.TxtKilometraje)
    End If
End If
End Sub

Private Sub TxtNumero_KeyPress(KeyAscii As Integer)
Dim idOrden As Double
If KeyAscii = 13 Then
    Me.txtnumero.Text = formato_item(Me.txtnumero.Text, 10)
    If Me.TxtSerie.Text <> "" Then
        strCadena = "SELECT * FROM OrdenCompra WHERE serie='" & Trim(Me.TxtSerie.Text) & "' AND numero='" & Trim(Me.txtnumero.Text) & "' AND Ruc='" & KEY_RUC & "' AND doc_cod='0110'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount > 0 Then
            idOrden = rst("IdOrden")
            Set rst = Nothing
            strCadena = "SELECT     MiTransporte.marca, MiTransporte.placa, Persona.NombrePersona, OrdenCompra.licencia " & _
            " FROM OrdenCompra INNER JOIN MiTransporte ON OrdenCompra.id_Transporte = MiTransporte.id_Transporte INNER JOIN " & _
            " Persona ON OrdenCompra.cConductor = Persona.cPersona WHERE IdOrden='" & idOrden & "'"
            Call ConfiguraRst(strCadena)
            If rst.RecordCount > 0 Then
               Me.LblTractor.Caption = rst("marca") + Space(2) + rst("placa")
                Me.LblChofer.Caption = rst("NombrePersona") + Space(2) + rst("licencia")
            End If
            Set rst = Nothing
            strCadena = "SELECT     MiTransporte.id_Transporte, MiTransporte.marca, MiTransporte.placa " & _
            " FROM OrdenCompra INNER JOIN MiTransporte ON OrdenCompra.id_Cisterna = MiTransporte.id_Transporte WHERE idOrden='" & idOrden & "'"
            Call ConfiguraRst(strCadena)
            If rst.RecordCount > 0 Then
                Me.LblCisterna.Caption = rst("placa")
            End If
            Set rst = Nothing
            Me.DtCSalidaIngreso.SetFocus
         Else
            Me.LblCisterna.Caption = ""
            Me.LblTractor.Caption = ""
            Me.LblChofer.Caption = ""
            MsgBox "Orden de Compra no registrada", vbInformation, "Consulte con el Administrador"
            Call Resalta(Me.txtnumero)
        End If
        
        
    End If
End If
End Sub

Private Sub TxtSerie_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.TxtSerie.Text = formato_item(Me.TxtSerie.Text, 4)
    Call Resalta(Me.txtnumero)
End If
End Sub
