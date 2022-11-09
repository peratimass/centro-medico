VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmCargosPersonal 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8565
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8610
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   8610
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtLinea 
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
      Left            =   1920
      TabIndex        =   22
      Top             =   7920
      Width           =   4815
   End
   Begin VB.Frame FrameFunciones 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8055
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
      Width           =   6975
      Begin VB.TextBox txtFuncion 
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
         Height          =   285
         Left            =   1680
         TabIndex        =   18
         Top             =   6600
         Width           =   4335
      End
      Begin VB.Frame FrameDetalle 
         BackColor       =   &H00FFFFFF&
         Height          =   5295
         Left            =   480
         TabIndex        =   3
         Top             =   600
         Visible         =   0   'False
         Width           =   6135
         Begin VB.Frame FrameDetalleFuncion 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4815
            Left            =   120
            TabIndex        =   4
            Top             =   360
            Visible         =   0   'False
            Width           =   4695
            Begin VB.TextBox txtCodigo 
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
               Height          =   285
               Left            =   1560
               TabIndex        =   10
               Top             =   480
               Visible         =   0   'False
               Width           =   1335
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
               Height          =   285
               Left            =   1560
               TabIndex        =   9
               Top             =   1005
               Width           =   3015
            End
            Begin VB.CommandButton cmdprocesar 
               Caption         =   "PROCESAR"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1560
               TabIndex        =   8
               Top             =   2400
               Width           =   2175
            End
            Begin VB.TextBox txtIdFuncion 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1560
               TabIndex        =   7
               Top             =   240
               Visible         =   0   'False
               Width           =   1335
            End
            Begin VB.CommandButton Command1 
               Caption         =   "CERRAR"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   1560
               TabIndex        =   6
               Top             =   2880
               Width           =   2175
            End
            Begin MSDataListLib.DataCombo dtcArea 
               Height          =   330
               Left            =   1560
               TabIndex        =   5
               Top             =   1440
               Width           =   3015
               _ExtentX        =   5318
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
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "CODIGO :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Left            =   615
               TabIndex        =   13
               Top             =   600
               Visible         =   0   'False
               Width           =   705
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "DESCRIPCION :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Left            =   210
               TabIndex        =   12
               Top             =   1080
               Width           =   1110
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "AREA :"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Left            =   840
               TabIndex        =   11
               Top             =   1560
               Width           =   480
            End
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfFunciones 
            Height          =   4335
            Left            =   120
            TabIndex        =   14
            Top             =   840
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   7646
            _Version        =   393216
            ForeColor       =   8388608
            FixedCols       =   0
            ForeColorFixed  =   8388608
            BackColorBkg    =   16777215
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
         Begin ComCtl3.CoolBar CoolBar2 
            Height          =   2535
            Left            =   4920
            TabIndex        =   15
            Top             =   600
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   4471
            BandCount       =   1
            BackColor       =   16777215
            ImageList       =   "ImageList1"
            Orientation     =   1
            _CBWidth        =   810
            _CBHeight       =   2535
            _Version        =   "6.0.8169"
            MinHeight1      =   750
            Width1          =   2475
            NewRow1         =   0   'False
            BandStyle1      =   1
            Begin ComctlLib.Toolbar Toolbar1 
               Height          =   2490
               Left            =   0
               TabIndex        =   16
               Top             =   0
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   7250
               ButtonWidth     =   1323
               ButtonHeight    =   1376
               ImageList       =   "ImageList1"
               _Version        =   327682
               BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
                  NumButtons      =   4
                  BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                     Caption         =   "Nuevo"
                     Key             =   "(Nuevo)"
                     Object.Tag             =   ""
                     ImageKey        =   "(Nuevo)"
                  EndProperty
                  BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                     Caption         =   "Modificar"
                     Key             =   "(Modificar)"
                     Object.Tag             =   ""
                     ImageKey        =   "(Modificar)"
                  EndProperty
                  BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                     Caption         =   "Salir"
                     Key             =   "(Salir)"
                     Object.Tag             =   ""
                     ImageKey        =   "(Salir)"
                  EndProperty
                  BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
                     Key             =   ""
                     Object.Tag             =   ""
                     Style           =   3
                  EndProperty
               EndProperty
            End
         End
         Begin VB.Label lblfunciones 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "FUNCIONES A EJECUTAR"
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
            Left            =   240
            TabIndex        =   17
            Top             =   240
            Width           =   1830
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfActividades 
         Height          =   5655
         Left            =   480
         TabIndex        =   19
         Top             =   600
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   9975
         _Version        =   393216
         ForeColor       =   8388608
         FixedCols       =   0
         ForeColorFixed  =   8388608
         BackColorBkg    =   16777215
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
      Begin VB.Label lblcargo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "FUNCIONES A EJECUTAR"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   600
         TabIndex        =   21
         Top             =   240
         Width           =   1830
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BUSCAR :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   165
         Left            =   645
         TabIndex        =   20
         Top             =   6600
         Width           =   705
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00DFDFE0&
         BorderColor     =   &H00808080&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   555
         Left            =   480
         Top             =   6480
         Width           =   6135
      End
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   7080
      Left            =   7440
      TabIndex        =   0
      Top             =   480
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   12488
      BandCount       =   1
      ImageList       =   "ImageList1"
      Orientation     =   1
      _CBWidth        =   1005
      _CBHeight       =   7080
      _Version        =   "6.0.8169"
      MinHeight1      =   945
      Width1          =   7020
      NewRow1         =   0   'False
      BandStyle1      =   1
      Begin ComctlLib.Toolbar TlbAcciones 
         Height          =   5730
         Left            =   105
         TabIndex        =   1
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   10107
         ButtonWidth     =   1614
         ButtonHeight    =   1429
         ImageList       =   "ImageList1"
         _Version        =   327682
         BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
            NumButtons      =   6
            BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Nuevo"
               Key             =   "(Nuevo)"
               Object.Tag             =   ""
               ImageKey        =   "(Nuevo)"
            EndProperty
            BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Modificar"
               Key             =   "(Modificar)"
               Object.Tag             =   ""
               ImageKey        =   "(Modificar)"
            EndProperty
            BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Funciones"
               Key             =   "(Funciones)"
               Object.Tag             =   ""
               ImageKey        =   "(Funciones)"
            EndProperty
            BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Eliminar"
               Key             =   "(Eliminar)"
               Object.Tag             =   ""
               ImageKey        =   "(Eliminar)"
            EndProperty
            BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Caption         =   "Salir"
               Key             =   "(Salir)"
               Object.Tag             =   ""
               ImageKey        =   "(Salir)"
            EndProperty
            BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
               Key             =   ""
               Object.Tag             =   ""
               Style           =   4
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   4320
      Top             =   4800
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
            Picture         =   "FrmCargosPersonal.frx":0000
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCargosPersonal.frx":0454
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCargosPersonal.frx":0774
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCargosPersonal.frx":0BC8
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCargosPersonal.frx":101C
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCargosPersonal.frx":133C
            Key             =   "(Cancelar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCargosPersonal.frx":165C
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCargosPersonal.frx":197C
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmCargosPersonal.frx":1C9C
            Key             =   "(Buscar)"
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgLinea 
      Height          =   7095
      Left            =   240
      TabIndex        =   23
      Top             =   480
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   12515
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
   Begin VB.Label LblFecha 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BUSCAR :"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   165
      Left            =   405
      TabIndex        =   25
      Top             =   8040
      Width           =   705
   End
   Begin VB.Label LblEmpresa 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ROLES USUARIOS"
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
      Left            =   225
      TabIndex        =   24
      Top             =   120
      Width           =   1485
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   6960
      Top             =   8280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   5
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmCargosPersonal.frx":1FBC
            Key             =   "(Funciones)"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmCargosPersonal.frx":2C0E
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmCargosPersonal.frx":2F28
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmCargosPersonal.frx":3242
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "FrmCargosPersonal.frx":4D94
            Key             =   "(Nuevo)"
         EndProperty
      EndProperty
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00DFDFE0&
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   555
      Left            =   240
      Top             =   7800
      Width           =   6975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   8565
      Left            =   0
      Top             =   0
      Width           =   8610
   End
End
Attribute VB_Name = "FrmCargosPersonal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede

Private Sub Command1_Click()
Me.FrameDetalleFuncion.Visible = False
End Sub

Private Sub cmdprocesar_Click()
Dim id_permiso As String
Me.FrameDetalleFuncion.Visible = False
If Val(Me.txtcodigo.Text) < 1 Then
    'Ingresar
    
            strCadena = "SELECT * FROM funciones_empresa_detalle WHERE ruc='" & KEY_RUC & "' ORDER BY id_permiso DESC LIMIT 0,1"
            Call ConfiguraRstT(strCadena)
            If rstT.RecordCount > 0 Then
                id_permiso = formato_item(Val(rstT("id_permiso")) + 1, 5)
            Else
                id_permiso = "00001"
            End If
            
            'Generar funciones
            strCadena = "SELECT  * FROM persona_cargos WHERE id_empresa='" & KEY_RUC & "'"
            Call ConfiguraRst(strCadena)
            If rst.RecordCount > 0 Then
                rst.MoveFirst
                For i = 0 To rst.RecordCount - 1
                    strCadena = "INSERT INTO funciones_empresa_detalle(id_permiso,id_funcion,id_cargo,descripcion,ruc)VALUES('" & id_permiso & "','" & Me.HfActividades.TextMatrix(Me.HfActividades.Row, 0) & "','" & rst("id_cargo") & "','" & UCase(Trim(Me.TxtDescripcion.Text)) & "','" & KEY_RUC & "')"
                    CnBd.Execute (strCadena)
                     
                    rst.MoveNext
                Next i
            End If
    
Else
  
  strCadena = "SELECT * FROM funciones_empresa_detalle WHERE id_detalle='" & Val(Me.txtcodigo.Text) & "' AND ruc='" & KEY_RUC & "'"
  Call ConfiguraRst(strCadena)
  If rst.RecordCount > 0 Then
        strCadena = "UPDATE funciones_empresa_detalle SET descripcion='" & UCase(Trim(Me.TxtDescripcion.Text)) & "' WHERE id_permiso='" & rst("id_permiso") & "' AND  ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
         
  End If
  
  
End If

Call llenar_detalle_funcion(Me.HfFunciones, Me.HfgLinea.TextMatrix(Me.HfgLinea.Row, 0), Me.HfActividades.TextMatrix(Me.HfActividades.Row, 0))
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Me.FrameFunciones.Visible = False
    Me.FrameDetalleFuncion.Visible = False
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 100
Call actualizar
End Sub

Private Sub llenar_funciones(ByVal Grilla As MSHFlexGrid, ByVal id_cargo As String)
On Error GoTo salir
Dim color As String
Grilla.SelectionMode = flexSelectionFree
Grilla.MergeCells = flexMergeFree




Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub
End If
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
   ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 800
           Grilla.ColWidth(1) = 3500
           Grilla.ColWidth(2) = 1000
           
          Next
         cabecera = "CODIGO" & vbTab & "DESCRIPCION" & vbTab & "ESTADO"
         Grilla.AddItem cabecera
         For k = 0 To 2
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
               c = 2
               NumeroCampo = 2
            If rst("estado") = "no" Then
                estado = Chr(168)
            Else
                estado = Chr(254)
            End If
            Fila = rst("id_funcion") & vbTab & UCase(rst("descripcion")) & vbTab & estado
            Grilla.AddItem Fila
            If c = NumeroCampo Then
                        With Grilla
                            .Row = i + 1 ' se posiciona en la fila
                            .col = c '  .. en la columna
                            .CellFontName = "Wingdings"
                            .CellFontSize = 14
                            .CellAlignment = flexAlignCenterCenter
                            ' edita la celda
                             If rst("estado") = "no" Then
                                estado = Chr(168)
                            Else
                                estado = Chr(254)
                            End If
                        End With
         End If
         Fila = ""
          If rst("estado") = "si" Then
               color = &HC0FFC0
          Else
               color = &HC0C0FF
          End If
            For j = 0 To 2
                Grilla.col = j
                Grilla.Row = i + 1
                Grilla.CellBackColor = color
            Next j
            Fila = ""
        rst.MoveNext
        Next i
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub
Private Sub llenar_detalle_funcion(ByVal Grilla As MSHFlexGrid, ByVal id_cargo As String, ByVal id_funcion As String)
On Error GoTo salir
Dim color As String
Grilla.SelectionMode = flexSelectionFree
Grilla.MergeCells = flexMergeFree
strCadena = "SELECT id_detalle,descripcion,estado FROM funciones_empresa_detalle WHERE id_funcion='" & id_funcion & "' AND id_cargo='" & id_cargo & "'   AND ruc='" & KEY_RUC & "' ORDER BY id_detalle"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub
End If
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
   ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 3500
           Grilla.ColWidth(2) = 800
           
          Next
         cabecera = "CODIGO" & vbTab & "DESCRIPCION" & vbTab & "ESTADO"
         Grilla.AddItem cabecera
         For k = 0 To 2
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
         Next k
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
               c = 2
               NumeroCampo = 2
            If rst("estado") = "no" Then
                estado = Chr(168)
            Else
                estado = Chr(254)
            End If
            Fila = rst("id_detalle") & vbTab & UCase(rst("descripcion")) & vbTab & estado
            Grilla.AddItem Fila
            If c = NumeroCampo Then
                        With Grilla
                            .Row = i + 1 ' se posiciona en la fila
                            .col = c '  .. en la columna
                            .CellFontName = "Wingdings"
                            .CellFontSize = 14
                            .CellAlignment = flexAlignCenterCenter
                            ' edita la celda
                             If rst("estado") = "no" Then
                                estado = Chr(168)
                            Else
                                estado = Chr(254)
                            End If
                        End With
         End If
         Fila = ""
          If rst("estado") = "si" Then
               color = &HC0FFC0
          Else
               color = &HC0C0FF
          End If
            For j = 0 To 2
                Grilla.col = j
                Grilla.Row = i + 1
                Grilla.CellBackColor = color
            Next j
            Fila = ""
        rst.MoveNext
        Next i
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub
Private Sub HfActividades_DblClick()
If Me.FrameDetalle.Visible = True Then
    Me.FrameDetalle.Visible = False
Else
    Me.FrameDetalle.Visible = True
End If
End Sub

Private Sub HfActividades_Click()
If Me.HfActividades.col = 2 Then
    Call ActualizarCampo(Me.HfgLinea.TextMatrix(Me.HfgLinea.Row, 0), Me.HfActividades.TextMatrix(Me.HfActividades.Row, 0))
End If
End Sub
Private Sub ActualizarCampo(ByVal id_cargo As String, ByVal id_funcion As String)
     Dim estado As String
      strCadena = "SELECT * FROM cargo_funcion WHERE id_cargo='" & id_cargo & "' AND id_funcion='" & id_funcion & "' AND ruc='" & KEY_RUC & "'"
      Call ConfiguraRst(strCadena)
      If rst.RecordCount > 0 Then
        If rst("estado") = "si" Then
            estado = "no"
            Me.HfActividades.TextMatrix(Me.HfActividades.Row, 2) = Chr(168)
            For j = 0 To 2
                HfActividades.col = j
                HfActividades.Row = Me.HfActividades.Row
                HfActividades.CellBackColor = &HC0C0FF
            Next j
        Else
            estado = "si"
            Me.HfActividades.TextMatrix(Me.HfActividades.Row, 2) = Chr(254)
            For j = 0 To 2
                HfActividades.col = j
                HfActividades.Row = Me.HfActividades.Row
                HfActividades.CellBackColor = &HC0FFC0
            Next j
        End If
        
        strCadena = "UPDATE cargo_funcion SET estado='" & estado & "' WHERE id_cargo='" & id_cargo & "' AND id_funcion='" & id_funcion & "' AND ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
         
        
        If estado = "si" Then
            Me.lblfunciones.Caption = Me.HfActividades.TextMatrix(Me.HfActividades.Row, 1)
            Me.txtIdFuncion.Text = Me.HfActividades.TextMatrix(Me.HfActividades.Row, 0)
            Me.FrameDetalle.Visible = True
            Me.FrameDetalleFuncion.Visible = False
            Call llenar_detalle_funcion(Me.HfFunciones, Me.HfgLinea.TextMatrix(Me.HfgLinea.Row, 0), Me.HfActividades.TextMatrix(Me.HfActividades.Row, 0))
        Else
            Me.FrameDetalle.Visible = False
        End If
        
      End If
End Sub
Private Sub actualizar_funcion(ByVal id_detalle As Double)
     Dim estado As String
      strCadena = "SELECT * FROM funciones_empresa_detalle WHERE id_detalle='" & id_detalle & "' AND ruc='" & KEY_RUC & "'"
      Call ConfiguraRst(strCadena)
      If rst.RecordCount > 0 Then
        If rst("estado") = "si" Then
            estado = "no"
            Me.HfFunciones.TextMatrix(Me.HfFunciones.Row, 2) = Chr(168)
            For j = 0 To 2
                Me.HfFunciones.col = j
                HfFunciones.Row = Me.HfFunciones.Row
                HfFunciones.CellBackColor = &HC0C0FF
            Next j
        Else
            estado = "si"
            Me.HfFunciones.TextMatrix(Me.HfFunciones.Row, 2) = Chr(254)
            For j = 0 To 2
                HfFunciones.col = j
                HfFunciones.Row = Me.HfFunciones.Row
                HfFunciones.CellBackColor = &HC0FFC0
            Next j
        End If
        
        strCadena = "UPDATE funciones_empresa_detalle SET estado='" & estado & "' WHERE id_detalle='" & id_detalle & "' AND ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
         
        
        
        
      End If
End Sub


Private Sub HfFunciones_Click()
If Val(Me.HfFunciones.TextMatrix(Me.HfFunciones.Row, 0)) > 0 Then
    Call actualizar_funcion(Me.HfFunciones.TextMatrix(Me.HfFunciones.Row, 0))
End If
End Sub

Private Sub Text1_Change()

End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As ComctlLib.Button)
Select Case Button.key
    Case KEY_NEW
      Procedencia = nuevo
      FrmDetalleCargo.Show
    Case KEY_UPDATE
      Procedencia = modificar
      FrmDetalleCargo.Show
    
    Case "(Funciones)"
        Me.FrameFunciones.Visible = True
        'Call verificar_funcion(Me.HfgLinea.TextMatrix(Me.HfgLinea.Row, 0))
        Me.lblcargo.Caption = Me.HfgLinea.TextMatrix(Me.HfgLinea.Row, 1)
        strCadena = "SELECT F.id_funcion,E.descripcion,F.estado FROM cargo_funcion F,funciones_empresa E WHERE F.id_funcion=E.id_funcion AND F.ruc='" & KEY_RUC & "' AND F.id_cargo='" & Me.HfgLinea.TextMatrix(Me.HfgLinea.Row, 0) & "' AND  E.ruc='" & KEY_RUC & "' ORDER BY F.id_funcion"
        Call llenar_funciones(HfActividades, Me.HfgLinea.TextMatrix(Me.HfgLinea.Row, 0))
        
    Case KEY_DELETE
      Procedencia = Eliminar
      FrmSeguridad.Show
      Exit Sub
    Case KEY_EXIT
        Unload Me
  End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
Select Case Button.key
    Case KEY_NEW
           Me.FrameDetalleFuncion.Visible = True
           Me.txtcodigo.Text = 0
           Me.TxtDescripcion.Text = ""
           Call Resalta(Me.TxtDescripcion)
           Exit Sub
    Case KEY_UPDATE
         Me.FrameDetalleFuncion.Visible = True
         Me.txtcodigo.Text = Val(Me.HfFunciones.TextMatrix(Me.HfFunciones.Row, 0))
         Me.TxtDescripcion.Text = Me.HfFunciones.TextMatrix(Me.HfFunciones.Row, 1)
    Case KEY_EXIT
         Me.FrameDetalle.Visible = False
  End Select
  End Sub

Private Sub txtFuncion_Change()
'strCadena = "SELECT * FROM funciones_empresa WHERE ruc='" & KEY_RUC & "' AND id_cargo='" & Me.HfgLinea.TextMatrix(Me.HfgLinea.Row, 0) & "' AND descripcion LIKE '%" & Trim(Me.txtFuncion.text) & "%'"
strCadena = "SELECT F.id_funcion,E.descripcion,F.estado FROM cargo_funcion F,funciones_empresa E WHERE F.id_funcion=E.id_funcion AND F.ruc='" & KEY_RUC & "' AND F.id_cargo='" & Me.HfgLinea.TextMatrix(Me.HfgLinea.Row, 0) & "' AND E.descripcion LIKE '%" & Trim(Me.txtFuncion.Text) & "%' AND E.ruc='" & KEY_RUC & "' ORDER BY F.id_funcion"
Call llenar_funciones(Me.HfActividades, Me.HfgLinea.TextMatrix(Me.HfgLinea.Row, 0))
End Sub

Private Sub TxtLinea_Change()
strCadena = "SELECT id_cargo as Codigo,descripcion as Descripcion FROM persona_cargos WHERE descripcion LIKE '%" & Trim(Me.txtlinea.Text) & "%' AND id_empresa='" & KEY_RUC & "' ORDER BY descripcion ASC"
Call llenarGrid(Me.HfgLinea)
End Sub
Private Sub HfgLinea_Click()
If HfgLinea.Row > 0 Then
    
    TlbAcciones.Buttons("(Modificar)").Enabled = True
    TlbAcciones.Buttons(KEY_DELETE).Enabled = True
    
  End If
End Sub
Public Sub verificar_funcion(ByVal id_cargo As String)
strCadena = "SELECT * FROM funciones_empresa WHERE ruc='" & KEY_RUC & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    rst.MoveFirst
    For i = 0 To rst.RecordCount - 1
        strCadena = "SELECT count(*) FROM cargo_funcion WHERE id_cargo='" & id_cargo & "' AND id_funcion='" & rst("id_funcion") & "' AND ruc='" & KEY_RUC & "'"
        Call ConfiguraRstT(strCadena)
        If rstT(0) < 1 Then
            strCadena = "INSERT INTO cargo_funcion(id_cargo,id_funcion,ruc)VALUES('" & id_cargo & "','" & rst("id_funcion") & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
            
        End If
        rst.MoveNext
    Next i
End If
End Sub
Public Sub actualizar()
strCadena = "SELECT id_cargo as Codigo,descripcion as Descripcion FROM persona_cargos WHERE ruc='si' AND id_empresa='" & KEY_RUC & "' ORDER BY descripcion"
Call llenarGrid(Me.HfgLinea)
End Sub
Public Sub llenarGrid(ByVal Grilla As MSHFlexGrid)
On Error GoTo salir
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
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 1200
           Grilla.ColWidth(1) = 4500
           
        Next
        cabecera = "CODIGO" & vbTab & "DESCRIPCION "
        Grilla.AddItem cabecera
         For k = 0 To 1
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            Fila = rst("Codigo") & vbTab & UCase(rst("Descripcion"))
            Grilla.AddItem Fila
            
            rst.MoveNext
    Next i
        
  
  Grilla.Row = 1
  Grilla.col = 0
  Grilla.ColSel = 1
  Grilla.RowSel = 1
  
  
  Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing
End Sub







