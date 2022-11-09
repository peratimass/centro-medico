VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmAsignarSector 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7815
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   Begin VB.OptionButton OptZona 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Seleccionar"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6240
      TabIndex        =   12
      Top             =   2400
      Width           =   1335
   End
   Begin VB.OptionButton OptUrbanizacion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Seleccionar"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6240
      TabIndex        =   11
      Top             =   1800
      Width           =   1335
   End
   Begin VB.OptionButton OptDistrito 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Seleccionar"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6240
      TabIndex        =   10
      Top             =   1320
      Width           =   1335
   End
   Begin VB.OptionButton OptProvincia 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Seleccionar"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6240
      TabIndex        =   9
      Top             =   720
      Width           =   1335
   End
   Begin VB.OptionButton OptDepartamento 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Seleccionar"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   6240
      TabIndex        =   7
      Top             =   240
      Width           =   1335
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   240
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsignarSector.frx":0000
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsignarSector.frx":031C
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsignarSector.frx":077C
            Key             =   "(Inicio)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsignarSector.frx":0BDC
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsignarSector.frx":0EF8
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsignarSector.frx":1358
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsignarSector.frx":1674
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsignarSector.frx":1AD4
            Key             =   "(Red)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsignarSector.frx":1F34
            Key             =   "(Grabar)"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsignarSector.frx":2814
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsignarSector.frx":2B30
            Key             =   "(Buscar)"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsignarSector.frx":2E4C
            Key             =   "(Cancelar)"
         EndProperty
      EndProperty
   End
   Begin MSDataListLib.DataCombo DtcProvincia 
      Height          =   315
      Left            =   1920
      TabIndex        =   0
      Top             =   600
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Locked          =   -1  'True
      Appearance      =   0
      Style           =   2
      ForeColor       =   8388608
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DtcDistrito 
      Height          =   315
      Left            =   1920
      TabIndex        =   1
      Top             =   1200
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Locked          =   -1  'True
      Appearance      =   0
      Style           =   2
      ForeColor       =   8388608
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DtcDepartamento 
      Height          =   315
      Left            =   1920
      TabIndex        =   2
      Top             =   120
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Locked          =   -1  'True
      Appearance      =   0
      Style           =   2
      ForeColor       =   8388608
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DtcUrbanizacion 
      Height          =   315
      Left            =   1920
      TabIndex        =   3
      Top             =   1800
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Locked          =   -1  'True
      Appearance      =   0
      Style           =   2
      ForeColor       =   8388608
      Text            =   "DataCombo1"
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   870
      Left            =   6060
      TabIndex        =   5
      Top             =   2850
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   1535
      BandCount       =   1
      ForeColor       =   -2147483635
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      EmbossPicture   =   -1  'True
      _CBWidth        =   1515
      _CBHeight       =   870
      _Version        =   "6.0.8169"
      Child1          =   "TlbAcciones"
      MinHeight1      =   810
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   810
         Left            =   30
         TabIndex        =   6
         Top             =   30
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   1429
         ButtonWidth     =   1111
         ButtonHeight    =   1429
         Style           =   1
         ImageList       =   "ImgIconos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Grabar"
               Key             =   "(Grabar)"
               Object.ToolTipText     =   "Cancelar"
               ImageKey        =   "(Grabar)"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Salir ."
               Key             =   "(Salir)"
               ImageKey        =   "(Salir)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin MSDataListLib.DataCombo DtcZona 
      Height          =   315
      Left            =   1920
      TabIndex        =   8
      Top             =   2400
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Locked          =   -1  'True
      Appearance      =   0
      Style           =   2
      ForeColor       =   8388608
      Text            =   "DataCombo1"
   End
   Begin MSDataListLib.DataCombo DtcAlmacen 
      Height          =   315
      Left            =   1920
      TabIndex        =   18
      Top             =   3120
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   8388608
      Text            =   "DataCombo1"
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DEPOSITO :"
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
      Left            =   585
      TabIndex        =   17
      Top             =   3120
      Width           =   945
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderStyle     =   3  'Dot
      Height          =   735
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   2880
      Width           =   5775
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ZONA :"
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
      Left            =   1125
      TabIndex        =   16
      Top             =   2520
      Width           =   555
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "URBANIZACION :"
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
      Left            =   285
      TabIndex        =   15
      Top             =   1920
      Width           =   1395
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DISTRITO :"
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
      Left            =   735
      TabIndex        =   14
      Top             =   1320
      Width           =   945
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PROVINCIA :"
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
      Left            =   645
      TabIndex        =   13
      Top             =   720
      Width           =   1035
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   3855
      Left            =   0
      Top             =   0
      Width           =   7815
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DEPARTAMENTO"
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
      Left            =   315
      TabIndex        =   4
      Top             =   240
      Width           =   1365
   End
End
Attribute VB_Name = "FrmAsignarSector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 1800
strCadena = "SELECT id_depa as Codigo,descripcion as Descripcion FROM departamentos ORDER BY descripcion"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcDepartamento)
Me.DtcDepartamento.BoundText = FrmZonas.HfgZona.TextMatrix(FrmZonas.HfgZona.Row, 0)
Set rst = Nothing

strCadena = "SELECT id_provincia as Codigo,descripcion as  Descripcion FROM provincia WHERE id_departamento='" & Me.DtcDepartamento.BoundText & "' ORDER BY descripcion"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcProvincia)
Me.DtcProvincia.BoundText = FrmZonas.HfgZona.TextMatrix(FrmZonas.HfgZona.Row, 1)
Set rst = Nothing

strCadena = "SELECT id_distrito as Codigo,descripcion as  Descripcion FROM distrito WHERE id_provincia='" & Me.DtcProvincia.BoundText & "' ORDER BY descripcion"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcDistrito)
Me.DtcDistrito.BoundText = FrmZonas.HfgZona.TextMatrix(FrmZonas.HfgZona.Row, 2)
Set rst = Nothing


strCadena = "SELECT id_urbanizacion as Codigo,descripcion as  Descripcion FROM urbanizacion WHERE  id_distrito='" & Me.DtcDistrito.BoundText & "' ORDER BY descripcion"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcUrbanizacion)
Me.DtcUrbanizacion.BoundText = FrmZonas.HfgZona.TextMatrix(FrmZonas.HfgZona.Row, 3)
Set rst = Nothing

strCadena = "SELECT id_zona as Codigo,descripcion_zona as  Descripcion FROM zona WHERE  id_urbanizacion='" & Me.DtcUrbanizacion.BoundText & "' ORDER BY descripcion_zona"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcZona)
Me.DtcZona.BoundText = FrmZonas.HfgZona.TextMatrix(FrmZonas.HfgZona.Row, 4)
Set rst = Nothing

strCadena = "SELECT id_alm as Codigo, descripcion as Descripcion FROM almacen where ruc='" & KEY_RUC & "' ORDER BY descripcion"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcAlmacen)
Me.DtcDepartamento.Enabled = False
Me.DtcProvincia.Enabled = False
Me.DtcDistrito.Enabled = False
Me.DtcDistrito.Enabled = False
Me.DtcUrbanizacion.Enabled = False
Me.DtcZona.Enabled = False

End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo error
  Select Case Button.key
    Case KEY_SAVE
      If Me.DtcAlmacen.BoundText <> "" Then
      If Me.OptDepartamento.Value = True Then
            strCadena = "UPDATE departamento SET id_alm='" & Me.DtcAlmacen.BoundText & "' WHERE id_departamento='" & Me.DtcDepartamento.BoundText & "'"
            CnBd.Execute (strCadena)
             
            strCadena = "UPDATE provincia SET id_alm='" & Me.DtcAlmacen.BoundText & "' WHERE id_departamento='" & Me.DtcDepartamento.BoundText & "'"
            CnBd.Execute (strCadena)
             
            strCadena = "UPDATE distrito SET id_alm='" & Me.DtcAlmacen.BoundText & "' WHERE id_provincia='" & Me.DtcProvincia.BoundText & "'"
            CnBd.Execute (strCadena)
             
            strCadena = "UPDATE Ubanizacion SET id_alm='" & Me.DtcAlmacen.BoundText & "' WHERE id_distrito='" & Me.DtcDistrito.BoundText & "'"
            CnBd.Execute (strCadena)
             
            strCadena = "UPDATE Zona SET id_alm='" & Me.DtcAlmacen.BoundText & "' WHERE id_urbanizacion='" & Me.DtcUrbanizacion.BoundText & "'"
            CnBd.Execute (strCadena)
             
            Call FrmZonas.actualizar
            Unload Me
            Exit Sub
            
      End If
      If Me.OptProvincia.Value = True Then
            strCadena = "UPDATE provincia SET id_alm='" & Me.DtcAlmacen.BoundText & "' WHERE id_provincia='" & Me.DtcProvincia.BoundText & "'"
            CnBd.Execute (strCadena)
             
            strCadena = "UPDATE distrito SET id_alm='" & Me.DtcAlmacen.BoundText & "' WHERE id_provincia='" & Me.DtcProvincia.BoundText & "'"
            CnBd.Execute (strCadena)
             
            strCadena = "UPDATE Ubanizacion SET id_alm='" & Me.DtcAlmacen.BoundText & "' WHERE id_distrito='" & Me.DtcDistrito.BoundText & "'"
            CnBd.Execute (strCadena)
             
            strCadena = "UPDATE Zona SET id_alm='" & Me.DtcAlmacen.BoundText & "' WHERE id_urbanizacion='" & Me.DtcUrbanizacion.BoundText & "'"
            CnBd.Execute (strCadena)
             
            Call FrmZonas.actualizar
            Unload Me
            Exit Sub
            
      End If
      If Me.OptDistrito.Value = True Then
            strCadena = "UPDATE distrito SET id_alm='" & Me.DtcAlmacen.BoundText & "' WHERE id_distrito='" & Me.DtcDistrito.BoundText & "'"
            CnBd.Execute (strCadena)
             
            strCadena = "UPDATE Ubanizacion SET id_alm='" & Me.DtcAlmacen.BoundText & "' WHERE id_distrito='" & Me.DtcDistrito.BoundText & "'"
            CnBd.Execute (strCadena)
             
            strCadena = "UPDATE Zona SET id_alm='" & Me.DtcAlmacen.BoundText & "' WHERE id_urbanizacion='" & Me.DtcUrbanizacion.BoundText & "'"
            CnBd.Execute (strCadena)
             
            Call FrmZonas.actualizar
            Unload Me
            Exit Sub
            
      End If
      If Me.OptUrbanizacion.Value = True Then
            strCadena = "UPDATE Ubanizacion SET id_alm='" & Me.DtcAlmacen.BoundText & "' WHERE id_Urbanizacion='" & Me.DtcUrbanizacion.BoundText & "'"
            CnBd.Execute (strCadena)
             
            strCadena = "UPDATE Zona SET id_alm='" & Me.DtcAlmacen.BoundText & "' WHERE id_Urbanizacion='" & Me.DtcUrbanizacion.BoundText & "'"
            CnBd.Execute (strCadena)
             
            Call FrmZonas.actualizar
            Unload Me
            Exit Sub
            
      End If
        If Me.OptZona.Value = True Then
            strCadena = "UPDATE Zona SET id_alm='" & Me.DtcAlmacen.BoundText & "' WHERE id_zona='" & Me.DtcZona.BoundText & "'"
            CnBd.Execute (strCadena)
             
            Call FrmZonas.actualizar
            Unload Me
            Exit Sub
            
      End If
      End If
    Case KEY_EXIT
        Unload Me
  End Select
  Exit Sub
error:
  MsgBox Err.Number & " " & Err.Description, vbCritical, MSGERROR
  Exit Sub

End Sub
