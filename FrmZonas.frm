VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Begin VB.Form FrmZonas 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   14445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   14445
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   3360
      Top             =   2880
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
            Picture         =   "FrmZonas.frx":0000
            Key             =   "(Nuevo)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmZonas.frx":0454
            Key             =   "(Modificar)"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmZonas.frx":0774
            Key             =   "(Eliminar)"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmZonas.frx":0BC8
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmZonas.frx":101C
            Key             =   "(Aceptar)"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmZonas.frx":133C
            Key             =   "(Cancelar)"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmZonas.frx":165C
            Key             =   "(Quitar)"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmZonas.frx":197C
            Key             =   "(Agregar)"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmZonas.frx":1C9C
            Key             =   "(Buscar)"
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgZona 
      Height          =   5655
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   13215
      _ExtentX        =   23310
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
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   5625
      Left            =   13485
      TabIndex        =   1
      Top             =   450
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   9922
      BandCount       =   1
      ForeColor       =   8388608
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      Orientation     =   1
      EmbossPicture   =   -1  'True
      _CBWidth        =   900
      _CBHeight       =   5625
      _Version        =   "6.7.9782"
      Caption1        =   "Acciones"
      Child1          =   "TlbAcciones"
      MinHeight1      =   840
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   810
         Left            =   30
         TabIndex        =   2
         Top             =   375
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   1429
         ButtonWidth     =   1561
         ButtonHeight    =   1429
         Style           =   1
         ImageList       =   "ImgIconos"
         DisabledImageList=   "ImgIconos"
         HotImageList    =   "ImgIconos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   8
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Urbanizac"
               Key             =   "(Urbanizacion)"
               Object.ToolTipText     =   "Urbanizacion"
               ImageKey        =   "(Modificar)"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Zonas"
               Key             =   "(Zona)"
               ImageKey        =   "(Modificar)"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Sector"
               Key             =   "(Sector)"
               Object.ToolTipText     =   "Eliminar"
               ImageKey        =   "(Aceptar)"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Salir"
               Key             =   "(Salir)"
               ImageKey        =   "(Salir)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin MSDataListLib.DataCombo DtcDepartamento 
      Height          =   330
      Left            =   1560
      TabIndex        =   5
      Top             =   6360
      Width           =   3135
      _ExtentX        =   5530
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
   Begin MSDataListLib.DataCombo DtcDistrito 
      Height          =   330
      Left            =   10215
      TabIndex        =   6
      Top             =   6360
      Width           =   3135
      _ExtentX        =   5530
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
   Begin MSDataListLib.DataCombo DtcProvincia 
      Height          =   330
      Left            =   6075
      TabIndex        =   8
      Top             =   6360
      Width           =   3135
      _ExtentX        =   5530
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
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PROVINCIA :"
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
      Left            =   5085
      TabIndex        =   9
      Top             =   6400
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DISTRITO :"
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
      Left            =   9420
      TabIndex        =   7
      Top             =   6400
      Width           =   705
   End
   Begin VB.Label LblFecha 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DEPARTAMENTO :"
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
      Left            =   270
      TabIndex        =   4
      Top             =   6400
      Width           =   1215
   End
   Begin VB.Label LblEmpresa 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ZONAS REGISTRADAS"
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
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   1455
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   6855
      Left            =   0
      Top             =   0
      Width           =   14445
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00DFDFE0&
      FillColor       =   &H00DFDFE0&
      FillStyle       =   0  'Solid
      Height          =   675
      Left            =   240
      Top             =   6000
      Width           =   13215
   End
End
Attribute VB_Name = "FrmZonas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede
Dim cargando As Boolean

Private Sub DtcDepartamento_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    strCadena = "SELECT id_provincia as Codigo,descripcion as Descripcion FROM provincia WHERE id_departamento='" & Me.DtcDepartamento.BoundText & "'   ORDER BY descripcion"
    Call ConfiguraRst(strCadena)
    Call LlenaDataCombo(Me.DtcProvincia)

    strCadena = "SELECT id_distrito as Codigo,descripcion as Descripcion FROM distrito WHERE id_provincia='" & Me.DtcProvincia.BoundText & "' ORDER BY descripcion"
    Call ConfiguraRst(strCadena)
    Call LlenaDataCombo(Me.DtcDistrito)
    
strCadena = "SELECT     departamentos.id_depa, provincia.id_provincia, distrito.id_distrito, urbanizacion.id_urbanizacion, zona.id_zona, " & _
" departamentos.descripcion AS DEPARTAMENTO, provincia.descripcion AS PROVINCIA, distrito.descripcion AS DISTRITO, urbanizacion.descripcion AS URBANIZACION," & _
" zona.descripcion_zona AS ZONA,zona.id_alm  FROM         departamentos INNER JOIN provincia ON departamentos.id_depa = provincia.id_departamento INNER JOIN " & _
" distrito ON provincia.id_provincia = distrito.id_provincia INNER JOIN urbanizacion ON distrito.id_distrito = urbanizacion.id_distrito INNER JOIN " & _
" zona ON urbanizacion.id_urbanizacion = zona.id_urbanizacion WHERE departamentos.id_depa='" & Me.DtcDepartamento.BoundText & "' ORDER BY departamentos.descripcion,provincia.descripcion,urbanizacion.descripcion,zona.descripcion_zona"
 Call llenarGridME(Me.HfgZona, Me)

End If
End Sub

Private Sub DtcDistrito_Change()
'If Me.DtcDistrito.BoundText <> "" Then
'    strCadena = "SELECT     departamentos.id_depa, provincia.id_provincia, distrito.id_distrito, urbanizacion.id_urbanizacion, zona.id_zona, " & _
'" departamentos.descripcion AS DEPARTAMENTO, provincia.descripcion AS PROVINCIA, distrito.descripcion AS DISTRITO, urbanizacion.descripcion AS URBANIZACION," & _
'" zona.descripcion_zona AS ZONA,zona.id_alm  FROM         departamentos INNER JOIN provincia ON departamentos.id_depa = provincia.id_departamento INNER JOIN " & _
'" distrito ON provincia.id_provincia = distrito.id_provincia INNER JOIN urbanizacion ON distrito.id_distrito = urbanizacion.id_distrito INNER JOIN " & _
'" zona ON urbanizacion.id_urbanizacion = zona.id_urbanizacion WHERE departamentos.id_depa='" & Me.DtcDepartamento.BoundText & "' AND provincia.id_provincia='" & Me.DtcProvincia.BoundText & "' AND distrito.id_distrito='" & Me.DtcDistrito.BoundText & "' ORDER BY departamentos.descripcion,provincia.descripcion,urbanizacion.descripcion,zona.descripcion_zona"
' Call llenarGridME(Me.HfgZona, Me)

'End If
End Sub



Private Sub DtcDistrito_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
      
        strCadena = "SELECT     departamentos.id_depa, provincia.id_provincia, distrito.id_distrito, urbanizacion.id_urbanizacion, zona.id_zona, " & _
" departamentos.descripcion AS DEPARTAMENTO, provincia.descripcion AS PROVINCIA, distrito.descripcion AS DISTRITO, urbanizacion.descripcion AS URBANIZACION," & _
" zona.descripcion_zona AS ZONA,zona.id_alm FROM         departamentos INNER JOIN provincia ON departamentos.id_depa = provincia.id_departamento INNER JOIN " & _
" distrito ON provincia.id_provincia = distrito.id_provincia INNER JOIN urbanizacion ON distrito.id_distrito = urbanizacion.id_distrito INNER JOIN " & _
" zona ON urbanizacion.id_urbanizacion = zona.id_urbanizacion WHERE departamentos.id_depa='" & Me.DtcDepartamento.BoundText & "' and provincia.id_provincia='" & Me.DtcProvincia.BoundText & "' and  distrito.id_distrito='" & Me.DtcDistrito.BoundText & "'   ORDER BY departamentos.descripcion,provincia.descripcion,urbanizacion.descripcion,zona.descripcion_zona"
 Call llenarGridME(Me.HfgZona, Me)
cargando = False
End If

End Sub

Private Sub DtcProvincia_Change()
'If Me.DtcProvincia.BoundText <> "" Then
'    strCadena = "SELECT     departamentos.id_depa, provincia.id_provincia, distrito.id_distrito, urbanizacion.id_urbanizacion, zona.id_zona, " & _
" departamentos.descripcion AS DEPARTAMENTO, provincia.descripcion AS PROVINCIA, distrito.descripcion AS DISTRITO, urbanizacion.descripcion AS URBANIZACION," & _
" zona.descripcion_zona AS ZONA,zona.id_alm  FROM departamentos INNER JOIN provincia ON departamentos.id_depa = provincia.id_departamento INNER JOIN " & _
" distrito ON provincia.id_provincia = distrito.id_provincia INNER JOIN urbanizacion ON distrito.id_distrito = urbanizacion.id_distrito INNER JOIN " & _
" zona ON urbanizacion.id_urbanizacion = zona.id_urbanizacion WHERE departamentos.id_depa='" & Me.DtcDepartamento.BoundText & "' AND provincia.id_provincia='" & Me.DtcProvincia.BoundText & "' ORDER BY departamentos.descripcion,provincia.descripcion,urbanizacion.descripcion,zona.descripcion_zona"
' Call llenarGridME(Me.HfgZona, Me)
'strCadena = "SELECT id_distrito as Codigo,descripcion as Descripcion FROM distrito WHERE id_provincia='" & Me.DtcProvincia.BoundText & "' ORDER BY descripcion"
'Call ConfiguraRst(strCadena)
'Call LlenaDataCombo(Me.DtcDistrito)
'End If
End Sub

Private Sub DtcProvincia_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        strCadena = "SELECT id_distrito as Codigo,descripcion as Descripcion FROM distrito WHERE id_provincia='" & Me.DtcProvincia.BoundText & "' ORDER BY descripcion"
        Call ConfiguraRst(strCadena)
        Call LlenaDataCombo(Me.DtcDistrito)
        
        strCadena = "SELECT     departamentos.id_depa, provincia.id_provincia, distrito.id_distrito, urbanizacion.id_urbanizacion, zona.id_zona, " & _
" departamentos.descripcion AS DEPARTAMENTO, provincia.descripcion AS PROVINCIA, distrito.descripcion AS DISTRITO, urbanizacion.descripcion AS URBANIZACION," & _
" zona.descripcion_zona AS ZONA,zona.id_alm FROM         departamentos INNER JOIN provincia ON departamentos.id_depa = provincia.id_departamento INNER JOIN " & _
" distrito ON provincia.id_provincia = distrito.id_provincia INNER JOIN urbanizacion ON distrito.id_distrito = urbanizacion.id_distrito INNER JOIN " & _
" zona ON urbanizacion.id_urbanizacion = zona.id_urbanizacion WHERE departamentos.id_depa='" & Me.DtcDepartamento.BoundText & "' and provincia.id_provincia='" & Me.DtcProvincia.BoundText & "'    ORDER BY departamentos.descripcion,provincia.descripcion,urbanizacion.descripcion,zona.descripcion_zona"
 Call llenarGridME(Me.HfgZona, Me)
cargando = False
        
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 500
cargando = True

strCadena = "SELECT id_depa as Codigo,descripcion as Descripcion FROM departamentos ORDER BY descripcion"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcDepartamento)
Me.DtcDepartamento.BoundText = KEY_DEPARTAMENTO

strCadena = "SELECT id_provincia as Codigo,descripcion as Descripcion FROM provincia WHERE id_departamento='" & KEY_DEPARTAMENTO & "'   ORDER BY descripcion"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcProvincia)
Me.DtcProvincia.BoundText = KEY_PROVINCIA

strCadena = "SELECT id_distrito as Codigo,descripcion as Descripcion FROM distrito WHERE id_provincia='" & KEY_PROVINCIA & "' ORDER BY descripcion"
Call ConfiguraRst(strCadena)
Call LlenaDataCombo(Me.DtcDistrito)
Me.DtcDistrito.BoundText = KEY_DISTRITO

Call actualizar
End Sub
Public Sub actualizar()
strCadena = "SELECT     departamentos.id_depa, provincia.id_provincia, distrito.id_distrito, urbanizacion.id_urbanizacion, zona.id_zona, " & _
" departamentos.descripcion AS DEPARTAMENTO, provincia.descripcion AS PROVINCIA, distrito.descripcion AS DISTRITO, urbanizacion.descripcion AS URBANIZACION," & _
" zona.descripcion_zona AS ZONA,zona.id_alm FROM         departamentos INNER JOIN provincia ON departamentos.id_depa = provincia.id_departamento INNER JOIN " & _
" distrito ON provincia.id_provincia = distrito.id_provincia INNER JOIN urbanizacion ON distrito.id_distrito = urbanizacion.id_distrito INNER JOIN " & _
" zona ON urbanizacion.id_urbanizacion = zona.id_urbanizacion WHERE departamentos.id_depa='" & Me.DtcDepartamento.BoundText & "'    ORDER BY departamentos.descripcion,provincia.descripcion,urbanizacion.descripcion,zona.descripcion_zona"
 Call llenarGridME(Me.HfgZona, Me)
cargando = False
End Sub



Private Sub HfgZona_SelChange()
If Val(Me.HfgZona.TextMatrix(Me.HfgZona.Row, 0)) > 0 Then

 TlbAcciones.Buttons(KEY_ZONA).Enabled = True
 TlbAcciones.Buttons(KEY_URBANIZACION).Enabled = True
  TlbAcciones.Buttons(KEY_SECTOR).Enabled = True
End If
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.key
    Case KEY_URBANIZACION
       
      Procedencia = nuevo
      FrmDetalleUrbanizacion.Show
    Case Procedencia = nuevo
    
    Case KEY_ZONA
       Procedencia = nuevo
      FrmDetalleZonas.Show
    Case KEY_SECTOR
        FrmAsignarSector.Show
    Case KEY_EXIT
        Unload Me
  End Select
End Sub
Private Sub llenarGridME(ByVal Grilla As MSHFlexGrid, ByVal Formulario As Form)
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
            Grilla.ColWidth(0) = 0
            Grilla.ColWidth(1) = 0
            Grilla.ColWidth(2) = 0
            Grilla.ColWidth(3) = 0
            Grilla.ColWidth(4) = 0
            Grilla.ColWidth(5) = 1800
            Grilla.ColWidth(6) = 2500
            Grilla.ColWidth(7) = 2500
            Grilla.ColWidth(8) = 2500
            Grilla.ColWidth(9) = 2000
            Grilla.ColWidth(10) = 1500
          Next
         cabecera = "IDDEPARTAMENTO" & vbTab & "IDPROVINCIA" & vbTab & "IDDISTRITO" & vbTab & "IDURBANIZACION" & vbTab & "IDZONA" & vbTab & "DEPARTAMENTO" & vbTab & "PROVINCIA" & vbTab & "DISTRITO" & vbTab & "URBANIZACION" & vbTab & "ZONA" & vbTab & "DEPOSITO"
         Grilla.AddItem cabecera
         For k = 0 To 10
                                Grilla.col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 1 To rst.RecordCount
            deposito = ""
             If Val(rst("id_alm")) > 0 Then
                    strCadena = "SELECT  descripcion FROM almacen WHERE id_alm='" & rst("id_alm") & "' AND ruc='" & KEY_RUC & "'"
                    Call ConfiguraTemporal(strCadena)
                    If rstTemporal.RecordCount > 0 Then
                        deposito = rstTemporal("descripcion")
                    End If
                    Set rstTemporal = Nothing
             End If
             
             Fila = Fila & rst("id_depa") & vbTab & rst("id_provincia") & vbTab & rst("id_distrito") & vbTab & rst("id_urbanizacion") & vbTab & rst("id_zona") & vbTab & UCase(rst("DEPARTAMENTO")) & vbTab & UCase(rst("PROVINCIA")) & vbTab & UCase(rst("DISTRITO")) & vbTab & UCase(rst("URBANIZACION")) & vbTab & UCase(rst("ZONA")) & vbTab & deposito
            If (Fila = "") Then
                X = 1
            End If
            Grilla.AddItem Fila
            Fila = ""
            rst.MoveNext
        Next i
        Me.TlbAcciones.Buttons(KEY_URBANIZACION).Enabled = False
        Me.TlbAcciones.Buttons(KEY_ZONA).Enabled = False
        Me.TlbAcciones.Buttons(KEY_SECTOR).Enabled = False
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub





