VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmListadoUsuario 
   ClientHeight    =   4560
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4560
   ScaleWidth      =   6105
   Begin MSDataGridLib.DataGrid DtgAdministrar 
      Height          =   3615
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   6376
      _Version        =   393216
      AllowUpdate     =   0   'False
      ColumnHeaders   =   -1  'True
      HeadLines       =   1
      RowHeight       =   15
      RowDividerStyle =   6
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "IdUsuario"
         Caption         =   "CÓDIGO"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "NombreUsuario"
         Caption         =   "NOMBRE"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         BeginProperty Column00 
            DividerStyle    =   5
            Object.Visible         =   -1  'True
            ColumnWidth     =   854.929
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   3014.929
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   3210
      Left            =   5160
      TabIndex        =   1
      Top             =   240
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   5662
      BandCount       =   1
      Orientation     =   1
      _CBWidth        =   855
      _CBHeight       =   3210
      _Version        =   "6.0.8169"
      Child1          =   "TlbAcciones"
      MinHeight1      =   795
      Width1          =   1140
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   2850
         Left            =   30
         TabIndex        =   2
         Top             =   30
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   5027
         ButtonWidth     =   1032
         ButtonHeight    =   1005
         Appearance      =   1
         Style           =   1
         ImageList       =   "ImlAcciones"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "[Nuevo]"
               Object.ToolTipText     =   "Nuevo Elemento"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "[Modificar]"
               Object.ToolTipText     =   "Modificar Elemento"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "[Eliminar]"
               Object.ToolTipText     =   "Eliminar Elemento"
               ImageIndex      =   3
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ImlAcciones 
      Left            =   2760
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmListadoUsuario.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmListadoUsuario.frx":5C22
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmListadoUsuario.frx":64FC
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmListadoUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrstDataSource As Object
Private Const mstrFormName As String = "Listado de Usuarios"

Public Property Get rstDataSource() As Variant
    Set rstDataSource = mrstDataSource
End Property

Public Property Set rstDataSource(ByVal vNewValue As Variant)
    Set mrstDataSource = vNewValue
    Set Me.DtgAdministrar.DataSource = mrstDataSource
    Me.DtgAdministrar.ReBind
End Property

Public Sub LoadData()
Dim objNegocio As Object
On Error GoTo ErrHandler
    Set objNegocio = CreateObject("Negocio.clsSeguridad")
    Set Me.rstDataSource = objNegocio.Buscar()
    
    Set objNegocio = Nothing
Exit Sub
ErrHandler:
    Set objNegocio = Nothing
    ErrorMessage FrmListadoUsuario_LoadData, Err.Source & " FrmListadoUsuario:LoadData", Err.Description
End Sub

Public Sub ShowForm()
    LoadData
    Me.Show
End Sub

Private Sub Form_Load()
    CenterForm Me
    Me.Caption = mstrFormName
End Sub

Private Sub Form_Resize()
Dim intAuxTop As Integer
    On Error Resume Next
    intAuxTop = 100
    Me.DtgAdministrar.Move 100, intAuxTop, (Me.Width - Me.ClbAcciones.Width - 300), (Me.Height - intAuxTop - 540)
    Me.ClbAcciones.Move Me.DtgAdministrar.Width + 100, intAuxTop, Me.Width, (Me.Height - intAuxTop - 540)
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case KEY_ACCION_NUEVO
        Nuevo
    Case KEY_ACCION_MODIFICAR
        Modificar
    Case KEY_ACCION_ELIMINAR
        If MsgBox(MSG_CONFIRMACION, vbQuestion + vbYesNo, mstrFormName) = vbYes Then
            If Not Eliminar Then
                MsgBox MSG_OPERACION_FALLO, vbCritical, mstrFormName
            End If
        End If
End Select
End Sub

Private Function Nuevo() As Boolean
Dim objForm As FrmMantenimientoUsuario
On Error GoTo ErrHandler
    Set objForm = New FrmMantenimientoUsuario
    objForm.intIdUsuario = gintNullCodigo
    objForm.intAction = enumAcciones.Nuevo
    objForm.ShowForm
    Set objForm = Nothing
Exit Function
ErrHandler:
    Set objForm = Nothing
    ErrorMessage FrmListadoUsuario_Nuevo, Err.Source & " FrmListadoUsuario:Nuevo", Err.Description
End Function

Private Function Modificar() As Boolean
Dim objForm As FrmMantenimientoUsuario
On Error GoTo ErrHandler
    Set objForm = New FrmMantenimientoUsuario
    objForm.intIdUsuario = Me.rstDataSource!IdUsuario
    objForm.intAction = enumAcciones.Modificar
    objForm.ShowForm
    Set objForm = Nothing
Exit Function
ErrHandler:
    Set objForm = Nothing
    ErrorMessage FrmListadoUsuario_Modificar, Err.Source & " FrmListadoUsuario:Modificar", Err.Description
End Function

Private Function Eliminar() As Boolean
Dim objNegocio As Object
Dim intIdUsuario As Integer
On Error GoTo ErrHandler
    intIdUsuario = Me.rstDataSource!IdUsuario
    Set objNegocio = CreateObject("Negocio.clsSeguridad")
    Eliminar = objNegocio.Eliminar(intIdUsuario)
    If Eliminar Then
        LoadData
    End If
Exit Function
ErrHandler:
    Eliminar = False
    ErrorMessage FrmListadoUsuario_Eliminar, Err.Source & " FrmListadoUsuario:Eliminar", Err.Description
End Function


