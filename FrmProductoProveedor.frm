VERSION 5.00
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmProductoProveedor 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9660
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9660
   ScaleWidth      =   15615
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      Height          =   375
      Left            =   6840
      Picture         =   "FrmProductoProveedor.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Sin Proveedor"
      Top             =   4560
      Width           =   495
   End
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   6240
      Picture         =   "FrmProductoProveedor.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Mostrar Todos"
      Top             =   4560
      Width           =   495
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7800
      Top             =   8400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProductoProveedor.frx":0914
            Key             =   "(Delete)"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CmdMostrar 
      Height          =   375
      Left            =   14760
      Picture         =   "FrmProductoProveedor.frx":0D66
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4560
      Width           =   495
   End
   Begin VB.TextBox TxtRuc 
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
      Left            =   12720
      TabIndex        =   8
      Top             =   4560
      Width           =   1935
   End
   Begin VB.TextBox TxtApellido 
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
      Left            =   9840
      TabIndex        =   7
      Top             =   4560
      Width           =   1935
   End
   Begin VB.TextBox TxtProducto 
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
      Left            =   4080
      TabIndex        =   1
      Top             =   4560
      Width           =   2055
   End
   Begin VB.TextBox TxtCod 
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
      Left            =   1140
      TabIndex        =   0
      Top             =   4560
      Width           =   1575
   End
   Begin MSComctlLib.ImageList ImgIconos 
      Left            =   4560
      Top             =   3480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProductoProveedor.frx":0EB0
            Key             =   "(Salir)"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmProductoProveedor.frx":1304
            Key             =   "(Asignar)"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar ClbAcciones 
      Height          =   4065
      Left            =   7560
      TabIndex        =   9
      Top             =   240
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   7170
      BandCount       =   1
      ForeColor       =   8388608
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      Orientation     =   1
      EmbossPicture   =   -1  'True
      _CBWidth        =   900
      _CBHeight       =   4065
      _Version        =   "6.0.8169"
      Caption1        =   "Acciones"
      Child1          =   "TlbAcciones"
      MinHeight1      =   840
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar TlbAcciones 
         Height          =   780
         Left            =   30
         TabIndex        =   10
         Top             =   420
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   1376
         ButtonWidth     =   1455
         ButtonHeight    =   1376
         Style           =   1
         ImageList       =   "ImgIconos"
         DisabledImageList=   "ImgIconos"
         HotImageList    =   "ImgIconos"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "    Asignar"
               Key             =   "(Asignar)"
               Object.ToolTipText     =   "Modificar"
               ImageKey        =   "(Asignar)"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Salir"
               Key             =   "(Salir)"
               ImageKey        =   "(Salir)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   2145
      Left            =   14880
      TabIndex        =   12
      Top             =   5400
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   3784
      BandCount       =   1
      ForeColor       =   8388608
      ImageList       =   "ImgIconos"
      FixedOrder      =   -1  'True
      VariantHeight   =   0   'False
      Orientation     =   1
      EmbossPicture   =   -1  'True
      _CBWidth        =   615
      _CBHeight       =   2145
      _Version        =   "6.0.8169"
      Child1          =   "TlbAcciones"
      MinHeight1      =   555
      Width1          =   3180
      FixedBackground1=   0   'False
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   780
         Left            =   30
         TabIndex        =   13
         Top             =   540
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   1376
         ButtonWidth     =   1032
         ButtonHeight    =   1376
         Style           =   1
         ImageList       =   "ImageList1"
         DisabledImageList=   "ImageList1"
         HotImageList    =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   1
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Quitar"
               Key             =   "(Eliminar)"
               Object.ToolTipText     =   "Modificar"
               ImageKey        =   "(Delete)"
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdGrilla 
      Height          =   4095
      Left            =   120
      TabIndex        =   17
      Top             =   240
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   7223
      _Version        =   393216
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdPersona 
      Height          =   4095
      Left            =   8520
      TabIndex        =   18
      Top             =   240
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   7223
      _Version        =   393216
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdDetalle 
      Height          =   4215
      Left            =   120
      TabIndex        =   19
      Top             =   5160
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   7435
      _Version        =   393216
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfProveedorProducto 
      Height          =   4215
      Left            =   8520
      TabIndex        =   20
      Top             =   5160
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   7435
      _Version        =   393216
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
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
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PROVEEDORES"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   8505
      TabIndex        =   16
      Top             =   0
      Width           =   1305
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RUC/DNI :"
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
      Height          =   195
      Left            =   11880
      TabIndex        =   6
      Top             =   4560
      Width           =   765
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RAZON SOCIAL:"
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
      Height          =   195
      Left            =   8610
      TabIndex        =   5
      Top             =   4560
      Width           =   1215
   End
   Begin VB.Label LblFecha 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUCTO:"
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
      Height          =   195
      Left            =   2940
      TabIndex        =   4
      Top             =   4590
      Width           =   915
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO:"
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
      Height          =   195
      Left            =   345
      TabIndex        =   3
      Top             =   4590
      Width           =   705
   End
   Begin VB.Label LblEmpresa 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUCTOS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   345
      TabIndex        =   2
      Top             =   0
      Width           =   1125
   End
   Begin VB.Shape ShpDatos 
      BackColor       =   &H00DFDFE0&
      BorderStyle     =   3  'Dot
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   555
      Left            =   240
      Top             =   4440
      Width           =   7215
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   3  'Dot
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   555
      Left            =   8520
      Top             =   4440
      Width           =   6855
   End
End
Attribute VB_Name = "FrmProductoProveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Procedencia As EnumProcede

Private Sub CmdMostrar_Click()
Me.Height = 9500

If Me.HfProveedorProducto.Visible = True Then
    Me.HfProveedorProducto.Visible = False
    Me.HfdDetalle.Visible = False
    Me.Height = 5190
Else
    Me.HfProveedorProducto.Visible = True
    Me.HfdDetalle.Visible = True
    Me.Height = 9500
End If


End Sub

Private Sub Command1_Click()
Call mostrarTodos
End Sub

Private Sub Command2_Click()
Call ActualizarProd
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 50
Call ActualizarProd
Call ActualizarPer

End Sub
Sub ActualizarProd()

 If Procedencia = Selecionar Then
    strCadena = "SELECT id_producto, nombre_prod FROM producto WHERE id_proveedor='' AND nombre_prod LIKE '%" & Trim(Me.txtProducto.Text) & "%' AND ruc='" & KEY_RUC & "' ORDER BY id_producto LIMIT 0,20"
    
 Else
    strCadena = "SELECT id_producto, nombre_prod FROM producto P WHERE id_proveedor='0' AND ruc='" & KEY_RUC & "' AND ruc='" & KEY_RUC & "' ORDER BY P.nombre_prod LIMIT 0,15"
 End If
 Call llenarGrid_prod(Me.HfdGrilla, Me)
End Sub
Sub mostrarTodos()
strCadena = "SELECT id_producto, nombre_prod  FROM producto P WHERE P.ruc='" & KEY_RUC & "' ORDER BY P.nombre_prod LIMIT 0,15"
Call llenarGrid_prod(Me.HfdGrilla, Me)
End Sub
Sub ActualizarPer()

  strCadena = "SELECT dni,nombre_completo FROM persona P,entidad_empresa E WHERE P.dni=E.cod_unico AND E.id_empresa='" & KEY_RUC & "' AND  id_proveedor='si' ORDER BY P.nombre_completo LIMIT 0,20"
  Call llenarGrid_per(Me.HfdPersona, Me)
End Sub
Sub llenarGrid_prod(ByVal Grilla As MSHFlexGrid, ByVal Formulario As Form)
'On Error GoTo salir
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
     Me.TlbAcciones.Buttons(KEY_ASIGNAR).Enabled = False
    Exit Sub
Else
     Me.TlbAcciones.Buttons(KEY_ASIGNAR).Enabled = True
End If
   Grilla.Clear
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 700
           Grilla.ColWidth(1) = 5200
    Next
        cabecera = "CODIGO" & vbTab & "DESCRIPCION ARTICULO"
        Grilla.AddItem cabecera
         For k = 0 To 1
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            Fila = Fila & rst("id_producto") & vbTab & rst("nombre_prod")
            If (Fila = "") Then
                x = 1
            End If
          Grilla.AddItem Fila
            Fila = ""
            rst.MoveNext
             
        Next i
        
  
  Grilla.Row = 1
  Grilla.col = 0
  Grilla.ColSel = 1
 ' Grilla.RowSel = 1
  
  
 ' Exit Sub
'salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing

End Sub
Sub llenarGrid_per(ByVal Grilla As MSHFlexGrid, ByVal Formulario As Form)
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
           Grilla.ColWidth(1) = 4600
           
    Next
        cabecera = "RUC/DNI" & vbTab & "RAZON SOCIAL"
        Grilla.AddItem cabecera
         For k = 0 To 1
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            Fila = Fila & rst("dni") & vbTab & rst("nombre_completo")
            If (Fila = "") Then
                x = 1
            End If
          Grilla.AddItem Fila
            Fila = ""
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
Private Sub HfdGrilla_Click()
If Me.HfdGrilla.Rows > 1 Then
    Me.TlbAcciones.Buttons(KEY_ASIGNAR).Enabled = True
End If
End Sub

Private Sub HfdGrilla_SelChange()
If Me.HfdDetalle.Visible = True Then
    strCadena = "SELECT dni,nombre_completo FROM persona P,entidad_empresa E,producto_proveedor V  WHERE P.dni=E.cod_unico AND E.id_empresa='" & KEY_RUC & "' AND " & _
    "V.id_producto LIKE '%" & Trim(Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)) & "%' AND V.ruc='" & KEY_RUC & "' AND P.dni=V.id_proveedor ORDER BY P.nombre_completo"
    Call llenar_proveedor(Me.HfdDetalle, Me)
End If
End Sub
Private Sub HfdPersona_SelChange()
'If Me.HfProveedorProducto.Visible = True And Val(Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)) > 0 Then
strCadena = "SELECT V.id_producto,P.nombre_prod FROM producto P,producto_proveedor V WHERE P.id_producto=V.id_producto AND P.ruc='" & KEY_RUC & "' AND V.ruc='" & KEY_RUC & "' AND V.id_proveedor='" & Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0) & "'"
Call llena_producto(Me.HfProveedorProducto, Me)
'End If
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.key
    Case KEY_ASIGNAR
        Dim cprod As String
        Dim cProve As String
        cprod = Me.HfdGrilla.TextMatrix(Me.HfdGrilla.Row, 0)
        cProve = Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0)
        strCadena = "UPDATE producto SET id_proveedor='" & Trim(cProve) & "' WHERE id_producto='" & Trim(cprod) & "' AND ruc='" & KEY_RUC & "'"
        CnBd.Execute (strCadena)
         
        strCadena = "SELECT * FROM producto_proveedor WHERE id_producto='" & Trim(cprod) & "' AND id_proveedor='" & Trim(cProve) & "' AND ruc='" & KEY_RUC & "'"
        Call ConfiguraRst(strCadena)
        If rst.RecordCount < 1 Then
            strCadena = "INSERT INTO producto_proveedor(id_producto,id_proveedor,ruc) VALUES ('" & Trim(cprod) & "','" & Trim(cProve) & "','" & KEY_RUC & "')"
            CnBd.Execute (strCadena)
             
            Call HfdPersona_SelChange
        Else
            MsgBox "Producto ya Asigando a este Proveedor", vbInformation, "Mensaje para el Usuario"
            Procedencia = Selecionar
            Call ActualizarProd
            Exit Sub
            
        End If
        MsgBox "AGREGADO CON EXITO", vbInformation, KEY_EMPRESA
        If Trim(Me.txtProducto.Text) <> "" Then
            Call buscar_producto
        Else
            Call ActualizarProd
        End If
    Case KEY_EXIT
      Unload Me
  End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.key
   
      
    Case KEY_DELETE
      If MsgBox("ESTA SEGURO DE QUITAR ESTE PRODUCTO", vbQuestion + vbYesNo, KEY_EMPRESA) = vbYes Then
        strCadena = "DELETE  FROM producto_proveedor WHERE id_producto='" & Trim(Me.HfProveedorProducto.TextMatrix(Me.HfProveedorProducto.Row, 0)) & "' AND id_proveedor='" & Me.HfdPersona.TextMatrix(Me.HfdPersona.Row, 0) & "' AND ruc='" & KEY_RUC & "' "
        CnBd.Execute (strCadena)
         
        Call HfdPersona_SelChange
      End If
  End Select
End Sub

Private Sub TxtApellido_Change()
strCadena = "SELECT dni ,nombre_completo FROM persona P,entidad_empresa E WHERE P.dni=E.cod_unico AND E.id_empresa='" & KEY_RUC & "' AND nombre_completo LIKE '%" & Trim(Me.TxtApellido.Text) & "%' AND id_proveedor='si' ORDER BY nombre_completo ASC"
Call llenarGrid_per(Me.HfdPersona, Me)
End Sub
Private Sub TxtCod_Change()
  If Me.TxtCod.Text = "" Then
    Call ActualizarProd
    Exit Sub
  Else
  If Len(Me.TxtCod.Text) > 0 And KEY_BARRAS = "si" Then
    strCadena = "SELECT P.id_producto,P.nombre_prod FROM producto P,producto_barras B WHERE P.id_producto=B.id_producto AND B.cod_barra='" & Trim(Me.TxtCod.Text) & "' AND P.ruc='" & KEY_RUC & "' AND B.ruc='" & KEY_RUC & "' ORDER BY P.nombre_prod LIMIT 0,15"
  Else
    strCadena = "SELECT P.id_producto, P.nombre_prod  FROM producto P WHERE P.id_producto LIKE '%" & Trim(Me.TxtCod.Text) & "%' AND ruc='" & KEY_RUC & "' ORDER BY P.nombre_prod LIMIT 0,15"
  End If
  Call llenarGrid_prod(Me.HfdGrilla, Me)
  Me.TxtCod.SetFocus
End If
End Sub

Private Sub TxtProducto_Change()
Call buscar_producto
End Sub
Private Sub buscar_producto()
If Len(Me.txtProducto.Text) > 0 Then
    strCadena = "SELECT P.id_producto, P.nombre_prod  FROM producto P WHERE P.nombre_prod LIKE '%" & Trim(Me.txtProducto.Text) & "%' AND ruc='" & KEY_RUC & "' ORDER BY P.nombre_prod LIMIT 0,15"
Else
    strCadena = "SELECT P.id_producto, P.nombre_prod  FROM producto P WHERE P.nombre_prod LIKE '%" & Trim(Me.txtProducto.Text) & "%' AND ruc='" & KEY_RUC & "' ORDER BY P.nombre_prod LIMIT 0,15"
End If
Call llenarGrid_prod(Me.HfdGrilla, Me)
Me.txtProducto.SetFocus
End Sub

Private Sub TxtRuc_Change()
strCadena = "SELECT dni ,nombre_completo FROM persona P,entidad_empresa E WHERE P.dni=E.cod_unico AND E.id_empresa='" & KEY_RUC & "' AND dni LIKE '%" & Trim(Me.TxtRuc.Text) & "%' AND id_proveedor='si' ORDER BY nombre_completo ASC"
Call llenar_proveedor(Me.HfdPersona, Me)
End Sub
Sub llenar_proveedor(ByVal Grilla As MSHFlexGrid, ByVal Formulario As Form)
On Error GoTo salir
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
   
     
    Exit Sub
Else
     Me.Toolbar1.Buttons(KEY_DELETE).Enabled = True
End If
   Grilla.Clear
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 1200
           Grilla.ColWidth(1) = 5200
    Next
        cabecera = "DNI/RUC" & vbTab & "RAZON SOCIAL/PERSONA"
        Grilla.AddItem cabecera
         For k = 0 To 1
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            Fila = Fila & rst("dni") & vbTab & rst("nombre_completo")
            If (Fila = "") Then
                x = 1
            End If
          Grilla.AddItem Fila
            Fila = ""
            rst.MoveNext
    Next i
  Grilla.Row = 1
  Grilla.col = 0
  Grilla.ColSel = 1
  Grilla.RowSel = 1
  Exit Sub
salir:     MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing
End Sub

Sub llena_producto(ByVal Grilla As MSHFlexGrid, ByVal Formulario As Form)
On Error GoTo salir
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
     Me.Toolbar1.Buttons(KEY_DELETE).Enabled = False
     
    Exit Sub
Else
     Me.Toolbar1.Buttons(KEY_DELETE).Enabled = True
End If
   Grilla.Clear
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 700
           Grilla.ColWidth(1) = 5200
    Next
        cabecera = "CODIGO" & vbTab & "DESCRIPCION ARTICULO"
        Grilla.AddItem cabecera
         For k = 0 To 1
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            Fila = Fila & rst("id_producto") & vbTab & rst("nombre_prod")
            If (Fila = "") Then
                x = 1
            End If
          Grilla.AddItem Fila
            Fila = ""
            rst.MoveNext
    Next i
  Grilla.Row = 1
  Grilla.col = 0
  Grilla.ColSel = 1
  Grilla.RowSel = 1
  Exit Sub
salir:     MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing
End Sub

