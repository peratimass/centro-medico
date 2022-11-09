VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmPlanta 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "MOVIMIENTO PLANTA"
   ClientHeight    =   9645
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   16710
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9645
   ScaleWidth      =   16710
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdSalir 
      Caption         =   "Salir"
      Height          =   855
      Left            =   9960
      Picture         =   "FrmPlanta.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4850
      Width           =   855
   End
   Begin VB.CommandButton CmdAnular 
      Caption         =   "Anular"
      Height          =   855
      Left            =   9960
      Picture         =   "FrmPlanta.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2280
      Width           =   855
   End
   Begin VB.CommandButton CmdGuardar 
      Caption         =   "Guardar"
      Height          =   855
      Left            =   9960
      Picture         =   "FrmPlanta.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3135
      Width           =   855
   End
   Begin VB.Frame Frame1 
      ForeColor       =   &H00800000&
      Height          =   9135
      Left            =   11160
      TabIndex        =   0
      Top             =   360
      Width           =   5295
      Begin VB.Timer Timer_Gas 
         Interval        =   2000
         Left            =   1560
         Top             =   3600
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgDetalle 
         Height          =   3615
         Left            =   840
         TabIndex        =   4
         Top             =   4920
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   6376
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
      Begin VB.Label Label5 
         BackColor       =   &H00FFFF00&
         BorderStyle     =   1  'Fixed Single
         Height          =   4035
         Left            =   380
         TabIndex        =   27
         Top             =   240
         Width           =   4335
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H000000C0&
         BackStyle       =   1  'Opaque
         Height          =   3855
         Left            =   0
         Top             =   360
         Width           =   4140
      End
      Begin VB.Label lblVacio 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         ForeColor       =   &H80000008&
         Height          =   1035
         Left            =   965
         TabIndex        =   26
         Top             =   0
         Width           =   4140
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "50 000"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblKilogramos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "........."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   2370
         TabIndex        =   3
         Top             =   1800
         Width           =   1080
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "CONTROL DE EXISTENCIAS"
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
         Left            =   1920
         TabIndex        =   2
         Top             =   4560
         Width           =   2145
      End
      Begin VB.Shape Shape2 
         BorderStyle     =   3  'Dot
         Height          =   4215
         Left            =   720
         Top             =   4440
         Width           =   4335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "25 000"
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   2160
         Width           =   495
      End
   End
   Begin TabDlg.SSTab SSTab2 
      Height          =   9015
      Left            =   120
      TabIndex        =   5
      Top             =   480
      Width           =   10965
      _ExtentX        =   19341
      _ExtentY        =   15901
      _Version        =   393216
      TabHeight       =   520
      ForeColor       =   8388608
      TabCaption(0)   =   "ENVASADO"
      TabPicture(0)   =   "FrmPlanta.frx":0EDE
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Shape3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label3"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "HfPedido"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "HfProductos"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "CmdNuevo"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame2"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "CmdBuscar"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdQuitar"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "REVICION BALONES"
      TabPicture(1)   =   "FrmPlanta.frx":0EFA
      Tab(1).ControlEnabled=   0   'False
      Tab(1).ControlCount=   0
      TabCaption(2)   =   "SISTERNA AMBULANTE"
      TabPicture(2)   =   "FrmPlanta.frx":0F16
      Tab(2).ControlEnabled=   0   'False
      Tab(2).ControlCount=   0
      Begin VB.CommandButton cmdQuitar 
         Height          =   255
         Left            =   8880
         Picture         =   "FrmPlanta.frx":0F32
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   6240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   855
         Left            =   9840
         Picture         =   "FrmPlanta.frx":14BC
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   3500
         Width           =   855
      End
      Begin VB.Frame Frame2 
         Height          =   1815
         Left            =   120
         TabIndex        =   10
         Top             =   480
         Width           =   9255
         Begin VB.TextBox TxtNumero 
            Appearance      =   0  'Flat
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
            Height          =   315
            Left            =   6240
            MaxLength       =   80
            TabIndex        =   19
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox TxtSerie 
            Appearance      =   0  'Flat
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
            Height          =   315
            Left            =   5400
            MaxLength       =   80
            TabIndex        =   18
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox TxtDireccion 
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   1530
            MaxLength       =   80
            TabIndex        =   14
            Top             =   1260
            Width           =   4815
         End
         Begin VB.TextBox TxtCliente 
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   1530
            MaxLength       =   80
            TabIndex        =   13
            Top             =   900
            Width           =   4815
         End
         Begin VB.TextBox TxtCodCliente 
            Appearance      =   0  'Flat
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   150
            MaxLength       =   80
            TabIndex        =   12
            Top             =   900
            Width           =   1335
         End
         Begin MSDataListLib.DataCombo DtcComprobante 
            Height          =   315
            Left            =   2760
            TabIndex        =   11
            Top             =   240
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Text            =   "DataCombo1"
         End
         Begin MSDataListLib.DataCombo DtcAlmacen 
            Height          =   315
            Left            =   120
            TabIndex        =   21
            Top             =   240
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Text            =   "DataCombo1"
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Dirección:"
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
            Left            =   225
            TabIndex        =   15
            Top             =   1260
            Width           =   885
         End
      End
      Begin VB.CommandButton CmdNuevo 
         Caption         =   "Nuevo"
         Height          =   855
         Left            =   9840
         Picture         =   "FrmPlanta.frx":17C6
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   960
         Width           =   855
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfProductos 
         Height          =   2775
         Left            =   120
         TabIndex        =   9
         Top             =   2760
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   4895
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfPedido 
         Height          =   2655
         Left            =   120
         TabIndex        =   20
         Top             =   6000
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   4683
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
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Datos de Envasado"
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
         Left            =   90
         TabIndex        =   23
         Top             =   5640
         Width           =   1635
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Datos de Almacen"
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
         Left            =   135
         TabIndex        =   22
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Shape Shape3 
         Height          =   4575
         Left            =   9720
         Shape           =   4  'Rounded Rectangle
         Top             =   840
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmPlanta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede

Private Sub CmdGuardar_Click()
If Me.TxtCodCliente.Text <> "" And Me.TxtCliente.Text <> "" Then
    Call Save
End If
If Me.TxtCodCliente.Text = "" Or Me.TxtCliente.Text = "" Then
    MsgBox "Ingrese un Cliente para este Movimiento", vbInformation, "Mensaje para el Usuario"
    Call Resalta(Me.TxtCodCliente)
End If
End Sub
Private Sub Save()
Dim IdEnvase As Double
strCadena = "INSERT INTO DocumentoEnvase (doc_cod,alm_cod,serie,numero,cPersona,fecha,Ruc)VALUES('" & Me.DtcComprobante.BoundText & "','" & Trim(Me.DtcAlmacen.BoundText) & "','" & Me.TxtSerie.Text & "','" & Me.TxtNumero.Text & "','" & Trim(Me.TxtCodCliente.Text) & "','" & CVDate(Date) & "','" & KEY_RUC & "')"
CnBd.Execute (strCadena)
 
IdEnvase = IdInsert("DocumentoEnvase")
Call DetalleEnvase(IdEnvase)
nuevo_numero = formato_item(Val(Me.TxtNumero.Text) + 1, 10)
strCadena = "UPDATE  Det_alm_com SET numero='" & Trim(nuevo_numero) & "'  WHERE (serie='" & Trim(Me.TxtSerie.Text) & "' AND doc_cod='" & Trim(Me.DtcComprobante.BoundText) & "')"
CnBd.Execute (strCadena)
 
Call llenarProductos(Me.HfProductos, Me)
Call nuevo

End Sub
Private Sub DetalleEnvase(ByVal IdEnvase As Double)
strCadena = "SELECT * FROM Temporal_Envasado WHERE serie='" & Trim(Me.TxtSerie.Text) & "' AND numero='" & Trim(Me.TxtNumero.Text) & "' AND id_usuario='" & KEY_USUARIO & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    rst.MoveFirst
    For i = 0 To rst.RecordCount - 1
        strCadena = "INSERT INTO DocumentoEnvase_detalle(idEnvase,cProducto,cantidad,total)VALUES('" & IdEnvase & "','" & rst("cProducto") & "','" & rst("cantidad") & "','" & rst("totalKg") & "')"
        CnBd.Execute (strCadena)
         
        rst.MoveNext
    Next i
End If
End Sub
Private Sub cmdNuevo_Click()
Call nuevo
End Sub
Public Sub nuevo()
strCadena = "DELETE FROM Temporal_Envasado WHERE id_usuario='" & KEY_USUARIO & "'"
CnBd.Execute (strCadena)
 
Call llenarEnvasado(Me.HfPedido, Me)
Me.CmdGuardar.Enabled = False
Me.CmdAnular.Enabled = False
Me.TxtCodCliente.Text = ""
Me.TxtCliente.Text = ""
 strCadena = "SELECT * FROM Det_alm_com WHERE doc_cod='" & Me.DtcComprobante.BoundText & "' AND serie='" & Trim(Me.TxtSerie.Text) & "' AND Ruc='" & KEY_RUC & "'"
  Call ConfiguraRst(strCadena)
  If rst.RecordCount > 0 Then
    Me.TxtSerie.Text = rst("serie")
    Me.TxtNumero.Text = rst("numero")
  End If
Call Resalta(Me.TxtCodCliente)
End Sub
Private Sub CmdQuitar_Click()
strCadena = "DELETE FROM Temporal_envasado WHERE idtemporal='" & Val(Me.HfPedido.TextMatrix(Me.HfPedido.Row, 0)) & "'"
CnBd.Execute (strCadena)
 
Call llenarEnvasado(Me.HfPedido, Me)
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 50
    strCadena = "SELECT Alm_cod as Codigo, Alm_des as Descripcion FROM Almacen  ORDER BY Alm_des"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcAlmacen)
  Me.DtcAlmacen.BoundText = "0001"
  Set rst = Nothing
   
  
  strCadena = "SELECT doc_cod as Codigo, doc_des as Descripcion FROM Comprobantes " & _
  " WHERE doc_tienda='V' ORDER BY doc_abrev"
  Call ConfiguraRst(strCadena)
  Call LlenaDataCombo(Me.DtcComprobante)
  Me.DtcComprobante.BoundText = "0104"
  
  strCadena = "SELECT * FROM Det_alm_com WHERE doc_cod='" & Me.DtcComprobante.BoundText & "' AND Ruc='" & KEY_RUC & "'"
  Call ConfiguraRst(strCadena)
  If rst.RecordCount > 0 Then
    Me.TxtSerie.Text = rst("serie")
    Me.TxtNumero.Text = rst("numero")
  End If
  Call llenarProductos(Me.HfProductos, Me)
   Call llenarEnvasado(Me.HfPedido, Me)
 
 
  
End Sub
Public Sub llenarProductos(ByVal Grilla As MSHFlexGrid, ByVal Formulario As Form)
On Error GoTo salir
strCadena = "SELECT  Almacen_Productos.cProducto, Producto.DescripcionProducto, Unidad.sAbreviatura, Almacen_Productos.Stock  " & _
  " FROM         Almacen_Productos INNER JOIN " & _
  "Producto ON Almacen_Productos.cProducto = Producto.cProducto INNER JOIN Unidad ON Producto.cUnidad = Unidad.cUnidad " & _
 " WHERE Almacen_Productos.Alm_cod='" & Trim(KEY_ALM) & "' AND Almacen_Productos.ruc='" & Trim(KEY_RUC) & "' AND sub_producto='V' ORDER BY  DescripcionProducto "
    'Call llenarProductos(Me.HfProductos, Me)
  
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
           Grilla.ColWidth(0) = 700
           Grilla.ColWidth(1) = 5200
           Grilla.ColWidth(2) = 700
           Grilla.ColWidth(3) = 900
        Next
        cabecera = "CODIGO" & vbTab & "DESCRIPCION ARTICULO" & vbTab & "UND" & vbTab & "STOCK"
        Grilla.AddItem cabecera
         For k = 0 To 3
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
            Fila = Fila & rst("cProducto") & vbTab & rst("DescripcionProducto") & vbTab & rst("sAbreviatura") & vbTab & rst("Stock")
            If (Fila = "") Then
                x = 1
            End If
          Grilla.AddItem Fila
                        
                        If (Trim(rst("Stock")) < 2) Then
                            For k = 0 To 3
                                Grilla.col = k
                                Grilla.Row = i + 1
                                Grilla.CellBackColor = &H8080FF
                            Next k
                      End If
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
Public Sub llenarEnvasado(ByVal Grilla As MSHFlexGrid, ByVal Formulario As Form)
Dim PesoTotal As Double
On Error GoTo salir
strCadena = "SELECT Temporal_Envasado.idTemporal,Temporal_Envasado.cProducto, Producto.DescripcionProducto, Unidad.sAbreviatura, Temporal_Envasado.cantidad, " & _
"  Temporal_Envasado.totalKg FROM Temporal_Envasado INNER JOIN Producto ON Temporal_Envasado.cProducto = Producto.cProducto INNER JOIN " & _
" Unidad ON Producto.cUnidad = Unidad.cUnidad WHERE serie='" & Trim(Me.TxtSerie.Text) & "' AND numero='" & Trim(Me.TxtNumero.Text) & "'"
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
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 700
           Grilla.ColWidth(2) = 5200
           Grilla.ColWidth(3) = 700
           Grilla.ColWidth(4) = 900
           Grilla.ColWidth(5) = 900
        Next
        cabecera = "CODIGO" & vbTab & "CODIGO" & vbTab & "DESCRIPCION ARTICULO" & vbTab & "UND" & vbTab & "CANTIDAD" & vbTab & "TOTAL"
        Grilla.AddItem cabecera
         For k = 0 To 5
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        PesoTotal = 0
        For i = 0 To rst.RecordCount - 1
            Fila = rst("idTemporal") & vbTab & rst("cProducto") & vbTab & rst("DescripcionProducto") & vbTab & rst("sAbreviatura") & vbTab & rst("cantidad") & vbTab & rst("totalKg")
            If (Fila = "") Then
                x = 1
            End If
          PesoTotal = PesoTotal + rst("totalKg")
          Grilla.AddItem Fila
            Fila = ""
            rst.MoveNext
             
        Next i
            Fila = "" & vbTab & "" & vbTab & "ACUMULADO EN KG" & vbTab & "====" & vbTab & "====" & vbTab & Format(PesoTotal, "###0.00")
            Grilla.AddItem Fila
            For k = 0 To 5
            Grilla.col = k
            Grilla.Row = i + 1
            Grilla.CellBackColor = &HC0FFC0
        Next k
  Grilla.Row = 1
  Grilla.col = 0
  Grilla.ColSel = 1
  Grilla.RowSel = 1
    Exit Sub
salir: MsgBox "Ocurrio un Error en la Conexion", vbInformation, "Mensaje para el usuario"
Set rst = Nothing

End Sub


Private Sub HfPedido_Click()
If Val(Me.HfPedido.TextMatrix(Me.HfPedido.Row, 0)) > 0 Then
    Me.cmdQuitar.Visible = True
Else
    Me.cmdQuitar.Visible = False
End If
End Sub

Private Sub HfProductos_DblClick()
If Val(Me.HfProductos.TextMatrix(Me.HfProductos.Row, 0)) > 0 Then
    strCadena = "SELECT * FROM Producto WHERE cProducto='" & Me.HfProductos.TextMatrix(Me.HfProductos.Row, 0) & "' AND Ruc='" & KEY_RUC & "'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
    FrmAddEnvasado.lblcodProducto.Caption = Me.HfProductos.TextMatrix(Me.HfProductos.Row, 0)
    FrmAddEnvasado.LblProducto.Caption = Me.HfProductos.TextMatrix(Me.HfProductos.Row, 1)
    FrmAddEnvasado.lblpeso.Caption = Format(rst("prod_peso"), "#,##0.00")
    FrmAddEnvasado.Top = Me.HfProductos.Top
    'FrmAddEnvasado.Left = FrmAddEnvasado.Left + 500
    FrmAddEnvasado.Show
    End If
End If
End Sub

Private Sub HfProductos_SelChange()
'If Val(Me.HfProductos.TextMatrix(Me.HfProductos.Row, 0)) > 0 Then
    'Me.txt
'End If
End Sub



Private Sub Timer_Gas_Timer()
Dim Total As Double
Dim rstGas As New ADODB.Recordset
Total = 50000
strCadena = "SELECT  Stock FROM Almacen_Productos WHERE cProducto='00006' AND Ruc='" & KEY_RUC & "' AND alm_cod='" & KEY_ALM & "'"
 rstGas.Open strCadena, CnBd, adOpenKeyset, adLockOptimistic
If rstGas.RecordCount > 0 Then
    porcentaje = Val(Me.Shape1.Height) - Val(Me.Shape1.Height) * rstGas("Stock") / Total
    Me.lblVacio.Height = porcentaje
    Me.lblKilogramos.Caption = Format(rstGas("Stock"), "###0.00") + Space(2) + "KG"
End If
Set rstGas = Nothing

End Sub

Private Sub TxtCodCliente_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then

If Len(Me.TxtCodCliente.Text) = 11 Or Len(Me.TxtCodCliente.Text) = 8 Then
    strCadena = "SELECT cPersona,NombrePersona,sDireccionCliente1,Per_Ruc,Observacion FROM " & _
    " Persona WHERE Per_Ruc='" & Trim(Me.TxtCodCliente.Text) & "'"
Else
    strCadena = "SELECT cPersona,NombrePersona,sDireccionCliente1,Per_Ruc,Observacion FROM " & _
    " Persona WHERE cPersona='" & Trim(Me.TxtCodCliente.Text) & "'"
End If
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        
        Me.TxtCodCliente.Text = rst(0)
        Me.TxtCliente.Text = rst(1)
        Me.TxtDireccion.Text = rst(2)
        
        'Me.TxtObservacion.Text = rst(4)
        'Call Resalta(Me.TxtObservacion)
    Else
        Procedencia = Selecionar
        FrmPersona.Show
    End If
    
End If
End Sub
