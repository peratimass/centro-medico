VERSION 5.00
Begin VB.Form FrmGeneradorBarras 
   BorderStyle     =   0  'None
   Caption         =   "GENERADOR CODIGOS BARRA"
   ClientHeight    =   6255
   ClientLeft      =   2655
   ClientTop       =   870
   ClientWidth     =   10695
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C00000&
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6255
   ScaleWidth      =   10695
   ShowInTaskbar   =   0   'False
   Begin VitekeySoft.ChameleonBtn cmdImprimir 
      Height          =   735
      Left            =   5640
      TabIndex        =   15
      Top             =   2280
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1296
      BTYPE           =   5
      TX              =   "GENERAR IMPRESION"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
      MICON           =   "Form1.frx":030A
      PICN            =   "Form1.frx":0326
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.PictureBox CB 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1275
      Left            =   480
      ScaleHeight     =   1245
      ScaleWidth      =   3675
      TabIndex        =   14
      Top             =   2280
      Visible         =   0   'False
      Width           =   3705
   End
   Begin VB.TextBox txtColumnas 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Height          =   285
      Left            =   5040
      TabIndex        =   13
      Text            =   "2"
      Top             =   1320
      Width           =   660
   End
   Begin VB.TextBox txtcodBarra 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1635
      TabIndex        =   10
      Top             =   600
      Width           =   1575
   End
   Begin VB.OptionButton Option3 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "IMP-GRANDE"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   480
      TabIndex        =   9
      Top             =   4200
      Width           =   2175
   End
   Begin VB.OptionButton Option2 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "IMP-MEDIANA"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   3960
      Width           =   2175
   End
   Begin VB.OptionButton Option1 
      Appearance      =   0  'Flat
      BackColor       =   &H000080FF&
      Caption         =   "IMP-PEQUEÑA"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   3720
      Width           =   2175
   End
   Begin VB.TextBox txtFilas 
      Appearance      =   0  'Flat
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
      Height          =   285
      Left            =   1635
      TabIndex        =   5
      Top             =   1560
      Width           =   1500
   End
   Begin VB.TextBox TxtPrecio 
      Appearance      =   0  'Flat
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
      Height          =   285
      Left            =   1635
      TabIndex        =   2
      Top             =   1080
      Width           =   1500
   End
   Begin VB.TextBox txtDescripcion 
      Appearance      =   0  'Flat
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
      Height          =   285
      Left            =   3315
      TabIndex        =   1
      Top             =   600
      Width           =   5460
   End
   Begin VB.TextBox txtcodigo 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1635
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
   Begin VitekeySoft.ChameleonBtn ChameleonBtn1 
      Height          =   550
      Left            =   5760
      TabIndex        =   16
      Top             =   1200
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   979
      BTYPE           =   5
      TX              =   "AGREGAR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
      MICON           =   "Form1.frx":2AEF
      PICN            =   "Form1.frx":2B0B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn ChameleonBtn2 
      Height          =   550
      Left            =   7320
      TabIndex        =   17
      Top             =   1200
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   979
      BTYPE           =   5
      TX              =   "AGREGAR"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
      MICON           =   "Form1.frx":53F5
      PICN            =   "Form1.frx":5411
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdimpresionvisual 
      Height          =   615
      Left            =   5640
      TabIndex        =   18
      Top             =   3120
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1085
      BTYPE           =   5
      TX              =   "IMPRESION VISUAL"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
      MICON           =   "Form1.frx":785B
      PICN            =   "Form1.frx":7877
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VitekeySoft.ChameleonBtn cmdcerrar 
      Height          =   375
      Left            =   9960
      TabIndex        =   19
      Top             =   240
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BTYPE           =   5
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
      MICON           =   "Form1.frx":A161
      PICN            =   "Form1.frx":A17D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº ETIQUETAS X FILAS"
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
      Left            =   3480
      TabIndex        =   12
      Top             =   1320
      Width           =   1440
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   735
      Left            =   3315
      Top             =   1080
      Width           =   5415
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO BARRA :"
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
      Height          =   210
      Left            =   255
      TabIndex        =   11
      Top             =   720
      Width           =   1290
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Nº ETIQUETAS :"
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
      Height          =   210
      Left            =   360
      TabIndex        =   6
      Top             =   1665
      Width           =   1185
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PRECIO VENTA :"
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
      Height          =   210
      Left            =   345
      TabIndex        =   4
      Top             =   1185
      Width           =   1200
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CODIGO INTERNO :"
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
      Height          =   210
      Left            =   75
      TabIndex        =   3
      Top             =   240
      Width           =   1470
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      Height          =   6255
      Left            =   0
      Top             =   0
      Width           =   10695
   End
End
Attribute VB_Name = "FrmGeneradorBarras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Procedencia As EnumProcede

Private Sub ChameleonBtn4_Click()

End Sub

Private Sub cmdcerrar_Click()
Unload Me
End Sub

Private Sub cmdexit_Click()
    
   

End Sub

Private Sub cmdimpresionvisual_Click()
Dim In_fecha As String
Dim in_filas As Integer
strCadena = "DELETE FROM barra_impresion WHERE dni_save='" & KEY_USUARIO & "' and ruc='" & KEY_RUC & "'"
CnBd.Execute (strCadena)
In_fecha = get_last_buy(Trim(Me.txtcodigo.Text))
If Val(Me.txtFilas.Text) Mod 2 = 0 Then
    in_filas = Int(Val(Me.txtFilas.Text) / 2)
Else
    in_filas = Int(Val(Me.txtFilas.Text) / 2) + 1
End If



For i = 0 To in_filas - 1
    strCadena = "INSERT INTO barra_impresion(codigo,codigo_barra,descripcion,precio,fecha,dni_save,ruc)values " & _
    "('" & Trim(Me.txtcodigo.Text) & "','" & Trim(Me.txtcodBarra.Text) & "','" & Trim(Me.txtDescripcion.Text) & "','" & Val(Me.TxtPrecio.Text) & "','" & In_fecha & "','" & KEY_USUARIO & "','" & KEY_RUC & "')"
    CnBd.Execute (strCadena)
    
Next i



strCadena = "SELECT codigo,codigo_barra,descripcion,precio,fecha,'" & KEY_EMPRESA & "' FROM barra_impresion  WHERE dni_save='" & KEY_USUARIO & "' and  ruc='" & KEY_RUC & "' "
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    Ans = ShowMultiReport(rst, "rptBarras", , App.Path + "\Reportes\")
End If
End Sub

Private Sub Command1_Click()

End Sub



Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()

End Sub

Private Sub cmdPrint_Click()

End Sub

Private Sub cmdImprimir_Click()
For i = 0 To Val(Me.txtFilas.Text) - 1
    Call generar_stikers(Val(Me.txtcodigo.Text))
    Printer.TrackDefault = True 'si
    Printer.Font.Bold = True
    Printer.Print " "
    Printer.Font.Bold = False
    Printer.CurrentX = 0: Printer.CurrentY = 0
    Printer.PaintPicture CB.Image, 0, 0
    Printer.EndDoc
Next i
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 500

End Sub

Private Sub optSize_Click(index As Integer)
    

End Sub



Public Sub generar_stikers(ByVal in_activo As String)
Dim C128 As New BarCode128, tipo As eTipoDeCódigo128
Dim alto As Single, Den As Single
CB.Cls
tipo = cC128_B

Altura = 400
Densidad = 15

If IsNumeric(Altura) Then
    alto = CSng(Altura)
    If IsNumeric(Densidad) Then
        Den = CSng(Densidad)
        C128.GenerarBarras Trim(Me.txtcodBarra.Text), Trim(Me.txtDescripcion.Text), Trim(Me.txtcodigo.Text), KEY_FECHA, DC:=CB, ImprimirTexto:=True, Codificación:=tipo, y:=2, alto:=alto, Densidad:=Den
        
    Else
        C128.GenerarBarras Trim(Me.txtcodigo.Text), Trim(Me.txtDescripcion.Text), KEY_ALM, KEY_FECHA, DC:=CB, ImprimirTexto:=True, Codificación:=tipo, y:=2, alto:=alto
    End If
Else
    If IsNumeric(Densidad) Then
        Den = CSng(Densidad)
        C128.GenerarBarras Trim(Me.txtcodigo.Text), Trim(Me.txtDescripcion.Text), KEY_ALM, KEY_FECHA, CB, ImprimirTexto:=True, Codificación:=tipo, y:=2, Densidad:=Den
    Else
        C128.GenerarBarras Trim(Me.txtcodigo.Text), Trim(Me.txtDescripcion.Text), KEY_ALM, KEY_FECHA, CB, ImprimirTexto:=True, Codificación:=tipo, y:=2
    End If
End If
   
End Sub





Private Sub txtcodigo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.Option1.Value = True
    Call precionar(Trim(Me.txtcodigo.Text))
    
End If
End Sub

Public Sub precionar(ByVal in_producto As String)

If in_producto = "" Then
    Procedencia = Selecionar
    FrmProducto.Show
    Exit Sub
Else
    strCadena = "SELECT * FROM view_producto_barra WHERE id_producto='" & in_producto & "' and id_alm='" & KEY_ALM & "' and ruc='" & KEY_RUC & "' LIMIT 1"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        Me.txtcodigo.Text = rst("id_producto")
        Me.txtcodBarra.Text = rst("cod_barra")
        Me.TxtPrecio.Text = rst("precio_venta")
        Me.txtDescripcion.Text = rst("nombre_prod")
        Me.txtFilas.Text = 2
Else
    Procedencia = Selecionar
    FrmProducto.Show
    Exit Sub
End If
End If




End Sub


Private Sub listado_barras(ByVal Grilla As MSHFlexGrid)
strCadena = "SELECT * FROM generador_barras_temporal T,producto P,unidad U WHERE T.id_producto=P.id_producto AND P.id_unidad=U.id_und AND P.ruc='" & KEY_RUC & "' AND T.ruc='" & KEY_RUC & "' AND U.id_usu='" & KEY_RUC & "' AND T.dni_save='" & KEY_USUARIO & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Grilla.Clear
    Exit Sub
End If
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1300
           Grilla.ColWidth(2) = 4000
           Grilla.ColWidth(3) = 500
           Grilla.ColWidth(4) = 1200
           
       Next
        cabecera = "IDDETALLE" & vbTab & "CODIGO" & vbTab & "PRODUCTO" & vbTab & "UND" & vbTab & "ETIQUETAS"
        Grilla.AddItem cabecera
         For k = 0 To 4
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
        Next k
        rst.MoveFirst
        For i = 0 To rst.RecordCount - 1
          Fila = rst("id_detalle") & vbTab & rst("barra") & vbTab & rst("nombre_prod") & vbTab & rst("abreviatura") & vbTab & Format(rst("cantidad"), "#,##0.00")
          Grilla.AddItem Fila
          Fila = ""
          rst.MoveNext
        Next i
  End Sub

