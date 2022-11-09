VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form FrmDocumentos 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5910
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11535
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "FrmDocumentos.frx":0000
   ScaleHeight     =   5910
   ScaleWidth      =   11535
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtcambioLocal 
      BackColor       =   &H000080FF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   5520
      TabIndex        =   9
      Top             =   4320
      Width           =   1110
   End
   Begin VitekeySoft.ChameleonBtn Command1 
      Height          =   495
      Left            =   8280
      TabIndex        =   3
      Top             =   5280
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   873
      BTYPE           =   4
      TX              =   "ACEPTAR RESPONSABILIDAD"
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
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "FrmDocumentos.frx":78A20
      PICN            =   "FrmDocumentos.frx":78A3C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfgComprobante 
      Height          =   2775
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   4895
      _Version        =   393216
      ForeColor       =   8388608
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      ForeColorFixed  =   12582912
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
   Begin VB.Label lblcambioCompra 
      BackColor       =   &H0080C0FF&
      Caption         =   "3.2345"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   5520
      TabIndex        =   8
      Top             =   3840
      Width           =   1110
   End
   Begin VB.Label lblcambioVenta 
      BackColor       =   &H0080C0FF&
      Caption         =   "3.2345"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   330
      Left            =   5520
      TabIndex        =   7
      Top             =   3360
      Width           =   1110
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TIPO CAMBIO     [LOCAL ] :"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   960
      TabIndex        =   6
      Top             =   4320
      Width           =   4290
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TIPO CAMBIO SBS [COMPRA] :"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   960
      TabIndex        =   5
      Top             =   3840
      Width           =   4290
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TIPO CAMBIO SBS [VENTA ] :"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   315
      Left            =   960
      TabIndex        =   4
      Top             =   3360
      Width           =   4290
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      Height          =   1575
      Left            =   360
      Top             =   3240
      Width           =   10815
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00404040&
      BorderWidth     =   4
      Height          =   4590
      Left            =   0
      Top             =   0
      Width           =   10050
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "USTED ES RESPONSABLE DE LA EMISIÓN DE CUALQUIER COMPROBANTE DENTRO DE SU TURNO"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   360
      TabIndex        =   2
      Top             =   5040
      Width           =   9360
   End
   Begin VB.Label lblventanilla 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "VERIFIQUE SU SERIE DE COMPROBANTES."
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   4200
   End
End
Attribute VB_Name = "FrmDocumentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub actualizacomp(ByVal Grilla As MSHFlexGrid, ByVal idalmacen As String)
On Error GoTo salir
If KEY_COMPROBANTES_PROPIOS = "si" Then
    strCadena = "SELECT * FROM view_comprobante_almacen WHERE ruc='" & KEY_RUC & "' AND id_alm='" & KEY_VENTANILLA & "'"
Else
    strCadena = "SELECT * FROM view_comprobante_almacen WHERE ruc='" & KEY_RUC & "' AND id_alm='" & KEY_ALM & "'"
End If
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 0
    Exit Sub

End If
 
 
   Grilla.Rows = 0
      
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
           Grilla.ColWidth(0) = 0
           Grilla.ColWidth(1) = 1000
           Grilla.ColWidth(2) = 5000
           Grilla.ColWidth(3) = 1000
           Grilla.ColWidth(4) = 1200
           
        Next
         
         cabecera = "CODIGO" & vbTab & "SUNAT" & vbTab & "COMPROBANTE" & vbTab & "SERIE" & vbTab & "NUMERO"
         Grilla.AddItem cabecera
         
         For k = 0 To 4
            Grilla.col = k
            Grilla.Row = 0
            Grilla.CellBackColor = &HDFDFE0
          Next k
                            
        rst.MoveFirst
        For i = 1 To rst.RecordCount
             Fila = "" & vbTab & rst("id_doc") & vbTab & rst("doc_des") & vbTab & rst("serie") & vbTab & rst("numero")
             Grilla.AddItem Fila
             If rst("defecto") = "si" Then
        
                            For k = 0 To 4
                                Grilla.col = k
                                Grilla.Row = i
                                Grilla.CellBackColor = &H8080FF
                            Next k
        
             End If
            
        
        rst.MoveNext
        Next i
        
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub


Private Sub Command1_Click()
If Val(Me.txtcambioLocal.Text) = 0 Then
    MsgBox "EL TIPO CAMBIO LOCAL NO DEBE ESTAR EN 0" + Chr(13) + "EL SISTEMA ACTUALIZARÁ AL TC SUNAT", vbInformation
    Me.txtcambioLocal.Text = Val(Me.lblcambioCompra.Caption)
End If

strCadena = "UPDATE tipo_cambio SET valor_local='" & Val(Me.txtcambioLocal.Text) & "' WHERE fecha='" & KEY_FECHA & "' and id_creador='" & KEY_RUC & "'"
CnBd.Execute (strCadena)
KEY_CAMBIO_LOCAL = Val(Me.txtcambioLocal.Text)
MDIFrmPrincipal.StatusBar1.Panels(3) = "FECHA:" + Space(2) + Format$(KEY_FECHA, "dd-mm-yyyy") & Space(3) & "TIPO CAMBIO LOCAL:" & KEY_CAMBIO_LOCAL
MDIFrmPrincipal.Toolbar1.Enabled = True
Unload Me
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 200

Call Me.actualizacomp(Me.HfgComprobante, KEY_ALM)
Me.lblcambioVenta.Caption = KEY_CAMBIO_VENTA
Me.lblcambioCompra.Caption = KEY_CAMBIO_COMPRA
Me.txtcambioLocal.Text = KEY_CAMBIO_LOCAL

End Sub
