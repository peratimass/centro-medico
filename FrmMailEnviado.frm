VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMailEnviado 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6375
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11175
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   11175
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtBuscar 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   4920
      TabIndex        =   5
      Top             =   240
      Width           =   2295
   End
   Begin VB.TextBox TxtDetalle 
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
      ForeColor       =   &H00400000&
      Height          =   1095
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   5160
      Width           =   9375
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Salir"
      Height          =   435
      Left            =   9720
      TabIndex        =   1
      Top             =   5880
      Width           =   1395
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdMails 
      Height          =   4335
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   7646
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
   Begin MSDataListLib.DataCombo DtcPersona 
      Height          =   315
      Left            =   960
      TabIndex        =   3
      Top             =   240
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ForeColor       =   8388608
      Text            =   "DataCombo1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Persona:"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   225
      TabIndex        =   4
      Top             =   280
      Width           =   645
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   6375
      Left            =   0
      Top             =   0
      Width           =   11175
   End
End
Attribute VB_Name = "FrmMailEnviado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
Unload Me
End Sub

Private Sub DtcPersona_Change()
strCadena = "SELECT * FROM persona_mail WHERE dni='" & Trim(Me.DtcPersona.BoundText) & "' AND ruc='" & KEY_RUC & "' ORDER BY fecha ASC"
Call llenarGrid(Me.HfdMails, Me)
End Sub

Private Sub Form_Load()
CenterForm Me
Me.Top = 800
strCadena = "SELECT P.dni as Codigo, P.nombre_completo as Descripcion FROM persona P,entidad_empresa E WHERE P.dni=E.cod_unico AND E.id_empresa='" & KEY_RUC & "' ORDER BY nombre_completo"
  Call ConfiguraRstT(strCadena)
  Call LlenaDataComboT(Me.DtcPersona)
  strCadena = "SELECT * FROM persona_mail WHERE dni='" & Trim(FrmMail.TxtPersona.Text) & "' AND ruc='" & KEY_RUC & "' ORDER BY fecha ASC"
  Call llenarGrid(Me.HfdMails, Me)
End Sub
Private Sub llenarGrid(ByVal Grilla As MSHFlexGrid, ByVal formulario As Form)
On Error GoTo salir
Call ConfiguraRst(strCadena)
If rst.RecordCount < 1 Then
    Grilla.Rows = 1
    Grilla.Clear
    Exit Sub
End If
  n = 1
   Grilla.Clear
   Grilla.Refresh
   Grilla.Rows = 0
 
       ReDim arrColWidth(1 To rst.Fields.Count)
       For Each Campo In rst.Fields
            ' -- almacena el ancho de los campos
           Grilla.ColWidth(0) = 600
           Grilla.ColWidth(1) = 1000
           Grilla.ColWidth(2) = 2500
           Grilla.ColWidth(3) = 6500
           
           
       Next
         cabecera = "ITEM" & vbTab & "FECHA" & vbTab & "MOTIVO" & vbTab & "DETALLE"
         Grilla.AddItem cabecera
         For k = 0 To 3
                                Grilla.Col = k
                                Grilla.Row = 0
                                Grilla.CellBackColor = &HDFDFE0
                            Next k
                            
        rst.MoveFirst
        For i = 1 To rst.RecordCount
             fila = fila & Str(rst.RecordCount - i + 1) & vbTab & rst("fecha") & vbTab & rst("motivo") & vbTab & rst("detalle")
            If (fila = "") Then
                x = 1
            End If
            
          Grilla.AddItem fila
            
        fila = ""
        rst.MoveNext
        Next i
Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub


Private Sub HfdMails_SelChange()
If Len(Me.HfdMails.TextMatrix(Me.HfdMails.Row, 3)) > 0 Then
    Me.TxtDetalle.Text = Me.HfdMails.TextMatrix(Me.HfdMails.Row, 3)
Else
    Me.TxtDetalle.Text = ""
End If
End Sub

Private Sub TxtBuscar_Change()
strCadena = "SELECT P.dni as Codigo, P.nombre_completo as Descripcion FROM persona P,entidad_empresa E WHERE P.dni=E.cod_unico AND E.id_empresa='" & KEY_RUC & "' AND P.nombre_completo LIKE '%" & Trim(Me.TxtBuscar.Text) & "%' ORDER BY nombre_completo"
  Call ConfiguraRstT(strCadena)
  Call LlenaDataComboT(Me.DtcPersona)
End Sub
