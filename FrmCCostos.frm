VERSION 5.00
Begin VB.Form FrmCCostos 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "CENTRO DE COSTOS"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4860
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4860
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FrameCCostos 
      BackColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   120
      TabIndex        =   9
      Top             =   0
      Width           =   4575
      Begin VB.TextBox Monto3 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3360
         MaxLength       =   80
         TabIndex        =   6
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox TxtNaturaleza 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   120
         MaxLength       =   80
         TabIndex        =   0
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox TxtCostos1 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   2040
         MaxLength       =   80
         TabIndex        =   1
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtCostos3 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   2040
         MaxLength       =   80
         TabIndex        =   5
         Top             =   1320
         Width           =   1215
      End
      Begin VB.TextBox txtCostos2 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   2040
         MaxLength       =   80
         TabIndex        =   3
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtCostos4 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   2040
         MaxLength       =   80
         TabIndex        =   7
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox Monto1 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3360
         MaxLength       =   80
         TabIndex        =   2
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Monto2 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3360
         MaxLength       =   80
         TabIndex        =   4
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Monto4 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3360
         MaxLength       =   80
         TabIndex        =   8
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox TxtTotalCC 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3360
         MaxLength       =   80
         TabIndex        =   10
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label LblDescripcion2 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1275
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   1800
      End
      Begin VB.Label Label14 
         Caption         =   "Gº. NATURALEZA"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label18 
         Caption         =   "C.COSTOS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   2040
         TabIndex        =   13
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "IMPORTE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   3360
         TabIndex        =   12
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "TOTAL C.C"
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
         Left            =   2040
         TabIndex        =   11
         Top             =   2040
         Width           =   840
      End
   End
End
Attribute VB_Name = "FrmCCostos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public busqueda As EnumCostos
Public Procedencia As EnumProcede
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    FrmRegistroComprasList.cmdAgregar.SetFocus
    Unload Me
End If
End Sub

Private Sub Form_Load()
Me.Left = FrmRegistroComprasList.cmdAgregar.Left - Me.Width
Me.Top = FrmRegistroComprasList.cmdAgregar.Top - Me.Height
If FrmRegistroComprasList.strModificar = True Then
strCadena = "SELECT * FROM registrocomprascostosnaturaleza WHERE codigounico='" & FrmRegistroComprasList.HfdPersona.TextMatrix(FrmRegistroComprasList.HfdPersona.Row, 0) & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
   Me.TxtNaturaleza.Text = rst("cnaturaleza")
   
End If
Set rst = Nothing
strCadena = "SELECT * FROM registrocomprasnaturaleza_costos WHERE codigounico='" & FrmRegistroComprasList.HfdPersona.TextMatrix(FrmRegistroComprasList.HfdPersona.Row, 0) & "'"
Call ConfiguraRst(strCadena)
If rst.RecordCount > 0 Then
    rst.MoveFirst
    If rst.RecordCount = 1 Then
        Me.TxtCostos1.Text = rst("ccostos")
        ccostos1G = rst("ccostos")
        Me.Monto1.Text = Format(rst("monto"), "###0.00")
    End If
    If rst.RecordCount = 2 Then
        Me.TxtCostos1.Text = rst("ccostos")
        ccostos1G = rst("ccostos")
        Me.Monto1.Text = Format(rst("monto"), "###0.00")
        rst.MoveNext
        Me.txtCostos2.Text = rst("ccostos")
        ccostos2G = rst("ccostos")
        Me.Monto2.Text = Format(rst("monto"), "###0.00")
    End If
    If rst.RecordCount = 3 Then
        Me.TxtCostos1.Text = rst("ccostos")
         ccostos1G = rst("ccostos")
        Me.Monto1.Text = Format(rst("monto"), "###0.00")
        rst.MoveNext
        Me.txtCostos2.Text = rst("ccostos")
         ccostos2G = rst("ccostos")
        Me.Monto2.Text = Format(rst("monto"), "###0.00")
        rst.MoveNext
        Me.txtCostos3.Text = rst("ccostos")
         ccostos3G = rst("ccostos")
        Me.Monto3.Text = Format(rst("monto"), "###0.00")
    End If
    If rst.RecordCount = 4 Then
        Me.TxtCostos1.Text = rst("ccostos")
         ccostos1G = rst("ccostos")
        Me.Monto1.Text = Format(rst("monto"), "###0.00")
        rst.MoveNext
        Me.txtCostos2.Text = rst("ccostos")
         ccostos2G = rst("ccostos")
        Me.Monto2.Text = Format(rst("monto"), "###0.00")
        rst.MoveNext
        Me.txtCostos3.Text = rst("ccostos")
        ccostos3G = rst("ccostos")
        Me.Monto3.Text = Format(rst("monto"), "###0.00")
        rst.MoveNext
        Me.txtCostos4.Text = rst("ccostos")
        ccostos4G = rst("ccostos")
        Me.Monto4.Text = Format(rst("monto"), "###0.00")
    End If
        
    End If
End If



End Sub

Private Sub Monto1_Change()
Dim montoY As Single
Me.TxtTotalCC.Text = Format(Val(Me.Monto1.Text) + Val(Me.Monto2.Text) + Val(Me.Monto3.Text) + Val(Me.Monto4.Text), "###0.00")
 If Val(FrmRegistroComprasList.TxtValorcompra.Text) = 0 Then
    montoY = Val(FrmRegistroComprasList.TxtValorCompraNoAfecta.Text)
  Else
     montoY = Val(FrmRegistroComprasList.TxtValorcompra.Text)
    End If
 Val (FrmRegistroComprasList.TxtValorcompra.Text)
If Val(Me.TxtTotalCC.Text) > montoY Then
    MsgBox "Monto Ingresado Incorrecto", vbInformation, "Mensaje para el Usuario"
    Call Resalta(Me.Monto1)
End If

End Sub

Private Sub Monto1_KeyPress(KeyAscii As Integer)
Dim valorV As Single
If KeyAscii = 13 Then
    Me.Monto1.Text = Format(Val(Me.Monto1.Text), "###0.00")
    Call Resalta(Me.txtCostos2)
    If Val(FrmRegistroComprasList.TxtValorcompra.Text) = 0 Then
        valorV = Val(FrmRegistroComprasList.TxtValorCompraNoAfecta.Text)
    Else
        valorV = Val(FrmRegistroComprasList.TxtValorcompra.Text)
    End If
    If (Val(Me.TxtTotalCC.Text) = valorV) Then
       
        cNaturaleza = Trim(Me.TxtNaturaleza.Text)
        ccostos1 = Trim(Me.TxtCostos1.Text)
        cMonto1 = Val(Me.Monto1.Text)
        ccostos2 = Trim(Me.txtCostos2.Text)
        cMonto2 = Val(Me.Monto2.Text)
        ccostos3 = Trim(Me.txtCostos3.Text)
        cMonto3 = Val(Me.Monto3.Text)
        ccostos4 = Trim(Me.txtCostos4.Text)
        cMonto4 = Val(Me.Monto4.Text)
        Unload Me
        FrmRegistroComprasList.cmdAgregar.SetFocus
         
        Exit Sub
    End If
End If
End Sub

Private Sub Monto2_Change()
Me.TxtTotalCC.Text = Format(Val(Me.Monto1.Text) + Val(Me.Monto2.Text) + Val(Me.Monto3.Text) + Val(Me.Monto4.Text), "###0.00")
If Val(Me.TxtTotalCC.Text) > Val(FrmRegistroComprasList.TxtValorcompra.Text) Then
    MsgBox "Monto Ingresado Incorrecto", vbInformation, "Mensaje para el Usuario"
    Call Resalta(Me.Monto2)
End If
End Sub

Private Sub Monto2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.Monto2.Text = Format(Val(Me.Monto2.Text), "###0.00")
    
    If (Val(Me.TxtTotalCC.Text) = Val(FrmRegistroComprasList.TxtValorcompra.Text)) Then
        
        cNaturaleza = Trim(Me.TxtNaturaleza.Text)
        ccostos1 = Trim(Me.TxtCostos1.Text)
        cMonto1 = Val(Me.Monto1.Text)
        ccostos2 = Trim(Me.txtCostos2.Text)
        cMonto2 = Val(Me.Monto2.Text)
        ccostos3 = Trim(Me.txtCostos3.Text)
        cMonto3 = Val(Me.Monto3.Text)
        ccostos4 = Trim(Me.txtCostos4.Text)
        cMonto4 = Val(Me.Monto4.Text)
        Unload Me
        FrmRegistroComprasList.cmdAgregar.SetFocus
        Exit Sub
    End If
    Call Resalta(Me.txtCostos3)
End If
End Sub

Private Sub Monto3_Change()
Me.TxtTotalCC.Text = Format(Val(Me.Monto1.Text) + Val(Me.Monto2.Text) + Val(Me.Monto3.Text) + Val(Me.Monto4.Text), "###0.00")
If Val(Me.TxtTotalCC.Text) > Val(FrmRegistroComprasList.TxtValorcompra.Text) Then
    MsgBox "Monto Ingresado Incorrecto", vbInformation, "Mensaje para el Usuario"
    Call Resalta(Me.Monto3)
End If
End Sub

Private Sub Monto3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.Monto3.Text = Format(Val(Me.Monto3.Text), "###0.00")
    If (Val(Me.TxtTotalCC.Text) = Val(FrmRegistroComprasList.TxtValorcompra.Text)) Then
        
        cNaturaleza = Trim(Me.TxtNaturaleza.Text)
        ccostos1 = Trim(Me.TxtCostos1.Text)
        cMonto1 = Val(Me.Monto1.Text)
        ccostos2 = Trim(Me.txtCostos2.Text)
        cMonto2 = Val(Me.Monto2.Text)
        ccostos3 = Trim(Me.txtCostos3.Text)
        cMonto3 = Val(Me.Monto3.Text)
        ccostos4 = Trim(Me.txtCostos4.Text)
        cMonto4 = Val(Me.Monto4.Text)
        Unload Me
        FrmRegistroComprasList.cmdAgregar.SetFocus
        Exit Sub
    End If
    Call Resalta(Me.txtCostos4)
End If
End Sub

Private Sub Monto4_Change()
Me.TxtTotalCC.Text = Format(Val(Me.Monto1.Text) + Val(Me.Monto2.Text) + Val(Me.Monto3.Text) + Val(Me.Monto4.Text), "###0.00")
If Val(Me.TxtTotalCC.Text) > Val(FrmRegistroComprasList.TxtValorcompra.Text) Then
    MsgBox "Monto Ingresado Incorrecto", vbInformation, "Mensaje para el Usuario"
    Call Resalta(Me.Monto4)
End If
End Sub

Private Sub Monto4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.Monto3.Text = Format(Val(Me.Monto3.Text), "###0.00")
    If (Val(Me.TxtTotalCC.Text) = Val(FrmRegistroComprasList.TxtValorcompra.Text)) Then
        cNaturaleza = Trim(Me.TxtNaturaleza.Text)
        ccostos1 = Trim(Me.TxtCostos1.Text)
        cMonto1 = Val(Me.Monto1.Text)
        ccostos2 = Trim(Me.txtCostos2.Text)
        cMonto2 = Val(Me.Monto2.Text)
        ccostos3 = Trim(Me.txtCostos3.Text)
        cMonto3 = Val(Me.Monto3.Text)
        ccostos4 = Trim(Me.txtCostos4.Text)
        cMonto4 = Val(Me.Monto4.Text)
        Unload Me
        FrmRegistroComprasList.cmdAgregar.SetFocus
        Exit Sub
    Else
        MsgBox "Cantidades Incorrectas", vbInformation
        Call Resalta(Me.Monto4)
    End If
    
End If
End Sub

Private Sub TxtNaturaleza_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    busqueda = Buscar2
    FrmPlanContableCuentas.TxtPlanContable.Text = Trim(Me.TxtNaturaleza.Text)
    FrmPlanContableCuentas.Show
    FrmPlanContableCuentas.TxtPlanContable.SetFocus
    Exit Sub
End If
End Sub

Private Sub TxtCostos1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    busqueda = Buscar3
    FrmPlanContableCuentas.TxtPlanContable.Text = Trim(Me.TxtCostos1.Text)
    FrmPlanContableCuentas.Show
    FrmPlanContableCuentas.TxtPlanContable.SetFocus
    Exit Sub
End If
End Sub

Private Sub txtCostos2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    busqueda = Buscar4
    FrmPlanContableCuentas.TxtPlanContable.Text = Trim(Me.txtCostos2.Text)
    FrmPlanContableCuentas.Show
    FrmPlanContableCuentas.TxtPlanContable.SetFocus
    Exit Sub
End If
End Sub

Private Sub txtCostos3_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    busqueda = Buscar5
    FrmPlanContableCuentas.TxtPlanContable.Text = Trim(Me.txtCostos3.Text)
    FrmPlanContableCuentas.Show
    FrmPlanContableCuentas.TxtPlanContable.SetFocus
    Exit Sub
End If
End Sub

Private Sub txtCostos4_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    busqueda = Buscar6
    FrmPlanContableCuentas.TxtPlanContable.Text = Trim(Me.txtCostos4.Text)
    FrmPlanContableCuentas.Show
    FrmPlanContableCuentas.TxtPlanContable.SetFocus
    Exit Sub
End If
End Sub

Private Sub TxtNaturaleza_Change()
If Trim(Me.TxtNaturaleza.Text) <> "" Then
    strCadena = "SELECT * FROM plan_contable_det WHERE pc_codigo='" & Trim(Me.TxtNaturaleza.Text) & "' AND id_plancontable='0001'"
    Call ConfiguraRst(strCadena)
    If rst.RecordCount > 0 Then
        Me.LblDescripcion2.Caption = rst("plan_des")
    Else
        Me.LblDescripcion2.Caption = ""
    End If
    Set rst = Nothing
End If
End Sub
Public Sub Resalta(ByVal texto As TextBox)
texto.SelStart = 0
texto.SelLength = Len(Trim(texto))
texto.Text = texto.SelText
texto.SetFocus
End Sub
