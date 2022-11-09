VERSION 5.00
Begin VB.Form FrmTransferencias_detalle 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "DETALLE ITEM"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   5910
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtid_producto 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Left            =   4080
      TabIndex        =   9
      Top             =   720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Txtid_detalle 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
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
      Left            =   4080
      TabIndex        =   8
      Top             =   360
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "SALIR"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   7
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton cmdprocesar 
      Caption         =   "PROCESAR"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox TxtObservacion 
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
      Height          =   645
      Left            =   2280
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1320
      Width           =   3375
   End
   Begin VB.TextBox TxtCantidadRecibida 
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
      Left            =   2280
      TabIndex        =   3
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox TxtCantidadEnviada 
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
      Left            =   2280
      TabIndex        =   2
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "OBSERVACION :"
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
      Left            =   720
      TabIndex        =   4
      Top             =   1560
      Width           =   1200
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      Caption         =   "CANTIDAD RECIBIDA :"
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
      Left            =   195
      TabIndex        =   1
      Top             =   840
      Width           =   1890
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      Caption         =   "CANTIDAD ENVIADA :"
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
      Left            =   195
      TabIndex        =   0
      Top             =   360
      Width           =   1890
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      Height          =   2535
      Left            =   120
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "FrmTransferencias_detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdprocesar_Click()

If Val(Me.TxtCantidadRecibida.Text) >= 0 And Val(FrmTransferencias.TxtId_transferencia.Text) > 0 Then
    strCadena = "UPDATE movimiento_transferencia_temporal SET recibido='" & Val(Me.TxtCantidadRecibida.Text) & "',observacion='" & Trim(Me.txtObservacion.Text) & "' WHERE  id_temporal='" & Val(FrmTransferencias.HfDetalle.TextMatrix(FrmTransferencias.HfDetalle.Row, 0)) & "' AND ruc='" & KEY_RUC & "' AND dni_save='" & KEY_USUARIO & "'"
    CnBd.Execute (strCadena)
    Call FrmTransferencias.Llenar_Temporal(FrmTransferencias.HfDetalle)
    
Else
    If Val(Me.TxtCantidadEnviada.Text) > 0 Then
        
        If get_diferida(FrmTransferencias.txtid_venta.Text) = "no" Then
            strCadena = "UPDATE movimiento_transferencia_temporal SET cantidad='" & Val(Me.TxtCantidadEnviada.Text) & "',total=peso*'" & Val(Me.TxtCantidadEnviada.Text) & "',observacion='" & Trim(Me.txtObservacion.Text) & "' WHERE  id_temporal='" & Val(FrmTransferencias.HfDetalle.TextMatrix(FrmTransferencias.HfDetalle.Row, 0)) & "' AND ruc='" & KEY_RUC & "' AND dni_save='" & KEY_USUARIO & "'"
            CnBd.Execute (strCadena)
            Call FrmTransferencias.Llenar_Temporal(FrmTransferencias.HfDetalle)
        
        Else
        
        If control_stock_general(Trim(Me.txtid_producto.Text), Val(Me.TxtCantidadEnviada.Text), FrmTransferencias.DtcTipoDoc.BoundText) = True Then
        strCadena = "UPDATE movimiento_transferencia_temporal SET cantidad='" & Val(Me.TxtCantidadEnviada.Text) & "',total=peso*'" & Val(Me.TxtCantidadEnviada.Text) & "',observacion='" & Trim(Me.txtObservacion.Text) & "' WHERE  id_temporal='" & Val(FrmTransferencias.HfDetalle.TextMatrix(FrmTransferencias.HfDetalle.Row, 0)) & "' AND ruc='" & KEY_RUC & "' AND dni_save='" & KEY_USUARIO & "'"
        CnBd.Execute (strCadena)
        Call FrmTransferencias.Llenar_Temporal(FrmTransferencias.HfDetalle)
        End If
    
    End If
    End If
End If

Unload Me
Exit Sub



End Sub



Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Call Resalta(Me.TxtCantidadRecibida)
End Sub


Private Sub Form_Load()
CenterForm Me
Me.Top = 2000
Me.Txtid_detalle.Text = FrmTransferencias.HfDetalle.TextMatrix(FrmTransferencias.HfDetalle.Row, 0)
Me.TxtCantidadEnviada.Text = FrmTransferencias.HfDetalle.TextMatrix(FrmTransferencias.HfDetalle.Row, 3)
Me.TxtCantidadRecibida.Text = FrmTransferencias.HfDetalle.TextMatrix(FrmTransferencias.HfDetalle.Row, 4)

Exit Sub
End Sub

Private Sub TxtCantidadRecibida_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Call Resalta(Me.txtObservacion)
End If
End Sub

Private Sub TxtObservacion_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.cmdprocesar.SetFocus
End If
End Sub
