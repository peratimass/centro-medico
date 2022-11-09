VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form FrmControlAccesos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de Accesos"
   ClientHeight    =   3525
   ClientLeft      =   2055
   ClientTop       =   1860
   ClientWidth     =   5220
   Icon            =   "FrmControlAccesos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   5220
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HfdGrilla 
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   4260
      _Version        =   393216
      ForeColor       =   8388608
      FixedCols       =   0
      ForeColorFixed  =   8388608
      SelectionMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
End
Attribute VB_Name = "FrmControlAccesos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Procedencia As EnumProcede

Private Sub Form_Activate()
  StrCadena = "SELECT IdUsuario as Código,NombreUsuario as Nombre FROM Seguridad ORDER BY IdUsuario"
  Call ConfiguraRst(StrCadena)
  Set HfdGrilla.Recordset = Rst
  HfdGrilla.ColWidth(1) = 2500
  HfdGrilla.ColWidth(2) = 2000
  
  Set Rst = Nothing
End Sub

Private Sub Form_Load()
CenterForm Me
End Sub

Private Sub Form_Resize()
  On Error Resume Next
  Dim intAuxTop As Integer
  intAuxTop = 100
 ' HfdGrilla.Move 100, intAuxTop, (Me.Width - ClbAcciones.Width - 400), _
  (Me.Height - intAuxTop - 540)
 ' ClbAcciones.Move HfdGrilla.Width + 200, intAuxTop, Me.Width, _
  (Me.Height - intAuxTop - 540)
End Sub

Private Sub TlbAcciones_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Key
    Case KEY_NEW
      Procedencia = Nuevo
      Me.HfdGrilla.Col = 0
      FrmAcceso.intIdUsuario = Me.HfdGrilla
      FrmAcceso.Show
    
      
  End Select
End Sub

