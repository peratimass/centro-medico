VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFlex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Autor : Luciano L.E
'Módulo para editar un mHFlex
'''''''''''''''''''''''''''''''''''''

Private Declare Function GetWindowLong Lib "USER32" Alias "GetWindowLongA" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "USER32" Alias "SetWindowLongA" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "USER32" ( _
    ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal wFlags As Long) As Long

Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1


Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_CLIENTEDGE = &H200
Private Const WS_EX_STATICEDGE = &H20000


Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_NOZORDER = &H4
Private Const SWP_DRAWFRAME = SWP_FRAMECHANGED
Private Const SWP_FLAGS = SWP_NOZORDER Or _
                          SWP_NOSIZE Or _
                          SWP_NOMOVE Or _
                          SWP_DRAWFRAME



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


Private WithEvents mHFlex As MSHFlexGrid
Attribute mHFlex.VB_VarHelpID = -1
Private WithEvents mTextBox As TextBox
Attribute mTextBox.VB_VarHelpID = -1
Private WithEvents mDtPicker As DtPicker
Attribute mDtPicker.VB_VarHelpID = -1



Private sTemp As String

Enum EFormato
    estring = 0
    eBoolean = 1
    EDate = 2
    enumber = 3
    ecurrency = 4
End Enum

Private Type tcol
    Bloqueado As Boolean
    Formato As EFormato
End Type

Public Event validar(IRowActual As Long, _
                     IColActual As Integer, _
                     CellValue As Variant, _
                     Cancelar As Boolean)


Enum eEditMode
    eClick = 0
    eDobleClick = 1
End Enum

Private AntCol As Integer
Private AntRow As Long

Private mEditMode As eEditMode


Private mCol(100) As tcol



Function SetColumnas(Indice As Long, _
                     Optional Formato As EFormato, _
                     Optional Bloquear As Boolean = False)
                     
    If mHFlex Is Nothing Then
       MsgBox "No se ha indicado el FlexGrid. Ejecutar la función " & _
              "Iniciar previamente", vbCritical
       Exit Function
    End If
    
    mCol(Indice).Bloqueado = Bloquear
    mCol(Indice).Formato = Formato
    
    If Formato = eBoolean Then
        If Not mHFlex.DataSource Is Nothing Then
            With mHFlex
                 Dim i As Integer
                 Dim sBoolean As String
                 For i = 1 To mHFlex.Rows - 1
                     If .TextMatrix(i, Indice) = True Then
                        .TextMatrix(i, Indice) = "Sí"
                     ElseIf .TextMatrix(i, Indice) = False Then
                        .TextMatrix(i, Indice) = "No"
                     End If
                     
                 Next
        End With
      End If
    End If
    

 
    
End Function

Private Sub mDtPicker_Change()
        With mHFlex
            If mCol(.col).Formato = EDate Then
                .TextMatrix(.Row, .col) = mDtPicker.Value
            End If
        End With
End Sub

Private Sub mDTPicker_KeyUp(KeyCode As Integer, Shift As Integer)
       
      If mDtPicker.Visible Then
          With mHFlex
                Select Case KeyCode
                    Case 13, 9
                        ValidarDtPicker
                End Select
          End With
      End If
End Sub

Sub ValidarDtPicker()
    
    With mHFlex
        Dim Cancelar As Boolean
        
        RaiseEvent validar(AntRow, AntCol, mDtPicker.Value, Cancelar)
        
        If Cancelar Then
           .Row = AntRow
           .col = AntCol
           .Text = sTemp
        End If
        
    End With
    
     mDtPicker.Visible = False
    
End Sub

Private Sub mDtPicker_LostFocus()
    mDTPicker_KeyUp 13, 0
End Sub

Private Sub mHFlex_Click()
If Val(mHFlex.TextMatrix(mHFlex.Row, 0)) > 0 Then
    If mEditMode = eDobleClick Then
       Exit Sub
    ElseIf mEditMode = eClick Then
       Call mHFlex_DblClick
    End If
 End If
End Sub

Private Sub mHFlex_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mHFlex.Row <> mHFlex.RowSel Or mHFlex.col <> mHFlex.ColSel Then
       mHFlex.RowSel = mHFlex.Row
       mHFlex.ColSel = mHFlex.col
    End If
End Sub

Private Sub mtextbox_LostFocus()

Dim Cancelar As Boolean
    
    Dim opcion As EFormato
    
    opcion = mCol(AntCol).Formato
    
    Select Case mCol(AntCol).Formato
        Case estring
            RaiseEvent validar(AntRow, AntCol, mTextBox.Text, Cancelar)
            
            'If Cancelar Then
            '    mHFlex.Text = sTemp
            'End If
        ' números
        Case enumber
            If Not IsNumeric(mTextBox.Text) Then
               RaiseEvent validar(AntRow, AntCol, 0, Cancelar)
               mHFlex.TextMatrix(AntRow, AntCol) = 0
            Else
               RaiseEvent validar(AntRow, AntCol, mTextBox.Text, Cancelar)
            End If
            
        ' currency
        Case ecurrency
            If Not IsNumeric(mTextBox.Text) Then
                RaiseEvent validar(AntRow, AntCol, 0, Cancelar)
                mHFlex.TextMatrix(AntRow, AntCol) = 0
            Else
                RaiseEvent validar(AntRow, AntCol, mTextBox.Text, Cancelar)
            End If
            'If Cancelar Then
            '    mHFlex.Text = sTemp
            'End If
    End Select
    
    If Cancelar Then
        mHFlex.Row = AntRow
        mHFlex.col = AntCol
        mHFlex.Text = sTemp
    End If
    
    mTextBox.Visible = False

End Sub

Private Sub mHFlex_DblClick()

    Dim Cancelar As Boolean
    Dim iCol As Integer
    Dim in_pago As Double
    iCol = mHFlex.col
    
    If mCol(iCol).Bloqueado Then
        Exit Sub
    End If
    
    With mHFlex
        sTemp = .Text
        AntCol = .col
        AntRow = .Row
    End With
    in_pago = mHFlex.Text
    
    If mCol(iCol).Formato = eBoolean Then
        With mHFlex
            Dim sBoolean As String
            
            If LCase(.TextMatrix(.Row, .col)) = LCase("Sí") Then
               RaiseEvent validar(AntRow, AntCol, False, Cancelar)
            ElseIf LCase(.TextMatrix(.Row, .col)) = LCase("No") Then
               RaiseEvent validar(AntRow, AntCol, True, Cancelar)
            End If
            
            If Cancelar = False Then
               .TextMatrix(.Row, .col) = Format(UCase(.TextMatrix(.Row, .col)) = "NO" Or Trim(.TextMatrix(.Row, .col)) = vbNullString, "YES/NO")
            End If
        End With
        Exit Sub
    End If
    
    If mCol(iCol).Formato = EDate Then
        With mHFlex
            If Trim(.Text) <> vbNullString Then
                mDtPicker.Value = Format(.Text, "Short Date")
            Else
                mDtPicker.Value = Format(Date, "Short Date")
            End If
        
            mDtPicker.Move (.Left + .CellLeft - 25), _
                           (.Top + (.RowHeight(0) * (.Row - .TopRow + 1))), _
                           (.CellWidth - 8), _
                           (.CellHeight + 25)
                                
            mDtPicker.Visible = True
            mDtPicker.Enabled = True
            mDtPicker.SetFocus
            mDtPicker.Value = .TextMatrix(.Row, .col)
     
            Exit Sub
        End With
    End If


    mTextBox.Visible = False

    With mTextBox
        .Text = Format(in_pago, "#,##0.00")
        .Top = mHFlex.Top + mHFlex.CellTop - 25
        .Left = mHFlex.Left + mHFlex.CellLeft - 25
        .Width = mHFlex.CellWidth
        .Height = mHFlex.CellHeight
        .Visible = True
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With


End Sub

Private Sub mHFlex_Scroll()
    mDtPicker.Visible = False
    mTextBox.Visible = False

End Sub

Private Sub mHFlex_KeyPress(KeyAscii As Integer)
  
  If KeyAscii = 13 Then
     If mEditMode = eClick Then
        mHFlex_Click
     ElseIf mEditMode = eDobleClick Then
        mHFlex_DblClick
     End If
   End If
End Sub

Private Sub mTextBox_KeyPress(KeyAscii As Integer)
    If mCol(mHFlex.col).Formato = enumber Then
       If InStr("1234567890" & Chr(8), Chr(KeyAscii)) = 0 Then
          KeyAscii = 0
       End If
    End If
    If mCol(mHFlex.col).Formato = ecurrency Then
       If InStr("1234567890" & Chr(8) & Chr(44), Chr(KeyAscii)) = 0 Then
          KeyAscii = 0
       End If
    End If
End Sub



Private Sub mTextBox_Change()
     mHFlex.Text = mTextBox
End Sub

Private Sub mTextBox_KeyDown(KeyCode As Integer, Shift As Integer)
  
  If KeyCode = 13 Then
    mTextBox.Visible = False
    mHFlex.SetFocus
  ElseIf KeyCode = vbKeyEscape Then
    mHFlex.Text = sTemp
    mTextBox.Visible = False
    mHFlex.SetFocus
  End If

End Sub

Sub iniciar(HierarchicalFlex As MSHFlexGrid, FRM As Form, DtPicker As DtPicker)
    
    If mHFlex Is Nothing Then
       
       'Set mAdoCon = AdoConection
       Set mTextBox = FRM.Controls.Add("vb.textbox", "txtEditarCelda")
       HierarchicalFlex.FixedRows = 1
       'If HierarchicalFlex.FixedCols <> 0 Or HierarchicalFlex.FixedRows <> 1 Then
       '   MsgBox "Este ejemplo solo funciona con la propiedad FixedCol en " & _
                 "0 y FixedRow en 1. Establezca esas propiedades para el " & _
                 "Hierarchical FlexGrid. ( hacerlo en tiempo de diseño)", vbCritical
       '   Exit Sub
       'End If
       
       Set mHFlex = HierarchicalFlex
       
       With mHFlex
            .SelectionMode = flexSelectionFree
            Call Flat(.hwnd)
            .FixedCols = 0
       End With
       
       With mTextBox
            .BackColor = &HC0FFC0
            .ZOrder 0
            .BorderStyle = 0
       End With
       
       Set mDtPicker = DtPicker
       DtPicker.Visible = False
       Call Flat(DtPicker.hwnd)
    Else
       Exit Sub
    End If
End Sub

Sub iniciar_ii(HierarchicalFlex As MSHFlexGrid, FRM As Form, DtPicker As DtPicker)
    
    If mHFlex Is Nothing Then
       
       'Set mAdoCon = AdoConection
       Set mTextBox = FRM.Controls.Add("vb.textbox", "txtEditarCeldaii")
       HierarchicalFlex.FixedRows = 1
       'If HierarchicalFlex.FixedCols <> 0 Or HierarchicalFlex.FixedRows <> 1 Then
       '   MsgBox "Este ejemplo solo funciona con la propiedad FixedCol en " & _
                 "0 y FixedRow en 1. Establezca esas propiedades para el " & _
                 "Hierarchical FlexGrid. ( hacerlo en tiempo de diseño)", vbCritical
       '   Exit Sub
       'End If
       
       Set mHFlex = HierarchicalFlex
       
       With mHFlex
            .SelectionMode = flexSelectionFree
            Call Flat(.hwnd)
            .FixedCols = 0
       End With
       
       With mTextBox
            .BackColor = &HC0FFC0
            .ZOrder 0
            .BorderStyle = 0
       End With
       
       Set mDtPicker = DtPicker
       DtPicker.Visible = False
       Call Flat(DtPicker.hwnd)
    Else
       Exit Sub
    End If
End Sub


Private Sub Flat(ByVal hwnd As Long)
    Dim e As Long
        
        e = GetWindowLong(hwnd, GWL_EXSTYLE)
        e = e And Not WS_EX_CLIENTEDGE Or WS_EX_STATICEDGE
        SetWindowLong hwnd, GWL_EXSTYLE, e
        SetWindowPos hwnd, 0, 0, 0, 0, 0, _
                     SWP_NOACTIVATE Or SWP_NOZORDER Or SWP_FRAMECHANGED _
                      Or SWP_NOSIZE Or SWP_NOMOVE
End Sub


Property Get EditMode() As eEditMode
    EditMode = mEditMode
End Property

Property Let EditMode(Valor As eEditMode)
    mEditMode = Valor
End Property



