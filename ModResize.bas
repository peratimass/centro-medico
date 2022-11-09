Attribute VB_Name = "ModResize"
'MODULO DE FORMATO DE TAMAÑO
Public Sub DarFormato1(ByVal Grilla As MSHFlexGrid, ByVal longitud As Integer)
Dim x As Integer
Dim Formato As String
  Formato = ""
  For x = 1 To 4
    Formato = Formato + "0"
  Next x
  Grilla.col = longitud
  Grilla.Row = 0
  
For x = 0 To (Grilla.Rows - 2)
    Grilla.Row = Grilla.Row + 1
    Grilla.Text = Format(Trim(str(Val(Right(Grilla.Text, 4)))), Formato)
Next x

  Grilla.Refresh
  End Sub
Sub Reziseform(ByVal Formulario As Form)
CenterForm Formulario
Formulario.Width = 5505
Formulario.Height = 5970
End Sub
Sub llenarGrid(ByVal Grilla As MSHFlexGrid, ByVal Formulario As Form)
On Error GoTo salir
  Call ConfiguraRst(strCadena)
  Grilla.Clear
  Grilla.Rows = rst.RecordCount + 1
  Set Grilla.Recordset = rst
  Grilla.ColWidth(0) = 900
  Grilla.ColWidth(1) = 4500
'Call DarFormato1(Grilla, 0)
Grilla.Refresh
  Formulario.TlbAcciones.Buttons(KEY_UPDATE).Enabled = False
  Formulario.TlbAcciones.Buttons(KEY_DELETE).Enabled = False
  Exit Sub
salir: MsgBox "Por Favor Digite Correctamente", vbInformation, "Mensaje para el usuario"
End Sub
Sub FormReport(ByVal Formulario As Form)
CenterForm Formulario
Formulario.Width = 8130
Formulario.Height = 4995
End Sub

