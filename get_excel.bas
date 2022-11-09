Attribute VB_Name = "get_excel"


'devuelve un objeto Recordset con los datos de la hoja
Public Function Leer_Excel(ByVal PathXls As String, Hoja As String) As ADODB.Recordset

      On Error GoTo ErrorFunction
      Dim rs As ADODB.Recordset
      Set rs = New ADODB.Recordset
      Dim cs As String
    
      rs.CursorLocation = adUseClient
      rs.CursorType = adOpenKeyset
      rs.LockType = adLockBatchOptimistic

      cs = "DRIVER=Microsoft Excel Driver (*.xls);" & "DBQ=" & PathXls
      
      Hoja = "[" & Hoja & "$" & "]"
      
      rs.Open "SELECT * FROM " & Hoja, cs
      Set Leer_Excel = rs
      Set rs = Nothing
      Exit Function
ErrorFunction:
      MsgBox Err.Description, vbCritical
      Err.Clear
End Function

'devuelve un objeto Recordset con los datos del txt

Public Function LeerTxt(Directorio As String) As ADODB.Recordset
      On Error GoTo ErrorFunction
      Dim rs As ADODB.Recordset
      Set rs = New ADODB.Recordset
      Dim cn As ADODB.Connection
      Set cn = New ADODB.Connection
      
      cn.Open "DRIVER={Microsoft Text Driver (*.txt; *.csv)};" & _
                         "DBQ=" & Directorio & ";", "", ""
      rs.Open "select * from [archivo#txt]", cn, adOpenStatic, adLockReadOnly, adCmdText
      Set LeerTxt = rs
      
      Set rs = Nothing
      Set cn = Nothing
      
      Exit Function
ErrorFunction:
      MsgBox Err.Description, vbCritical
      Err.Clear
End Function




