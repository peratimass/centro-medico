Attribute VB_Name = "UpdateBusiness"
Public Sub actualizar_version(ByVal in_ruta As String, ByVal in_version_nueva As String, ByVal in_ruc As String)
Dim imagen As String
Dim str_ruta_img As String
Dim Archivo As String
Dim datos(3) As String

Archivo = App.Path & "\Vitekey Business.exe"
On Error GoTo sit

If VerificarArchivo(Archivo) = True Then
    Kill Archivo
End If
sit:
str_ruta_img = App.Path & "\Vitekey Business.exe"
DownloadFile in_ruta, str_ruta_img


  
  

End Sub

Public Function DownloadFile(Url As String, LocalFilename As String) As Boolean

Dim lngRetVal As Long
lngRetVal = URLDownloadToFile(0, Url, LocalFilename, 0, 0)
If lngRetVal = 0 Then DownloadFile = True
End Function
