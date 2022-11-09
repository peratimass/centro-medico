Attribute VB_Name = "ModCopyImage"
Public Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Boolean
    hNameMappings As Long
    lpszProgressTitle As String
End Type
  
'Declaración Api SHFileOperation
Public Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" _
                                                (lpFileOp As SHFILEOPSTRUCT) As Long
  
'Constantes
Public Const FO_COPY = &H2
Public Const FOF_ALLOWUNDO = &H40
  
  
' Subrutina que copia el archivo
Public Sub Copiar_Archivo(ByVal Origen As String, ByVal Destino As String)
  
Dim t_Op As SHFILEOPSTRUCT
  
    With t_Op
        .hwnd = 0
        .wFunc = FO_COPY
        .pFrom = Origen & vbNullChar & vbNullChar
        .pTo = Destino & vbNullChar & vbNullChar
        .fFlags = FOF_ALLOWUNDO
    End With
  
    ' Se ejecuta la función Api pasandole la estructura
    SHFileOperation t_Op
      
     'reducir el peso de la imagen
    
End Sub

