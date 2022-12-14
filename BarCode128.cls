VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BarCode128"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'////////////////////////////////////////////////////////////////////
'//
'// Gestionar el BarCode 128
'//
'// Dise?ado por Sergio Fustero Carreras (2006)
'//
'// Crea e imprime c?digo de barras 128 (tipos A y B)
'//     El tipo C est? pendiente de implementar
'// Este c?digo de barras se usa, por ejemplo, en las
'//     cartas certificadas, listas de correos, acuses de
'//     recibo, etc. de Correos (Espa?a)
'//
'// Puede imprimir directamente a impresora o a un
'//     contexto de dispositivo (p.e. PictureBox)
'//
'////////////////////////////////////////////////////////////////////

'variables locales para almacenar los valores de las propiedades
Private mvarTabla As colC128 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarFuente As String 'copia local

Public Enum eTipoDeC?digo128
    cC128_A = 1
    cC128_B = 2
    cC128_C = 3
End Enum

Public Function GenerarBarras(Texto As String, ByVal in_activo As String, ByVal in_dependencia As String, ByVal in_fecha As String, DC As Object, Optional Codificaci?n As eTipoDeC?digo128 = cC128_B, Optional x As Single = -1, Optional y As Single = -1, Optional Densidad As Single = 15, Optional alto As Single = 250, Optional ImprimirTexto As Boolean = False, Optional Fuente As String = "System", Optional Tama?o As Integer = 4) As Boolean
Dim des_almacen As String
'------------------------------
' Genera las barras (secuencias de l?neas negras y espacios en blanco)
'       en c?digo 128 del tipo A, B ? C para un contexto dado.
' El contexto puede ser un PictureBox, p.e.
'
' Par?metros:
'
'   -- Texto:
'       El texto del que se van a generar las barras
'   -- DC:
'       Contexto donde se imprimir?. Puede ser util para los que us?is vista previa.
'       Para imprimir directamente a la impresora, usa la funci?n GenerarBarrasImpresora
'           (puede generar barras mucho m?s peque?as)
'   -- Codificaci?n:
'       Indicar cC128_A, cC128_B ? cC128_C
'   -- X:
'       Coordenada horizontal
'   -- Y:
'       Coordenada vertical
'   -- Densidad:
'       Ancho de las barras (no todos los anchos funcionan bien)
'   -- Alto:
'       Alto de las barras
'   -- ImprimirTexto:
'       Indicamos si queremos que se imprima el texto debajo de las barras o no
'   -- Fuente:
'       Fuente para el texto (si se quiere imprimir)
'   -- Tama?o:
'       Tama?o de fuente para el texto (si se quiere imprimir)
'------------------------------
Dim Aqu? As Single, CBarras As String, c As Integer, D As Integer, paso, cx, Ancho
'DC.CurrentX = Densidad
' -- Establece el ancho de las l?neas
On Error Resume Next
DC.DrawWidth = 1    'Algunos contextos dan error
On Error GoTo 0
Ancho = Densidad
' -- Establecer posici?n vertical (fija durante todo el proceso)
If y = -1 Then
    y = DC.CurrentY
End If
' -- Establecer posici?n horizontal inicial
If x = -1 Then
    Aqu? = DC.CurrentX
Else
    Aqu? = x
End If
' -- Espacio en blanco antes del c?digo de barras
Aqu? = Aqu? + 10 * Densidad
' -- Conseguir el c?digo de barras
CBarras = Barras(Texto, Codificaci?n)
' -- Imprimir el c?digo de barras (se intercalan l?neas y espacios)

DC.CurrentX = 0
DC.Font.name = "Code128"
DC.Font.Size = 8 'Tama?o
    
DC.FontBold = True
DC.Print KEY_EMPRESA
DC.FontBold = False
y = 180
For c = 1 To Len(CBarras)
    ' -- Imprimir linea o espacio en blanco
    
    For D = 1 To Val(Mid(CBarras, c, 1))
        If (c Mod 2) = 0 Then ' 0 = Falso --> L?nea
                                     ' 1 = Verdadero --> Espacio
            DC.Line (Aqu?, y)-(Aqu?, y + alto), vbWhite
        Else
            DC.Line (Aqu?, y)-(Aqu?, y + alto), vbBlack
        End If
        Aqu? = Aqu? + Ancho
    Next
Next
For c = 1 To 10 'Espacio en blanco despu?s
    DC.Line (Aqu?, y)-(Aqu?, y + alto), vbWhite
    Aqu? = Aqu? + Ancho
Next
' -- Imprimir el texto, si procede
If ImprimirTexto Then
    DC.Font.name = "Code128"
    DC.Font.name = "Arial"
    DC.Font.Size = 7 'Tama?o
    DC.FontBold = False
    DC.CurrentY = y + alto
    'DC.CurrentX = Cx + (10 * Densidad) + (Aqu? - Cx - DC.TextWidth(Texto)) / 2
    DC.CurrentX = 0
    DC.Print Texto & Space(10) & "UC:" & get_last_buy(in_dependencia)
    DC.CurrentX = 0
    
    'strCadena = "SELECT descripcion FROM almacen WHERE id_alm='" & in_dependencia & "' and ruc='" & KEY_RUC & "'"
    'Call ConfiguraRstT(strCadena)
    'If rstT.RecordCount > 0 Then
    '    des_almacen = rstT("descripcion")
    'Else
    '    des_almacen = " "
    'End If
    
    'DC.Print Mid(des_almacen, 1, 23)
    'DC.CurrentX = 0
    
    DC.Print Mid(in_activo, 1, 25)
    
End If
End Function


Public Function GenerarBarrasImpresora(Texto As String, Optional Codificaci?n As eTipoDeC?digo128 = cC128_B, Optional x As Single = -1, Optional y As Single = -1, Optional Densidad As Single = 10, Optional alto As Single = 250, Optional ImprimirTexto As Boolean = False, Optional Fuente As String = "System", Optional Tama?o As Integer = 4) As Boolean
'------------------------------
' Genera las barras (secuencias de l?neas negras y espacios en blanco)
'       en c?digo 128 del tipo A, B ? C directamente en la impresora.
'
' Par?metros:
'
'   -- Texto:
'       El texto del que se van a generar las barras
'   -- Codificaci?n:
'       Indicar cC128_A, cC128_B ? cC128_C
'   -- X:
'       Coordenada horizontal
'   -- Y:
'       Coordenada vertical
'   -- Densidad:
'       Ancho de las barras (no todos los anchos funcionan bien)
'   -- Alto:
'       Alto de las barras
'   -- ImprimirTexto:
'       Indicamos si queremos que se imprima el texto debajo de las barras o no
'   -- Fuente:
'       Fuente para el texto (si se quiere imprimir)
'   -- Tama?o:
'       Tama?o de fuente para el texto (si se quiere imprimir)
'------------------------------
Dim Aqu? As Single, CBarras As String, c As Integer, D As Integer, paso, cx, Ancho
' -- Establece el ancho de las l?neas
Printer.DrawWidth = 5
Ancho = Densidad
' -- Establecer posici?n vertical (fija durante todo el proceso)
If y = -1 Then
    y = Printer.CurrentY
End If
' -- Establecer posici?n horizontal inicial
If x = -1 Then
    Aqu? = Printer.CurrentX
Else
    Aqu? = x
End If
' -- Espacio en blanco antes del c?digo de barras
Aqu? = Aqu? + 10 * Densidad
' -- Conseguir el c?digo de barras
CBarras = Barras(Texto, Codificaci?n)
' -- Imprimir el c?digo de barras (se intercalan l?neas y espacios)
For c = 1 To Len(CBarras)
    ' -- Imprimir linea o espacio en blanco
    For D = 1 To Val(Mid(CBarras, c, 1))
        If (c Mod 2) = 0 Then ' 0 = Falso --> L?nea
                                     ' 1 = Verdadero --> Espacio
            Printer.Line (Aqu?, y)-(Aqu?, y + alto), vbWhite
        Else
            Printer.Line (Aqu?, y)-(Aqu?, y + alto), vbBlack
        End If
        Aqu? = Aqu? + Ancho
    Next
Next
' -- Si es necesario, pondremos espacio en blanco despu?s
'For C = 1 To 10 'Espacio en blanco despu?s
'    Printer.Line (Aqu?, Y)-(Aqu?, Y + Alto), vbWhite
'    Aqu? = Aqu? + Ancho
'Next
' -- Imprimir el texto, si procede
If ImprimirTexto Then
    Printer.Font.name = Fuente
    Printer.Font.Size = Tama?o
    Printer.CurrentY = y + alto
    Printer.CurrentX = cx + (10 * Densidad) + (Aqu? - cx - Printer.TextWidth(Texto)) / 2
    Printer.Print Texto
End If
End Function


Public Function Barras(Texto As String, Optional Codificaci?n As eTipoDeC?digo128 = cC128_B) As String
'------------------------------
' Crea la secuencia de barras y espacios para el texto dado
'
' Falta de implementar el c?digo para el tipo C
'       (solamente admite n?meros y es mucho m?s reducido)
'------------------------------
Dim c As Integer, letra As Integer, suma As Long
' -- Valor inicial
Select Case Codificaci?n
    Case eTipoDeC?digo128.cC128_A
        Barras = mvarTabla(103 + 1).Patr?n
        suma = 103
    Case eTipoDeC?digo128.cC128_B
        Barras = mvarTabla(104 + 1).Patr?n
        suma = 104
    Case eTipoDeC?digo128.cC128_C
        Barras = mvarTabla(105 + 1).Patr?n
        suma = 105
End Select
' -- Secuencia del texto
If Codificaci?n = cC128_C Then
Else
    For c = 1 To Len(Texto)
        For letra = 1 To 63
            If mvarTabla(letra).C?digo = Mid(Texto, c, 1) Then
                Barras = Barras & mvarTabla(letra).Patr?n
                suma = suma + mvarTabla(letra).Valor * c
                Exit For
            End If
        Next
        
        If Codificaci?n = cC128_A Then
            For letra = 0 To 31
                If letra = Asc(Mid(Texto, c, 1)) Then
                    Barras = Barras & mvarTabla(letra).Patr?n
                    suma = suma + mvarTabla(letra).Valor * c
                    Exit For
                End If
            Next
        Else
            For letra = 64 To 94
                If mvarTabla(letra).C?digo = Mid(Texto, c, 1) Then
                    Barras = Barras & mvarTabla(letra).Patr?n
                    suma = suma + mvarTabla(letra).Valor * c
                    Exit For
                End If
            Next
        End If
        
        For letra = 95 To 102
            If Asc(Mid(Texto, c, 1)) = letra + 32 Then
                Barras = Barras & mvarTabla(letra).Patr?n
                suma = suma + mvarTabla(letra).Valor * c
                Exit For
            End If
        Next
    Next
End If
' -- C?digo de control
Barras = Barras & mvarTabla((suma Mod 103) + 1).Patr?n
' -- Final
Barras = Barras & mvarTabla(106 + 1).Patr?n
End Function


Public Function PonerFuente(Objeto As Object) As Boolean
'---------------------------
' Asigna la fuente (que se debe haber indicado con anterioridad
'       al objeto pasado como par?metro, si la tiene
'---------------------------
Dim c As Integer
PonerFuente = False
On Error Resume Next
For c = 0 To Objeto.FontCount - 1
    If UCase(Trim(mvarFuente)) = UCase(Trim(Objeto.Fonts(c))) Then
        Objeto.Font.name = mvarFuente
        PonerFuente = True
        Exit For
    End If
Next
If Err.Number <> 0 Then
    PonerFuente = False
End If
On Error GoTo 0
End Function

Public Function AsignarFuente(Fuente As String) As Boolean
Dim c As Integer
'------------------------------
' Asigna la fuente para escribir el dato debajo de las barras
'       (si se quiere)
'------------------------------
AsignarFuente = False
'Recorremos todas las fuentes de pantalla
'       para saber si existe la fuente indicada
For c = 0 To Screen.FontCount - 1
    If UCase(Trim(Screen.Fonts(c))) = UCase(Trim(Fuente)) Then
        AsignarFuente = True
        mvarFuente = Fuente
        Exit For
    End If
Next
End Function

Public Property Get Fuente() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignaci?n.
'Syntax: Debug.Print X.Fuente
    Fuente = mvarFuente
End Property

Public Function Texto(Frase As String, Optional Clase As eTipoDeC?digo128 = cC128_A) As String
'*******************************************************
'Devuelve el texto correspondiente para usarlo
'   con la fuente de c?digo de barras
'
'*******************************************************
Dim suma As Long, c As Integer, V As Integer, Resto As Integer
' -- Valor de inicio, seg?n tipo
Select Case Clase
    Case cC128_A
        suma = 203
    Case cC128_B
        suma = 204
    Case cC128_C
        suma = 205
End Select
Texto = Chr(suma)
' -- C?lculo de la suma a codificar
If Clase = cC128_C Then
    For c = 1 To Len(Frase) Step 2
        V = Val(Mid(Frase, c, 2))
        Texto = Texto & mvarTabla(V).C?digo
        suma = suma + V * c
    Next
Else
    For c = 1 To Len(Frase)
        Select Case Clase
            Case eTipoDeC?digo128.cC128_A
                V = Valor(Mid(Frase, c, 1))
                If V < 64 Then
                    Texto = Texto & Mid(Frase, c, 1)
                End If
            Case eTipoDeC?digo128.cC128_B
                V = Valor(Mid(Frase, c, 1))
                If V < 95 Then
                    Texto = Texto & Mid(Frase, c, 1)
                End If
            Case eTipoDeC?digo128.cC128_C
        End Select
        suma = suma + V * c
    Next
End If
' -- D?gito de control
Resto = suma Mod 103
Texto = Texto + Chr(Resto)
' -- C?digo final
Texto = Texto & Chr(106)
End Function

Private Function Patr?n(Texto As String) As String
'---------------------------
' Devuelve el patr?n para un car?cter dado
'---------------------------
Dim V As C128
Patr?n = ""
For Each V In mvarTabla
    If V.C?digo = Texto Then
        Patr?n = V.Patr?n
        Exit For
    End If
Next
End Function

Private Function Valor(Texto As String) As Integer
'---------------------------
' Devuelve el valor num?rico correspondiente
'       a un car?cter dado
'---------------------------
Dim V As C128
Valor = -1
For Each V In mvarTabla
    If V.C?digo = Texto Then
        Valor = V.Valor
        Exit For
    End If
Next
End Function

Public Sub InicializarTabla()
'---------------------------
' Crea la tabla con los patrones para cada s?mbolo
'
' El patr?n indica, alternativamente, l?neas y espacios
'---------------------------
Dim c As C128
Set c = mvarTabla.Add("A1")
c.Valor = 0
c.C?digo = " "
c.Patr?n = "212222"
Set c = mvarTabla.Add("A2")
c.Valor = 1
c.C?digo = "!"
c.Patr?n = "222122"
Set c = mvarTabla.Add("A3")
c.Valor = 2
c.C?digo = Chr(34)  ' Comillas dobles
c.Patr?n = "222221"
Set c = mvarTabla.Add("A4")
c.Valor = 3
c.C?digo = "#"
c.Patr?n = "121223"
Set c = mvarTabla.Add("A5")
c.Valor = 4
c.C?digo = "$"
c.Patr?n = "121322"
Set c = mvarTabla.Add("A6")
c.Valor = 5
c.C?digo = "%"
c.Patr?n = "131222"
Set c = mvarTabla.Add("A7")
c.Valor = 6
c.C?digo = "&"
c.Patr?n = "122213"
Set c = mvarTabla.Add("A8")
c.Valor = 7
c.C?digo = "'"
c.Patr?n = "122312"
Set c = mvarTabla.Add("A9")
c.Valor = 8
c.C?digo = "("
c.Patr?n = "132212"
Set c = mvarTabla.Add("A10")
c.Valor = 9
c.C?digo = ")"
c.Patr?n = "221213"
Set c = mvarTabla.Add("A11")
'Set C = New C128
c.Valor = 10
c.C?digo = "*"
c.Patr?n = "221312"
Set c = mvarTabla.Add("A12")
c.Valor = 11
c.C?digo = "+"
c.Patr?n = "231212"
Set c = mvarTabla.Add("A13")
c.Valor = 12
c.C?digo = ","
c.Patr?n = "112232"
Set c = mvarTabla.Add("A14")
c.Valor = 13
c.C?digo = "-"
c.Patr?n = "122132"
Set c = mvarTabla.Add("A15")
c.Valor = 14
c.C?digo = "."
c.Patr?n = "122231"
Set c = mvarTabla.Add("A16")
c.Valor = 15
c.C?digo = "/"
c.Patr?n = "113222"
Set c = mvarTabla.Add("A17")
c.Valor = 16
c.C?digo = "0"
c.Patr?n = "123122"
Set c = mvarTabla.Add("A18")
c.Valor = 17
c.C?digo = "1"
c.Patr?n = "123221"
Set c = mvarTabla.Add("A19")
c.Valor = 18
c.C?digo = "2"
c.Patr?n = "223211"
Set c = mvarTabla.Add("A20")
c.Valor = 19
c.C?digo = "3"
c.Patr?n = "221132"
Set c = mvarTabla.Add("A21")
'Set C = New C128
c.Valor = 20
c.C?digo = "4"
c.Patr?n = "221231"
Set c = mvarTabla.Add("A22")
c.Valor = 21
c.C?digo = "5"
c.Patr?n = "213212"
Set c = mvarTabla.Add("A23")
c.Valor = 22
c.C?digo = "6"
c.Patr?n = "223112"
Set c = mvarTabla.Add("A24")
c.Valor = 23
c.C?digo = "7"
c.Patr?n = "312131"
Set c = mvarTabla.Add("A25")
c.Valor = 24
c.C?digo = "8"
c.Patr?n = "311222"
Set c = mvarTabla.Add("A26")
c.Valor = 25
c.C?digo = "9"
c.Patr?n = "321122"
Set c = mvarTabla.Add("A27")
c.Valor = 26
c.C?digo = ":"
c.Patr?n = "321221"
Set c = mvarTabla.Add("A28")
c.Valor = 27
c.C?digo = ";"
c.Patr?n = "312212"
Set c = mvarTabla.Add("A29")
c.Valor = 28
c.C?digo = "<"
c.Patr?n = "322112"
Set c = mvarTabla.Add("A30")
c.Valor = 29
c.C?digo = "="
c.Patr?n = "322211"
Set c = mvarTabla.Add("A31")
'Set C = New C128
c.Valor = 30
c.C?digo = ">"
c.Patr?n = "212123"
Set c = mvarTabla.Add("A32")
c.Valor = 31
c.C?digo = "?"
c.Patr?n = "212321"
Set c = mvarTabla.Add("A33")
c.Valor = 32
c.C?digo = "@"
c.Patr?n = "232121"
Set c = mvarTabla.Add("A34")
c.Valor = 33
c.C?digo = "A"
c.Patr?n = "111323"
Set c = mvarTabla.Add("A35")
c.Valor = 34
c.C?digo = "B"
c.Patr?n = "131123"
Set c = mvarTabla.Add("A36")
c.Valor = 35
c.C?digo = "C"
c.Patr?n = "131321"
Set c = mvarTabla.Add("A37")
c.Valor = 36
c.C?digo = "D"
c.Patr?n = "112313"
Set c = mvarTabla.Add("A38")
c.Valor = 37
c.C?digo = "E"
c.Patr?n = "132113"
Set c = mvarTabla.Add("A39")
c.Valor = 38
c.C?digo = "F"
c.Patr?n = "132311"
Set c = mvarTabla.Add("A40")
c.Valor = 39
c.C?digo = "G"
c.Patr?n = "211313"
Set c = mvarTabla.Add("A41")
c.Valor = 40
c.C?digo = "H"
c.Patr?n = "231113"
Set c = mvarTabla.Add("A42")
c.Valor = 41
c.C?digo = "I"
c.Patr?n = "231311"
Set c = mvarTabla.Add("A43")
c.Valor = 42
c.C?digo = "J"
c.Patr?n = "112133"
Set c = mvarTabla.Add("A44")
c.Valor = 43
c.C?digo = "K"
c.Patr?n = "112331"
Set c = mvarTabla.Add("A45")
c.Valor = 44
c.C?digo = "L"
c.Patr?n = "132131"
Set c = mvarTabla.Add("A46")
c.Valor = 45
c.C?digo = "M"
c.Patr?n = "113123"
Set c = mvarTabla.Add("A47")
c.Valor = 46
c.C?digo = "N"
c.Patr?n = "113321"
Set c = mvarTabla.Add("A48")
c.Valor = 47
c.C?digo = "O"
c.Patr?n = "133121"
Set c = mvarTabla.Add("A49")
c.Valor = 48
c.C?digo = "P"
c.Patr?n = "313121"
Set c = mvarTabla.Add("A50")
c.Valor = 49
c.C?digo = "Q"
c.Patr?n = "211331"
Set c = mvarTabla.Add("A51")
c.Valor = 50
c.C?digo = "R"
c.Patr?n = "231131"
Set c = mvarTabla.Add("A52")
c.Valor = 51
c.C?digo = "S"
c.Patr?n = "213113"
Set c = mvarTabla.Add("A53")
c.Valor = 52
c.C?digo = "T"
c.Patr?n = "213311"
Set c = mvarTabla.Add("A54")
c.Valor = 53
c.C?digo = "U"
c.Patr?n = "213131"
Set c = mvarTabla.Add("A55")
c.Valor = 54
c.C?digo = "V"
c.Patr?n = "311123"
Set c = mvarTabla.Add("A56")
c.Valor = 55
c.C?digo = "W"
c.Patr?n = "311321"
Set c = mvarTabla.Add("A57")
c.Valor = 56
c.C?digo = "X"
c.Patr?n = "331121"
Set c = mvarTabla.Add("A58")
c.Valor = 57
c.C?digo = "Y"
c.Patr?n = "312113"
Set c = mvarTabla.Add("A59")
c.Valor = 58
c.C?digo = "Z"
c.Patr?n = "312311"
Set c = mvarTabla.Add("A60")
c.Valor = 59
c.C?digo = "["
c.Patr?n = "332111"
Set c = mvarTabla.Add("A61")
c.Valor = 60
c.C?digo = "\"
c.Patr?n = "314111"
Set c = mvarTabla.Add("A62")
c.Valor = 61
c.C?digo = "]"
c.Patr?n = "221411"
Set c = mvarTabla.Add("A63")
c.Valor = 62
c.C?digo = "^"
c.Patr?n = "431111"
Set c = mvarTabla.Add("A64")
c.Valor = 63
c.C?digo = "_"
c.Patr?n = "111224"
Set c = mvarTabla.Add("A65")
c.Valor = 64
c.C?digo = "'"
c.Patr?n = "111422"
Set c = mvarTabla.Add("A66")
c.Valor = 65
c.C?digo = "a"
c.Patr?n = "121124"
Set c = mvarTabla.Add("A67")
c.Valor = 66
c.C?digo = "b"
c.Patr?n = "121421"
Set c = mvarTabla.Add("A68")
c.Valor = 67
c.C?digo = "c"
c.Patr?n = "141122"
Set c = mvarTabla.Add("A69")
c.Valor = 68
c.C?digo = "d"
c.Patr?n = "141221"
Set c = mvarTabla.Add("A70")
c.Valor = 69
c.C?digo = "e"
c.Patr?n = "112214"
Set c = mvarTabla.Add("A71")
c.Valor = 70
c.C?digo = "f"
c.Patr?n = "112412"
Set c = mvarTabla.Add("A72")
c.Valor = 71
c.C?digo = "g"
c.Patr?n = "122114"
Set c = mvarTabla.Add("A73")
c.Valor = 72
c.C?digo = "h"
c.Patr?n = "122411"
Set c = mvarTabla.Add("A74")
c.Valor = 73
c.C?digo = "i"
c.Patr?n = "142112"
Set c = mvarTabla.Add("A75")
c.Valor = 74
c.C?digo = "j"
c.Patr?n = "142211"
Set c = mvarTabla.Add("A76")
c.Valor = 75
c.C?digo = "k"
c.Patr?n = "241211"
Set c = mvarTabla.Add("A77")
c.Valor = 76
c.C?digo = "l"
c.Patr?n = "221114"
Set c = mvarTabla.Add("A78")
c.Valor = 77
c.C?digo = "m"
c.Patr?n = "413111"
Set c = mvarTabla.Add("A79")
c.Valor = 78
c.C?digo = "n"
c.Patr?n = "241112"
Set c = mvarTabla.Add("A80")
c.Valor = 79
c.C?digo = "o"
c.Patr?n = "134111"
Set c = mvarTabla.Add("A81")
c.Valor = 80
c.C?digo = "p"
c.Patr?n = "111242"
Set c = mvarTabla.Add("A82")
c.Valor = 81
c.C?digo = "q"
c.Patr?n = "121142"
Set c = mvarTabla.Add("A83")
c.Valor = 82
c.C?digo = "r"
c.Patr?n = "121241"
Set c = mvarTabla.Add("A84")
c.Valor = 83
c.C?digo = "s"
c.Patr?n = "114212"
Set c = mvarTabla.Add("A85")
c.Valor = 84
c.C?digo = "t"
c.Patr?n = "124112"
Set c = mvarTabla.Add("A86")
c.Valor = 85
c.C?digo = "u"
c.Patr?n = "124211"
Set c = mvarTabla.Add("A87")
c.Valor = 86
c.C?digo = "v"
c.Patr?n = "411212"
Set c = mvarTabla.Add("A88")
c.Valor = 87
c.C?digo = "w"
c.Patr?n = "421112"
Set c = mvarTabla.Add("A89")
c.Valor = 88
c.C?digo = "x"
c.Patr?n = "421211"
Set c = mvarTabla.Add("A90")
c.Valor = 89
c.C?digo = "y"
c.Patr?n = "212141"
Set c = mvarTabla.Add("A91")
c.Valor = 90
c.C?digo = "z"
c.Patr?n = "214121"
Set c = mvarTabla.Add("A92")
c.Valor = 91
c.C?digo = "{"
c.Patr?n = "412121"
Set c = mvarTabla.Add("A93")
c.Valor = 92
c.C?digo = "|"
c.Patr?n = "111143"
Set c = mvarTabla.Add("A94")
c.Valor = 93
c.C?digo = "}"
c.Patr?n = "111341"
Set c = mvarTabla.Add("A95")
c.Valor = 94
c.C?digo = "~"
c.Patr?n = "131141"
Set c = mvarTabla.Add("A96")
c.Valor = 95
c.C?digo = ""
c.Patr?n = "114113"
Set c = mvarTabla.Add("A97")
c.Valor = 96
c.C?digo = ""
c.Patr?n = "114311"
Set c = mvarTabla.Add("A98")
c.Valor = 97
c.C?digo = ""
c.Patr?n = "411113"
Set c = mvarTabla.Add("A99")
c.Valor = 98
c.C?digo = ""
c.Patr?n = "411311"
Set c = mvarTabla.Add("A100")
c.Valor = 99
c.C?digo = ""
c.Patr?n = "113141"
Set c = mvarTabla.Add("A101")
c.Valor = 100
c.C?digo = ""
c.Patr?n = "114131"
Set c = mvarTabla.Add("A102")
c.Valor = 101
c.C?digo = ""
c.Patr?n = "311141"
Set c = mvarTabla.Add("A103")
c.Valor = 102
c.C?digo = ""
c.Patr?n = "411131"
Set c = mvarTabla.Add("A104")
c.Valor = 103
c.C?digo = ""
c.Patr?n = "211412"
Set c = mvarTabla.Add("A105")
c.Valor = 104
c.C?digo = ""
c.Patr?n = "211214"
Set c = mvarTabla.Add("A106")
c.Valor = 105
c.C?digo = ""
c.Patr?n = "211232"
Set c = mvarTabla.Add("A107")
c.Valor = 106
c.C?digo = ""
c.Patr?n = "2331112"
'Set C = mvarTabla.Add("A108")
End Sub

Public Property Set Tabla(ByVal vData As colC128)
'se usa al asignar un objeto a la propiedad, en la parte izquierda de una instrucci?n Set.
'Syntax: Set x.Tabla = Form1
    Set mvarTabla = vData
End Property

Public Property Get Tabla() As colC128
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignaci?n.
'Syntax: Debug.Print X.Tabla
    Set Tabla = mvarTabla
End Property

Private Sub Class_Initialize()
mvarFuente = "PF_C128.ttf"
Set mvarTabla = New colC128
InicializarTabla
End Sub


