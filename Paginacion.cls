VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
 Option Explicit

'Propiedades
Public PaginaActual As Long
Public Total_Paginas As Long
Public Registros_Por_Pagina As Long
Public sql As String

'Enum para la navegación y para refrescar los datos
Enum E_Navegacion
    PRIMER_PAGINA = 0
    ANTERIOR = 1
    SIGUIENTE = 2
    ULTIMA_PAGINA = 3
    REFRESCAR = 4
End Enum


'Variables locales
Private m_MSFlexGrid As Object
Private m_Accion As E_Navegacion
Private m_Connection As ADODB.Connection

'Evento que muestra el total de registros, la página actual _
 y el total de paginas del recordset ( por ejemplo en un control Label )
Public Event MoveComplet(ByVal PaginaActual As Long, _
                         ByVal TotalPaginas As Long, _
                         ByVal Total_Registros As Long)


'Sub que llena el flexgrid ( recibe como parámetro la conexión ADO)
'***************************************************************************
Private Sub Cargar_FlexGrid()

On Error GoTo ErrSub

Dim rst As New ADODB.Recordset
Dim Registros As Long
Dim Columna As Integer
    
    
    ' Muestra un mensaje de Error por si no hay un Ms Flexgrid asignado
    If m_MSFlexGrid Is Nothing Then
        
        MsgBox "No se especificó ningún contol " & _
               "Ms FlexGrid para mostrar los datos. ", vbCritical
        Exit Sub
        
    End If
    
    ' Muestra un mensaje de Error por si no hay una conexión ADO establecida
    If m_Connection Is Nothing Then
        
        MsgBox "No se especificó ninguna conexión ADODB. " & _
               " Especificar una. ", vbCritical
        Exit Sub
    End If
    
    'Verifica que se especificó el comando sQL
    If sql = vbNullString Then
        ' Si no se especificó , Muestra el error y sale
        MsgBox "No se especificó ninguna consulta SQL", vbCritical
        Exit Sub
    End If
    
    
    If PaginaActual = 0 Then
       PaginaActual = 1
    End If
    
    Total_Paginas = PaginaActual
    
    'Limpia el Flexgrid
    m_MSFlexGrid.Clear
 
    ' Acá le indica al recordset la cantidad de _
     registros que tendrá cada página
    rst.PageSize = Registros_Por_Pagina
    
    ' Llena el Recordset
    Call rst.Open(sql, m_Connection, adOpenKeyset)
    
    ' Para que no de error
    If Total_Paginas <= rst.PageCount Then
        rst.AbsolutePage = Total_Paginas
    End If
    'Asigna el toal de páginas a la variable
    Total_Paginas = rst.PageCount
    
    m_MSFlexGrid.Rows = 1
    m_MSFlexGrid.Cols = rst.Fields.Count
    
    'Recorre las columnas del flexgrid para crear los encabezados
    For Columna = 0 To rst.Fields.Count - 1
        If Not IsNull(rst.Fields(Columna).Name) Then
            m_MSFlexGrid.TextMatrix(0, Columna) = rst.Fields(Columna).Name
        End If
    Next
    
    'Recorre el Recordset para añadir los registros
    Do While Not rst.EOF
    
        Registros = Registros + 1
        m_MSFlexGrid.Rows = m_MSFlexGrid.Rows + 1
        
        For Columna = 0 To rst.Fields.Count - 1
            
            If Not IsNull(rst.Fields(Columna).Value) Then
                m_MSFlexGrid.TextMatrix(Registros, Columna) = _
                                      rst.Fields(Columna).Value
            End If
    
        Next
        'Siguiente registro
        rst.MoveNext
        'Si llegó al límite de registros de la ágina, sale del DO While
        If (Registros >= Registros_Por_Pagina) Or (rst.EOF) Then
            Exit Do
        End If
    
    Loop
    
    ' Ejecuta el evento con los datos de información
    RaiseEvent MoveComplet(PaginaActual, Total_Paginas, rst.RecordCount)
    
    'Cierra y descarga
    On Error Resume Next
        rst.Close
        Set rst = Nothing
    On Error GoTo 0

Exit Sub

'Error
ErrSub:

MsgBox "Número de error: " & Err.Number & vbNewLine & _
       " Descripción del error: " & Err.Description, vbCritical

End Sub



'********************************************
' Sub y funciones
'********************************************

'Navega a la primer página del recordset
Sub FirstPage()
    m_Accion = PRIMER_PAGINA
    
    Call Calcular_Pagina
    Call Cargar_FlexGrid
End Sub
'Navega a la Anterior página del recordset
Sub PreviousPage()
    m_Accion = ANTERIOR
    
    Call Calcular_Pagina
    Call Cargar_FlexGrid

End Sub

'Siguiente página del recordset
Sub NexPage()
    m_Accion = SIGUIENTE
    
    Call Calcular_Pagina
    Call Cargar_FlexGrid

End Sub

' Va a la  última página del recordset
Sub LastPage()
    m_Accion = ULTIMA_PAGINA
    
    Call Calcular_Pagina
    Call Cargar_FlexGrid

End Sub

'Refresca la grilla
Sub Refresh()
    m_Accion = REFRESCAR
    
    Call Calcular_Pagina
    Call Cargar_FlexGrid
End Sub

' Navega al número de página indicado
Sub GoToPage(Pagina As Long)
    PaginaActual = Pagina
    Call Cargar_FlexGrid
End Sub

Private Sub Calcular_Pagina()
    
    ' Verifica la opción de navegación que se eligíó, para _
      establecer el valor de la variable --> PaginaActual
    
    Select Case m_Accion
    
        Case PRIMER_PAGINA
            PaginaActual = 1
        Case ANTERIOR
            If PaginaActual > 1 Then PaginaActual = PaginaActual - 1
        Case SIGUIENTE
            If PaginaActual < Total_Paginas Then PaginaActual = PaginaActual + 1
        Case ULTIMA_PAGINA
            PaginaActual = Total_Paginas
    End Select
    
End Sub


'********************************************
'Propiedades
'********************************************

' Propiedad para indicar la Conexión ADO
Public Property Get Connection() As ADODB.Connection
Connection = m_Connection
End Property

Public Property Set Connection(Conexion_ADODB As ADODB.Connection)
Set m_Connection = Conexion_ADODB
End Property

'Propiedad para indicar el Flexgrid
Public Property Get MSFlexGrid() As Object
    Set MSFlexGrid = m_MSFlexGrid
End Property
Public Property Set MSFlexGrid(v_MSFlexGrid As Object)
    Set m_MSFlexGrid = v_MSFlexGrid
End Property

Public Property Let MSFlexGrid(v_MSFlexGrid As Object)
    Set m_MSFlexGrid = v_MSFlexGrid
End Property

' Unload del módulo
Private Sub Class_Terminate()

If Not m_MSFlexGrid Is Nothing Then
   Set m_MSFlexGrid = Nothing
End If

If Not m_Connection Is Nothing Then
   Set m_Connection = Nothing
   On Error Resume Next
      m_Connection.Close
End If

End Sub

