VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "C128"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'variables locales para almacenar los valores de las propiedades
Private mvarC?digo As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarValor As Integer 'copia local
Private mvarPatr?n As String 'copia local
Public Property Let Patr?n(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignaci?n.
'Syntax: X.Patr?n = 5
    mvarPatr?n = vData
End Property


Public Property Get Patr?n() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignaci?n.
'Syntax: Debug.Print X.Patr?n
    Patr?n = mvarPatr?n
End Property



Public Property Let Valor(ByVal vData As Integer)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignaci?n.
'Syntax: X.Valor = 5
    mvarValor = vData
End Property


Public Property Get Valor() As Integer
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignaci?n.
'Syntax: Debug.Print X.Valor
    Valor = mvarValor
End Property



Public Property Let C?digo(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignaci?n.
'Syntax: X.C?digo = 5
    mvarC?digo = vData
End Property


Public Property Get C?digo() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignaci?n.
'Syntax: Debug.Print X.C?digo
    C?digo = mvarC?digo
End Property



