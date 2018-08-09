VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Configuracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


'Propiedades "globales" de configuración de la aplicación
Dim sTmpTituloAplicacion As String
Dim sTmpTituloAplicacionActiva As String
Dim sTmpFormatoFec As String
Dim sTmpSeparadorFec As String
Dim sTmpParametroLIKE As String
Dim sTmpCadenaConexion As String
Dim sTmpNombreCia As String
Dim iTmpIdCia As Integer
Dim sTmpUsuario As String
Dim sTmpMaquina As String
Dim sTmpLenguaje As String
Dim iTmpIdDelegacion As Integer
Dim nTmpLadoCursor As Integer

Dim M_ErrNum As Variant
Dim M_ErrDesc As Variant



'***********************************************
'Definición de Propiedades de Lectura/Escritura
'***********************************************
'---------------------------------------------------------
'Property Let TituloAplicaion()
'       Toma el Título de la Aplicación
'
'Autor: CSA
'---------------------------------------------------------
Public Property Let TituloAplicacion(valor As String)
    
    sTmpTituloAplicacion = valor

End Property
'---------------------------------------------------------
'Property Get TituloAplicacion()
'         Devuelve el Título de la Aplicación
'
'Autor: CSA
'---------------------------------------------------------
Public Property Get TituloAplicacion() As String
    
    TituloAplicacion = sTmpTituloAplicacion
    
End Property
'---------------------------------------------------------
'Property Let FormatoFec()
'         Toma el Formato de Fecha
'
'Autor: CSA
'---------------------------------------------------------
Public Property Let FormatoFec(valor As String)
    
    sTmpFormatoFec = valor

End Property
'---------------------------------------------------------
'Property Get FormatoFec()
'         Devuelve el Formato de Fecha
'
'Autor: CSA
'---------------------------------------------------------
Public Property Get FormatoFec() As String
    
    FormatoFec = sTmpFormatoFec
    
End Property
'---------------------------------------------------------
'Property Let SeparadorFec()
'         Toma el Separador de Fecha
'
'Autor: CSA
'---------------------------------------------------------
Public Property Let SeparadorFec(valor As String)
    
    sTmpSeparadorFec = valor

End Property
'---------------------------------------------------------
'Property Get SeparadorFec()
'         Devuelve el Separador de Fecha
'
'Autor: CSA
'---------------------------------------------------------
Public Property Get SeparadorFec() As String
    
    SeparadorFec = sTmpSeparadorFec
    
End Property
'---------------------------------------------------------
'Property Let ParametroLIKE()
'       Toma el Parametro Like
'
'Autor: CSA
'---------------------------------------------------------
Public Property Let ParametroLIKE(valor As String)
    
    sTmpParametroLIKE = valor

End Property
'---------------------------------------------------------
'Property Get ParametroLIKE()
'         Devuelve el Parametro Like
'
'Autor: CSA
'---------------------------------------------------------
Public Property Get ParametroLIKE() As String
    
    ParametroLIKE = sTmpParametroLIKE
    
End Property
'---------------------------------------------------------
'Property Let CadenaConexion()
'         Toma la Cadena de Conexión
'
'Autor: CSA
'---------------------------------------------------------
Public Property Let CadenaConexion(valor As String)
    
    sTmpCadenaConexion = valor

End Property
'---------------------------------------------------------
'Property Get CadenaConexion()
'         Devuelve la Cadena de conexión
'
'Autor: CSA
'---------------------------------------------------------
Public Property Get CadenaConexion() As String
    
    CadenaConexion = sTmpCadenaConexion
    
End Property
'---------------------------------------------------------
'Property Let NombreCia()
'         Toma el nombre de la Compañía aciva
'
'Autor: CSA
'---------------------------------------------------------
Public Property Let NombreCia(valor As String)
    
    sTmpNombreCia = valor

End Property
'---------------------------------------------------------
'Property Get NombreCia()
'         Devuelve el nombre de la Compañía activa
'
'Autor: CSA
'---------------------------------------------------------
Public Property Get NombreCia() As String
    
    NombreCia = sTmpNombreCia
    
End Property
'---------------------------------------------------------
'Property Let TituloAplicacionActiva()
'         Toma el título de la Aplicación Activa
'
'Autor: CSA
'---------------------------------------------------------
Public Property Let TituloAplicacionActiva(valor As String)
    
    sTmpTituloAplicacionActiva = valor
End Property
'---------------------------------------------------------
'Property Get TituloAplicacionActiva()
'         Devuelve el título de la Aplicación Activa
'
'Autor: CSA
'---------------------------------------------------------
Public Property Get TituloAplicacionActiva() As String
    
    TituloAplicacionActiva = sTmpTituloAplicacionActiva
    
End Property
'---------------------------------------------------------
'Property Let IdCia()
'         Toma el identificador de la Compañia Activa
'
'Autor: CSA
'---------------------------------------------------------
Public Property Let IdCia(valor As Integer)
    
    iTmpIdCia = valor

End Property
'---------------------------------------------------------
'Property Get IdCia()
'         Devuelve el identificador de la aplicación activa
'
'Autor: CSA
'---------------------------------------------------------
Public Property Get IdCia() As Integer
    
    IdCia = iTmpIdCia
    
End Property
'---------------------------------------------------------
'Property Let Usuario()
'         Toma el Usuario que inició el Sistema
'
'Autor: CSA
'---------------------------------------------------------
Public Property Let Usuario(valor As String)
    
    sTmpUsuario = valor

End Property
'---------------------------------------------------------
'Property Get Usuario()
'         Devuelve el Usuario que inició el Sistema
'
'Autor: CSA
'---------------------------------------------------------
Public Property Get Usuario() As String
    
    Usuario = sTmpUsuario
    
End Property
'---------------------------------------------------------
'Property Let Maquina()
'         Toma el Nombre de la máquina en la que "corre" la aplicación
'
'Autor: CSA
'---------------------------------------------------------
Public Property Let Maquina(valor As String)
    
    sTmpMaquina = valor

End Property
'---------------------------------------------------------
'Property Get Maquina()
'         Devuelve el Nombre de la máquina en la que "corre" la aplicación
'
'Autor: CSA
'---------------------------------------------------------
Public Property Get Maquina() As String
    
    Maquina = sTmpMaquina
    
End Property
'---------------------------------------------------------
'Property Let IdDelegacion()
'         Toma el identificador de la delegaciónen la que "corre" la aplicación
'
'Autor: CSA
'---------------------------------------------------------
Public Property Let IdDelegacion(valor As Integer)
    
    iTmpIdDelegacion = valor

End Property
'---------------------------------------------------------
'Property Get IdDelegacion()
'         Devuelve el identificador de la delegaciónen la que "corre" la aplicación
'
'Autor: CSA
'---------------------------------------------------------
Public Property Get IdDelegacion() As Integer
    
    IdDelegacion = iTmpIdDelegacion
    
End Property

'---------------------------------------------------------
'Property Let LadoCursor()
'         Toma el identificador de la Compañia Activa
'
'Autor: CSA
'---------------------------------------------------------
Public Property Let LadoCursor(valor As Integer)
    
    nTmpLadoCursor = valor

End Property
'---------------------------------------------------------
'Property Get LadoCursor()
'         Devuelve el identificador de la aplicación activa
'
'Autor: CSA
'---------------------------------------------------------
Public Property Get LadoCursor() As Integer
    
    LadoCursor = nTmpLadoCursor
    
End Property



'***********************************************
'Definición de Propiedades de Escritura
'***********************************************

'---------------------------------------------------------
'Property Get VerrNum()
'         Devuelve el error (Número) de la clase que se ha producido
'
'Autor: CSA
'---------------------------------------------------------
Public Property Get VerrNum() As Variant
    
    VerrNum = M_ErrNum

End Property
'---------------------------------------------------------
'Property Get VerrDesc()
'         Devuelve el error (Descripción) de la clase que se ha producido
'
'Autor: CSA
'---------------------------------------------------------
Public Property Get VerrDesc() As Variant
    
    VerrDesc = M_ErrDesc
    
End Property

Public Property Get Lenguaje() As String
    Lenguaje = sTmpLenguaje
End Property

Public Property Let Lenguaje(ByVal valor As String)
    sTmpLenguaje = valor
End Property
