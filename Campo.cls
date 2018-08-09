VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Campo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Variables Contenedoras de propiedades
Dim M_ErrNum As Variant
Dim M_ErrDesc As Variant
'------------------------------------------------------------
' Función: bCSA_ValidarCampoNumerico
'          Comprueba si el dato pasado por parámetro es numérico o no.
'          La función devuelve True si el campo es correcto.
'          En caso contrario devuelve False.
'
' Autor: CSA
'------------------------------------------------------------
Function bCSA_ValidarCampoNumerico(vCampoNumerico As Variant) As Boolean

   If IsNumeric(vCampoNumerico) = False Then
        bCSA_ValidarCampoNumerico = False
    Else
        bCSA_ValidarCampoNumerico = True
    End If
 
End Function








'------------------------------------------------------------
' Función: bCSA_ValidarCampoNulo
'          Comprueba si el dato pasado por parametro es nulo o no.
'          La función devuelve True si el campo es correcto.
'          En caso contrario devuelve False.
'
' Autor: CSA
'------------------------------------------------------------
Function bCSA_ValidarCampoNulo(vCampoNulo As Variant) As Boolean

    If Len(vCampoNulo) = 0 Then
        bCSA_ValidarCampoNulo = False
    Else
        bCSA_ValidarCampoNulo = True
    End If

End Function


'------------------------------------------------------------
' Función: bCSA_ValidarCampoNuloNúmerico
'          Comprueba si el dato pasado por parámetro es nulo o no.
'          La función devuelve True si el campo no es nulo.
'          En caso contrario devuelve False.
'
' Autor: CSA
'------------------------------------------------------------
Function bCSA_ValidarCampoNuloNumerico(vCampoNulo As Variant) As Boolean


    If Len(vCampoNulo) = 0 Then
        bCSA_ValidarCampoNuloNumerico = False
    Else
        If vCampoNulo = 0 Then
            bCSA_ValidarCampoNuloNumerico = False
        Else
            bCSA_ValidarCampoNuloNumerico = True
        End If
    End If
 

End Function








'------------------------------------------------------------
' Función: bCSA_ValidarCampoFecha
'          Comprueba si el dato pasado por parámetro es de tipo fecha o no.
'          La función devuelve True si el campo es correcto.
'          En caso contrario devuelve False.
' Autor: Carmen Castañeda
'
'------------------------------------------------------------
Function bCSA_ValidarCampoFecha(vCampoFecha As Variant) As Boolean

    If IsDate(vCampoFecha) Then
        bCSA_ValidarCampoFecha = True
    Else
        bCSA_ValidarCampoFecha = False
    End If

End Function






'------------------------------------------------------------
' Función: bCSA_ValidarCampoNuloFecha
'          Comprueba si el dato pasado por parámetro es Nulo para Fecha
'          es decir si es (__/__/____) o no.
'          La función devuelve True si el campo es no nulo.
'          En caso contrario devuelve False.
'
' Autor: CSA
'------------------------------------------------------------
Public Function bCSA_ValidarCampoNuloFecha(vCampoFecha As Variant) As Boolean
    
    If vCampoFecha = "__/__/____" Or vCampoFecha = "__-__-____" _
    Or vCampoFecha = "__/__/__" Or vCampoFecha = "__-__-__" Then
        bCSA_ValidarCampoNuloFecha = False
    Else
        bCSA_ValidarCampoNuloFecha = True
    End If

End Function





''------------------------------------------------------------
' Función: bCSA_ValidarIntervaloFecha
'          Comprueba si los datos pasados por parámetro
'          estan en orden creciente de fecha o no.
'          La función devuelve True si son correctos.
'          En caso contrario devuelve False.
'
' Autor: CSA
'------------------------------------------------------------
Public Function bCSA_ValidarIntervaloFecha(vCampoFechaInicio As Variant, vCampoFechaFin As Variant) As Boolean
    
    
    If vCampoFechaInicio = "__/__/____" Or vCampoFechaInicio = "__/__/__" Then
        bCSA_ValidarIntervaloFecha = True
    Else
        If vCampoFechaFin <> "__/__/____" And vCampoFechaFin <> "__/__/__" Then
            If CDate(vCampoFechaInicio) > CDate(vCampoFechaFin) Then
                bCSA_ValidarIntervaloFecha = False
            Else
                bCSA_ValidarIntervaloFecha = True
            End If
        Else
            bCSA_ValidarIntervaloFecha = False
        End If
    End If
    
End Function


'------------------------------------------------------------
' Función: sCSA_SustituirCaracteres
'          Devuelve una cadena pasada por parametro después de
'          haber sustituido los caracteres pasados (Old) por otros (New).
'
' Autor: CSA
'------------------------------------------------------------
Function sCSA_SustituirCaracteres(sCadenaAnalizada As String, sOld As String, sNew As String) As String

    
    Dim nPosicion As Integer

    
    'Activar la rutina de control de errores
    On Error GoTo ErrorZone
    
    
    If Not IsNull(sCadenaAnalizada) Then
        nPosicion = InStr(1, sCadenaAnalizada, sOld)
        Do Until nPosicion = 0
            sCadenaAnalizada = Mid(sCadenaAnalizada, 1, nPosicion - 1) & sNew & _
                Mid(sCadenaAnalizada, nPosicion + 1, Len(sCadenaAnalizada) - nPosicion)
            nPosicion = InStr(nPosicion + 1, sCadenaAnalizada, sOld)
        Loop
        
        sCSA_SustituirCaracteres = sCadenaAnalizada
    Else
        sCSA_SustituirCaracteres = ""
    End If
   
   
    Exit Function
    

ErrorZone:
    M_ErrNum = Err.Number
    M_ErrDesc = Err.Description
    sCSA_SustituirCaracteres = sCadenaAnalizada
   
End Function


'------------------------------------------------------------
' Función: sCSA_NumerALetra
'          Devuelve la cantidad pasada por parámetro en letras.
'
' Autor: CSA
'------------------------------------------------------------
Function sCSA_NumeroALetra(nCantidadAnalizada As Long, bCambio As Boolean) As String

    Dim scadena As String
    Dim sGrupo As String
    Dim nContador As Integer
    Dim nCantidad As Integer
    
    Dim mtrUno(9) As String
    Dim mtrDos(9) As String
    Dim mtrTres(9) As String
    Dim mtrCuatro(9) As String

    ' Activar la rutina de control de errores
    On Error GoTo ErrorZone
    
    
    '**Carga de matrices**

    'Matriz de unidades
    If bCambio = True Then
        mtrUno(1) = "UNO"
    Else
        mtrUno(1) = "UNA"
    End If
    mtrUno(2) = "DOS"
    mtrUno(3) = "TRES"
    mtrUno(4) = "CUATRO"
    mtrUno(5) = "CINCO"
    mtrUno(6) = "SEIS"
    mtrUno(7) = "SIETE"
    mtrUno(8) = "OCHO"
    mtrUno(9) = "NUEVE"
    
    
    'Matriz de 11-19
    mtrDos(1) = "ONCE"
    mtrDos(2) = "DOCE"
    mtrDos(3) = "TRECE"
    mtrDos(4) = "CATORCE"
    mtrDos(5) = "QUINCE"
    mtrDos(6) = "DIECISEIS"
    mtrDos(7) = "DIECISIETE"
    mtrDos(8) = "DIECIOCHO"
    mtrDos(9) = "DIECINUEVE"
    
    
    'Matriz de Decenas
    mtrTres(1) = "DIEZ"
    mtrTres(2) = "VEINTI"
    mtrTres(3) = "TREINTA"
    mtrTres(4) = "CUARENTA"
    mtrTres(5) = "CINCUENTA"
    mtrTres(6) = "SESENTA"
    mtrTres(7) = "SETENTA"
    mtrTres(8) = "OCHENTA"
    mtrTres(9) = "NOVENTA"
   
   
    'Matriz de Centenas
    mtrCuatro(1) = "CIENTO"
    If bCambio = True Then
        mtrCuatro(2) = "DOSCIENTOS"
        mtrCuatro(3) = "TRESCIENTOS"
        mtrCuatro(4) = "CUATROCIENTOS"
        mtrCuatro(5) = "QUINIENTOS"
        mtrCuatro(6) = "SEISCIENTOS"
        mtrCuatro(7) = "SETECIENTOS"
        mtrCuatro(8) = "OCHOCIENTOS"
        mtrCuatro(9) = "NOVECIENTOS"
    Else
        mtrCuatro(2) = "DOSCIENTAS"
        mtrCuatro(3) = "TRESCIENTAS"
        mtrCuatro(4) = "CUATROCIENTAS"
        mtrCuatro(5) = "QUINIENTAS"
        mtrCuatro(6) = "SEISCIENTAS"
        mtrCuatro(7) = "SETECIENTAS"
        mtrCuatro(8) = "OCHOCIENTAS"
        mtrCuatro(9) = "NOVECIENTAS"
    End If
    
    
    sCSA_NumeroALetra = ""
    
    If nCantidadAnalizada > 999999999999# Then
        sCSA_NumeroALetra = ""
        Exit Function
    Else
        scadena = Format(nCantidadAnalizada, "@@@@@@@@@@@@")
    End If
    
    
    For nContador = 0 To 3
        sGrupo = Mid(scadena, (nContador * 3) + 1, 3)
        If sGrupo <> "   " Then
            GoSub Conversion
            If nContador = 0 Then
                sCSA_NumeroALetra = sCSA_NumeroALetra & "MIL "
            ElseIf nContador = 1 Then
                If nCantidadAnalizada > 999999 And nCantidadAnalizada < 2000000 Then
                    sCSA_NumeroALetra = "UN MILLON "
                Else
                    sCSA_NumeroALetra = sCSA_NumeroALetra & "MILLONES "
                End If
            ElseIf nContador = 2 Then
                sCSA_NumeroALetra = sCSA_NumeroALetra & "MIL "
            End If
        End If
    Next nContador
    
    
    Exit Function
    

'Rutina de conversión de grupos
Conversion:
    nCantidad = CInt(sGrupo)
    If nCantidad = 100 Then
        sCSA_NumeroALetra = sCSA_NumeroALetra & "CIEN "
    ElseIf nCantidad > 100 Then
        sCSA_NumeroALetra = sCSA_NumeroALetra & mtrCuatro(Mid(sGrupo, 1, 1)) & " "
    End If
    
    If Val(Mid(sGrupo, 2, 2)) = 10 Then
        sCSA_NumeroALetra = sCSA_NumeroALetra & "DIEZ "
    ElseIf Val(Mid(sGrupo, 2, 2)) = 20 Then
        sCSA_NumeroALetra = sCSA_NumeroALetra & "VEINTE "
    ElseIf Val(Mid(sGrupo, 2, 2)) < 20 And Val(Mid(sGrupo, 2, 2)) > 10 Then
        sCSA_NumeroALetra = sCSA_NumeroALetra & mtrDos(Mid(sGrupo, 3, 1)) & " "
    ElseIf Val(Mid(sGrupo, 2, 2)) < 10 Then
        sCSA_NumeroALetra = sCSA_NumeroALetra & mtrUno(Mid(sGrupo, 3, 1)) & " "
    Else
        sCSA_NumeroALetra = sCSA_NumeroALetra & mtrTres(Mid(sGrupo, 2, 1))
        If Val(Mid(sGrupo, 3, 1)) <> 0 Then
            If Val(Mid(sGrupo, 2, 1)) > 2 Then
                sCSA_NumeroALetra = sCSA_NumeroALetra & " Y " & mtrUno(Mid(sGrupo, 3, 1)) & " "
            ElseIf Val(Mid(sGrupo, 2, 1)) = 2 Then
                sCSA_NumeroALetra = sCSA_NumeroALetra & mtrUno(Mid(sGrupo, 3, 1)) & " "
            End If
        End If
    End If
        
    Return
    

'Rutina de errores
ErrorZone:
    M_ErrNum = Err.Number
    M_ErrDesc = Err.Description
    sCSA_NumeroALetra = ""
   
End Function



'------------------------------------------------------------
' Función: sCSA_NumeroAMes
'          Devuelve el mes en letra correspondiente al número
'          pasado por parámetro
'
' Autor: CSA
'------------------------------------------------------------
Function sCSA_NumeroAMes(nCantidadAnalizada As Integer) As String

    
    Dim mtrMes(12) As String

    ' Activar la rutina de control de errores
    On Error GoTo ErrorZone
    
    
    '**Carga de matrices**
    
    'Matriz de meses
    mtrMes(1) = "Enero"
    mtrMes(2) = "Febrero"
    mtrMes(3) = "Marzo"
    mtrMes(4) = "Abril"
    mtrMes(5) = "Mayo"
    mtrMes(6) = "Junio"
    mtrMes(7) = "Julio"
    mtrMes(8) = "Agosto"
    mtrMes(9) = "Septiembre"
    mtrMes(10) = "Octubre"
    mtrMes(11) = "Noviembre"
    mtrMes(12) = "Diciembre"
    
    
    If nCantidadAnalizada < 1 Or nCantidadAnalizada > 12 Then
        sCSA_NumeroAMes = ""
    Else
        sCSA_NumeroAMes = mtrMes(nCantidadAnalizada)
    End If
    
    Exit Function
    

ErrorZone:
    M_ErrNum = Err.Number
    M_ErrDesc = Err.Description
    sCSA_NumeroAMes = ""
   
End Function


'--------------------------------------------------------------
'Property Get VerrDesc
'         Devuelve la Descripcion del Error producido
'
'Autor: CSA
'--------------------------------------------------------------
Public Property Get VerrDesc() As Variant
    
    VerrDesc = M_ErrDesc
    
End Property


'--------------------------------------------------------------
'Property Get VerrNum
'         Devuelve el numero del error producido
'
'Autor: CSA
'--------------------------------------------------------------
Public Property Get VerrNum() As Variant
    
    VerrNum = M_ErrNum

End Property
