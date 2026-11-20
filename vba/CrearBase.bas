Attribute VB_Name = "CrearBase"
Option Compare Database
Option Explicit

' ==================================================================
' Crea la base de ejemplo completa dentro de Access:
'   1. Las 4 tablas, con su llave primaria
'   2. Importa los CSV de la carpeta \datos
'   3. Crea las 6 consultas a partir de los archivos de \consultas
'
' Uso: guarda una base de Access en blanco en la carpeta del proyecto,
' importa este modulo (Alt+F11 > Archivo > Importar archivo...),
' pon el cursor dentro de CrearTodo y pulsa F5.
' Autor: Elius Mendez
' ==================================================================

Public Sub CrearTodo()
    Dim raiz As String
    raiz = CurrentProject.Path & "\"

    CrearTablas
    ImportarCsv raiz & "datos\puntos_venta.csv", "PuntosVenta"
    ImportarCsv raiz & "datos\direcciones.csv", "Direcciones"
    ImportarCsv raiz & "datos\marcas.csv", "Marcas"
    ImportarCsv raiz & "datos\geolocalizacion.csv", "Geolocalizacion"
    CrearConsultas raiz & "consultas\"

    Application.RefreshDatabaseWindow
    MsgBox "Base lista: 4 tablas y 6 consultas creadas.", vbInformation, "Crear base"
End Sub

Private Sub CrearTablas()
    BorrarTabla "PuntosVenta"
    BorrarTabla "Direcciones"
    BorrarTabla "Marcas"
    BorrarTabla "Geolocalizacion"

    CurrentDb.Execute "CREATE TABLE PuntosVenta (" & _
        "PuntoID LONG CONSTRAINT pk_puntos PRIMARY KEY, Estado TEXT(20), FechaEstado DATETIME, " & _
        "CodigoClase TEXT(5), RazonSocial TEXT(100), Comentarios TEXT(100), PuntoIDAnterior LONG)", dbFailOnError

    CurrentDb.Execute "CREATE TABLE Direcciones (" & _
        "PuntoID LONG CONSTRAINT pk_direcciones PRIMARY KEY, NombreComercial TEXT(80), Red TEXT(10), " & _
        "Subred TEXT(10), Direccion1 TEXT(100), Direccion2 TEXT(100), Ciudad TEXT(50), Provincia TEXT(30), " & _
        "CodigoPostal TEXT(10), Telefono1 TEXT(15), Telefono2 TEXT(15), TipoPunto TEXT(20))", dbFailOnError

    CurrentDb.Execute "CREATE TABLE Marcas (" & _
        "PuntoID LONG CONSTRAINT pk_marcas PRIMARY KEY, MarcaID TEXT(5))", dbFailOnError

    CurrentDb.Execute "CREATE TABLE Geolocalizacion (" & _
        "PuntoID LONG CONSTRAINT pk_geo PRIMARY KEY, Latitud DOUBLE, Longitud DOUBLE, Canal TEXT(10))", dbFailOnError
End Sub

Private Sub BorrarTabla(nombre As String)
    On Error Resume Next
    CurrentDb.TableDefs.Delete nombre
    On Error GoTo 0
End Sub

' Lee el archivo completo. Funciona igual con saltos de linea de Windows (CRLF) o de Linux (LF).
Private Function LeerLineas(ruta As String) As String()
    Dim f As Integer, contenido As String
    f = FreeFile
    Open ruta For Binary Access Read As #f
    contenido = Space$(LOF(f))
    Get #f, , contenido
    Close #f
    contenido = Replace(contenido, vbCr, "")
    LeerLineas = Split(contenido, vbLf)
End Function

Private Sub ImportarCsv(ruta As String, tabla As String)
    Dim lineas() As String, encabezados() As String, valores() As String
    Dim rs As DAO.Recordset, i As Long, j As Long

    lineas = LeerLineas(ruta)
    encabezados = Split(lineas(0), ",")
    Set rs = CurrentDb.OpenRecordset(tabla, dbOpenDynaset)

    For i = 1 To UBound(lineas)
        If Len(Trim$(lineas(i))) > 0 Then
            valores = Split(lineas(i), ",")
            rs.AddNew
            For j = 0 To UBound(encabezados)
                rs.Fields(encabezados(j)).Value = Convertir(valores(j), rs.Fields(encabezados(j)).Type)
            Next j
            rs.Update
        End If
    Next i
    rs.Close
End Sub

' Convierte el texto del CSV al tipo de la columna. Las fechas vienen como aaaa-mm-dd
' y los decimales con punto, asi funciona sin importar la configuracion regional de Windows.
Private Function Convertir(texto As String, tipo As Integer) As Variant
    If Len(texto) = 0 Then
        Convertir = Null
    ElseIf tipo = dbDate Then
        Convertir = DateSerial(CInt(Left$(texto, 4)), CInt(Mid$(texto, 6, 2)), CInt(Mid$(texto, 9, 2)))
    ElseIf tipo = dbDouble Then
        Convertir = Val(texto)
    ElseIf tipo = dbLong Then
        Convertir = CLng(texto)
    Else
        Convertir = texto
    End If
End Function

' Crea una consulta guardada por cada archivo .sql (Access no acepta comentarios "--", se quitan).
Private Sub CrearConsultas(carpeta As String)
    Dim archivo As String, nombre As String, sql As String
    Dim lineas() As String, i As Long

    archivo = Dir(carpeta & "*.sql")
    Do While Len(archivo) > 0
        lineas = LeerLineas(carpeta & archivo)
        sql = ""
        For i = 0 To UBound(lineas)
            If Left$(Trim$(lineas(i)), 2) <> "--" Then sql = sql & lineas(i) & vbCrLf
        Next i

        nombre = "q" & Left$(archivo, Len(archivo) - 4)
        On Error Resume Next
        CurrentDb.QueryDefs.Delete nombre
        On Error GoTo 0
        CurrentDb.CreateQueryDef nombre, sql

        archivo = Dir()
    Loop
End Sub
