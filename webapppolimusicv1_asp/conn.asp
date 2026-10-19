<%
' ============================================================
' conn.asp
' Archivo de conexion a la base de datos BDD_PoliMusic (SQL Server)
' Contiene la funcion GetConnection() que abre y devuelve un
' objeto ADODB.Connection listo para usar.
' Se reutiliza desde las demas paginas con:
'   <!--#include file="conn.asp"-->
' ============================================================

' Funcion que crea y abre la conexion a la base de datos
Function GetConnection()
    Dim conn, errMsg
    Set conn = Server.CreateObject("ADODB.Connection")

    ' Cadena de conexion OLE DB a SQL Server local (driver 19)
    conn.ConnectionString = _
        "Provider=MSOLEDBSQL19;" & _
        "Data Source=localhost;" & _
        "Initial Catalog=BDD_PoliMusic;" & _
        "User Id=usr_polimusic;" & _
        "Password=Politecnica1;" & _
        "Use Encryption for Data=Optional;"

    ' Intentar abrir la conexion
    On Error Resume Next
    conn.Open
    If Err.Number <> 0 Then
        errMsg = Err.Description
        Set conn = Nothing
        On Error GoTo 0
        Response.Write "<html><body style=""font-family:Arial;"">"
        Response.Write "<h2>Error de conexion</h2>"
        Response.Write "<p>No se pudo conectar a la base de datos. "
        Response.Write "Por favor, intente mas tarde.</p>"
        Response.Write "<!-- " & Server.HTMLEncode(errMsg) & " -->"
        Response.Write "</body></html>"
        Response.End
    End If
    On Error GoTo 0

    ' Devolver la conexion abierta
    Set GetConnection = conn
End Function
%>