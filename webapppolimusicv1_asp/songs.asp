<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<!--#include file="conn.asp" -->
<%
' ============================================================
' songs.asp
' Pagina principal: lista todas las canciones de TBL_SONG.
' - Si no hay sesion activa, redirige a login.asp.
' - Muestra "Welcome " + el nombre de usuario de la sesion.
' - Muestra un enlace "Logout" a logout.asp.
' - Muestra el subtitulo "Songs list".
' - Muestra una tabla HTML con columnas:
'   ID | Song Name | Play | Reproductions
' - En "Play" incluye un reproductor <audio controls> con la
'   ruta de la cancion (quitando el "~/" inicial).
' - En "Reproductions" muestra el valor de PLAYS.
' - Recorre el recordset con Do While Not rs.EOF.
' - Cierra el recordset y la conexion al terminar.
' ============================================================

Response.CodePage = 65001
Response.Charset = "UTF-8"

' --- Verificar sesion activa ---
If Session("USERNAME") = "" Then
    ' No hay sesion: redirigir al login
    Response.Redirect "login.asp"
    Response.End
End If

Dim strUsuario
strUsuario = Session("USERNAME")

' --- Consultar todas las canciones de TBL_SONG ---
Dim conn, rs
Set conn = GetConnection()

Dim strSQL
strSQL = "SELECT ID_SONG, SONG_NAME, SONG_PATH, PLAYS FROM TBL_SONG ORDER BY ID_SONG"

On Error Resume Next
Set rs = conn.Execute(strSQL)
If Err.Number <> 0 Then
    ' Error en la consulta: mensaje amigable
    On Error GoTo 0
    Response.Write "<html><body style=""font-family:Arial;"">"
    Response.Write "<h2>Error</h2>"
    Response.Write "<p>No se pudieron cargar las canciones. "
    Response.Write "Por favor, intente mas tarde.</p>"
    Response.Write "<p><a href=""logout.asp"">Cerrar sesion</a></p>"
    Response.Write "</body></html>"
    If conn.State <> 0 Then conn.Close
    Set conn = Nothing
    Response.End
End If
On Error GoTo 0

' --- Funcion auxiliar para convertir la ruta ---
' Convierte "~/songFiles/a.mp3" en "songFiles/a.mp3"
Function ConvertirRuta(strRutaOriginal)
    Dim strRuta
    strRuta = strRutaOriginal
    ' Quitar el "~/ " inicial si existe
    If Left(strRuta, 2) = "~/" Then
        strRuta = Mid(strRuta, 3)
    End If
    ' Devolver la ruta escapada para HTML
    ConvertirRuta = Server.HTMLEncode(strRuta)
End Function
%>
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>PoliMusic - Songs</title>
    <link rel="stylesheet" href="style.css">
</head>
<body>

<div class="songs-container">

    <% ' --- Encabezado con el nombre de usuario (escapado) --- %>
    <div class="header">
        <h1>Welcome <%= Server.HTMLEncode(strUsuario) %></h1>
        <a href="logout.asp" class="logout-link">Logout</a>
    </div>

    <% ' --- Subtitulo --- %>
    <h2>Songs list</h2>

    <% ' --- Tabla de canciones --- %>
    <table class="songs-table">
        <thead>
            <tr>
                <th>ID</th>
                <th>Song Name</th>
                <th>Play</th>
                <th>Reproductions</th>
            </tr>
        </thead>
        <tbody>
        <%
        ' Recorrer el recordset con Do While Not rs.EOF
        Do While Not rs.EOF
            Dim intId, strNombre, strRuta, intPlays
            intId     = rs("ID_SONG").Value
            strNombre = rs("SONG_NAME").Value
            strRuta   = rs("SONG_PATH").Value
            intPlays  = rs("PLAYS").Value
        %>
            <tr>
                <td><%= Server.HTMLEncode(CStr(intId)) %></td>
                <td><%= Server.HTMLEncode(strNombre) %></td>
                <td>
                    <% ' Reproductor de audio con la ruta convertida %>
                    <audio controls>
                        <source src="<%= ConvertirRuta(strRuta) %>" type="audio/mpeg">
                        Tu navegador no soporta el elemento de audio.
                    </audio>
                </td>
                <td><%= Server.HTMLEncode(CStr(intPlays)) %></td>
            </tr>
        <%
            ' Avanzar al siguiente registro
            rs.MoveNext
        Loop
        %>
        </tbody>
    </table>

    <% ' Si no habia canciones, mostrar un mensaje %>
    <% If rs Is Nothing Then %>
        <p>No hay canciones disponibles.</p>
    <% ElseIf rs.State = 0 Then %>
        <p>No hay canciones disponibles.</p>
    <% End If %>

</div>

<%
' --- Cerrar el recordset y la conexion al terminar ---
If Not rs Is Nothing Then
    If rs.State <> 0 Then rs.Close
    Set rs = Nothing
End If
If Not conn Is Nothing Then
    If conn.State <> 0 Then conn.Close
    Set conn = Nothing
End If
%>

</body>
</html>
