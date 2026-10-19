<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<!--#include file="conn.asp" -->
<%
' ============================================================
' login.asp
' Pagina de inicio de sesion.
' - Muestra un formulario HTML (POST) con usuario y contrasena.
' - Valida contra TBL_USER usando ADODB.Command con parametros
'   (nunca se concatena texto del usuario en el SQL).
' - Si las credenciales son correctas, guarda en sesion el
'   USERNAME y el USER_TYPE, y redirige a songs.asp.
' - Si son incorrectas, vuelve a mostrar el formulario con un
'   mensaje de error.
' - Todo lo que se muestra en pantalla pasa por Server.HTMLEncode.
' ============================================================

Response.CodePage = 65001
Response.Charset = "UTF-8"

Dim strMensaje   ' Mensaje de error a mostrar al usuario
Dim strUsuario   ' Valor ingresado en el campo usuario (para repoblar el form)

strMensaje = ""
strUsuario = ""

' Si se envio el formulario (POST), procesar el login
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then

    ' Recoger los valores del formulario (sin confianza, se parametrizan)
    strUsuario = Trim(Request.Form("username"))
    Dim strContrasena
    strContrasena = Trim(Request.Form("password"))

    ' Validar que no esten vacios
    If strUsuario = "" Or strContrasena = "" Then
        strMensaje = "Por favor, complete todos los campos."
    Else
        ' --- Validar credenciales contra TBL_USER con ADODB.Command ---
        Dim conn, cmd, rs
        Set conn = GetConnection()

        Set cmd = Server.CreateObject("ADODB.Command")
        Set cmd.ActiveConnection = conn

        ' Consulta parametrizada: evita inyeccion SQL
        cmd.CommandText = _
            "SELECT ID_USER, USERNAME, USER_TYPE " & _
            "FROM TBL_USER " & _
            "WHERE USERNAME = ? AND PASSWORD = ?"
        cmd.CommandType = 1  ' adCmdText

        ' Anadir los parametros (tipo: adVarWChar = 202, longitud maxima 100)
        cmd.Parameters.Append cmd.CreateParameter("p_user", 202, 1, 100, strUsuario)
        cmd.Parameters.Append cmd.CreateParameter("p_pass", 202, 1, 100, strContrasena)

        ' Ejecutar la consulta
        On Error Resume Next
        Set rs = cmd.Execute
        If Err.Number <> 0 Then
            ' Error en la consulta: mensaje amigable
            On Error GoTo 0
            strMensaje = "Error al validar las credenciales. Intente mas tarde."
        Else
            On Error GoTo 0

            ' Verificar si se encontro un registro coincidente
            If Not rs.EOF Then
                ' Credenciales correctas: guardar datos en sesion
                Session("USERNAME")  = rs("USERNAME").Value
                Session("USER_TYPE") = rs("USER_TYPE").Value

                ' Cerrar recordset y conexion
                rs.Close
                Set rs = Nothing
                conn.Close
                Set conn = Nothing

                ' Redirigir al listado de canciones
                Response.Redirect "songs.asp"
                Response.End
            Else
                ' No coincide: credenciales incorrectas
                strMensaje = "Usuario o contrasena incorrectos."
            End If

            ' Cerrar recordset y conexion
            If rs.State <> 0 Then rs.Close
            Set rs = Nothing
        End If

        ' Cerrar conexion
        If conn.State <> 0 Then conn.Close
        Set conn = Nothing
    End If
End If
%>
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>PoliMusic - Login</title>
    <link rel="stylesheet" href="style.css">
</head>
<body>

<div class="login-container">
    <h1>PoliMusic</h1>
    <h2>Iniciar Sesion</h2>

    <% ' Mostrar mensaje de error si existe (escapado con HTMLEncode) %>
    <% If strMensaje <> "" Then %>
        <p class="error-message"><%= Server.HTMLEncode(strMensaje) %></p>
    <% End If %>

    <% ' Formulario de login (metodo POST) %>
    <form method="post" action="login.asp">
        <label for="username">Usuario:</label>
        <input type="text" id="username" name="username"
               value="<%= Server.HTMLEncode(strUsuario) %>"
               required autofocus>

        <label for="password">Contrasena:</label>
        <input type="password" id="password" name="password" required>

        <button type="submit">Entrar</button>
    </form>
</div>

</body>
</html>
