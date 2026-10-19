<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<%
' ============================================================
' default.asp
' Pagina de inicio: redirige segun el estado de la sesion.
' - Si hay sesion iniciada -> songs.asp
' - Si no hay sesion        -> login.asp
' ============================================================

Response.CodePage = 65001
Response.Charset = "UTF-8"

' Verificar si el usuario tiene sesion activa
If Session("USERNAME") <> "" Then
    ' Sesion activa: ir al listado de canciones
    Response.Redirect "songs.asp"
Else
    ' Sin sesion: ir al login
    Response.Redirect "login.asp"
End If

Response.End
%>
