<%@ Language="VBScript" CodePage="65001" %>
<% Option Explicit %>
<%
' ============================================================
' logout.asp
' Cierra la sesion del usuario y redirige a login.asp.
' ============================================================

Response.CodePage = 65001
Response.Charset = "UTF-8"

' Abandonar la sesion (destruye todas las variables de sesion)
Session.Abandon

' Redirigir a la pagina de login
Response.Redirect "login.asp"
Response.End
%>
