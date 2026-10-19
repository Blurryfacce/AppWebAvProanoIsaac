<%
' =====================================================================
'  mysongs_db.asp
'  Helpers de base de datos para mysongs.asp.
'  Reutiliza la conexion que ya existe en conn.asp, por eso no hay
'  que volver a configurar la cadena de conexion.
' =====================================================================

' Solo si la deteccion automatica falla: pega aqui la misma cadena de
' conexion que tienes en conn.asp. Si la pagina funciona, dejala vacia.
Dim PS_CONN_STRING
PS_CONN_STRING = ""

Dim psConn, psConnError, psAffected
Set psConn = Nothing
psConnError = ""
psAffected = 0

' Devuelve el objeto ADODB.Connection al que apunta "expr" (variable o funcion), o Nothing
Function PS_TryObject(expr)
    Dim o
    Set PS_TryObject = Nothing
    On Error Resume Next
    Err.Clear
    Set o = Eval(expr)
    If Err.Number = 0 Then
        If TypeName(o) = "Connection" Then Set PS_TryObject = o
    End If
    Err.Clear
    On Error GoTo 0
End Function

' Devuelve el texto de "expr" si parece una cadena de conexion, o ""
Function PS_TryString(expr)
    Dim s
    PS_TryString = ""
    On Error Resume Next
    Err.Clear
    s = Eval(expr)
    If Err.Number = 0 And VarType(s) = vbString Then
        If InStr(1, s, "Provider", 1) > 0 Or InStr(1, s, "Data Source", 1) > 0 _
           Or InStr(1, s, "Server", 1) > 0 Or InStr(1, s, "Driver", 1) > 0 Then
            PS_TryString = s
        End If
    End If
    Err.Clear
    On Error GoTo 0
End Function

Function PS_OpenString(cs)
    Dim c
    Set PS_OpenString = Nothing
    On Error Resume Next
    Err.Clear
    Set c = Server.CreateObject("ADODB.Connection")
    c.Open cs
    If Err.Number = 0 Then
        Set PS_OpenString = c
    Else
        psConnError = Err.Description
    End If
    Err.Clear
    On Error GoTo 0
End Function

' Busca la conexion creada en conn.asp (nombres mas comunes) y la abre si hace falta
Function PS_OpenConnection()
    Dim names, n, o, cs
    Set PS_OpenConnection = Nothing

    If PS_CONN_STRING <> "" Then
        Set PS_OpenConnection = PS_OpenString(PS_CONN_STRING)
        Exit Function
    End If

    names = Array("conn", "objConn", "oConn", "cn", "cnn", "con", "cnx", "dbConn", "conexion", _
                  "objConexion", "connection", "db", "GetConnection()", "OpenConnection()", _
                  "GetConn()", "OpenConn()", "AbrirConexion()", "ObtenerConexion()", "Conectar()")
    For Each n In names
        Set o = PS_TryObject(n)
        If Not o Is Nothing Then
            On Error Resume Next
            Err.Clear
            If o.State = 0 Then o.Open
            If Err.Number = 0 And o.State = 1 Then
                Set PS_OpenConnection = o
                On Error GoTo 0
                Exit Function
            End If
            psConnError = Err.Description
            Err.Clear
            On Error GoTo 0
        End If
    Next

    names = Array("connStr", "strConn", "strConnection", "connString", "connectionString", _
                  "cadenaConexion", "strConexion", "strCadena", "cadena", "sConn", "cs")
    For Each n In names
        cs = PS_TryString(n)
        If cs <> "" Then
            Set PS_OpenConnection = PS_OpenString(cs)
            Exit Function
        End If
    Next

    If psConnError = "" Then psConnError = "No connection found in conn.asp. Set PS_CONN_STRING in mysongs_db.asp."
End Function

' Crea un ADODB.Command parametrizado (evita SQL injection)
Function PS_Cmd(sql, params)
    Dim cmd, i, v, size
    Set cmd = Server.CreateObject("ADODB.Command")
    Set cmd.ActiveConnection = psConn
    cmd.CommandText = sql
    cmd.CommandType = 1
    If IsArray(params) Then
        For i = 0 To UBound(params)
            v = params(i)
            If VarType(v) = vbString Then
                size = Len(v)
                If size < 1 Then size = 1
                cmd.Parameters.Append cmd.CreateParameter("p" & i, 202, 1, size, v)
            Else
                cmd.Parameters.Append cmd.CreateParameter("p" & i, 3, 1, 4, CLng(v))
            End If
        Next
    End If
    Set PS_Cmd = cmd
End Function

' Ejecuta INSERT/UPDATE/DELETE. Devuelve "" si todo bien o el mensaje de error.
Function PS_Run(sql, params)
    Dim cmd, ra
    PS_Run = ""
    psAffected = 0
    On Error Resume Next
    Err.Clear
    Set cmd = PS_Cmd(sql, params)
    cmd.Execute ra, , 128
    If Err.Number <> 0 Then
        PS_Run = Err.Description
    Else
        psAffected = ra
    End If
    Err.Clear
    On Error GoTo 0
End Function

Function PS_ToLong(ByVal v)
    PS_ToLong = -1
    v = Trim(v & "")
    If v <> "" And IsNumeric(v) Then
        If CDbl(v) >= 0 And CDbl(v) <= 2147483647 Then PS_ToLong = CLng(Fix(CDbl(v)))
    End If
End Function

Function PS_H(v)
    PS_H = Server.HTMLEncode(v & "")
End Function

' Carpeta de los audios: la primera que exista (en este proyecto es SongFiles)
Dim psAudioBase
psAudioBase = ""

Function PS_AudioBase()
    Dim fso, folders, f
    If psAudioBase = "" Then
        psAudioBase = "SongFiles/"
        On Error Resume Next
        Set fso = Server.CreateObject("Scripting.FileSystemObject")
        folders = Array("SongFiles", "assets/SongFiles", "assets/audio", "audio", "songs")
        For Each f In folders
            If fso.FolderExists(Server.MapPath(f)) Then
                psAudioBase = f & "/"
                Exit For
            End If
        Next
        Err.Clear
        On Error GoTo 0
    End If
    PS_AudioBase = psAudioBase
End Function

' Ruta publica del audio: si file_path ya trae carpeta se usa tal cual,
' si es solo el nombre del archivo se busca en la carpeta de audios
Function PS_AudioUrl(ByVal fp)
    fp = Replace(Trim(fp & ""), "\", "/")
    If InStr(fp, "/") = 0 Then fp = PS_AudioBase() & fp
    PS_AudioUrl = Server.URLPathEncode(fp)
End Function

' Opciones <option> con los archivos de audio disponibles (para el formulario del admin)
Function PS_AudioOptions()
    Dim fso, path, f, ext, html
    html = ""
    On Error Resume Next
    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    path = Server.MapPath(PS_AudioBase())
    If fso.FolderExists(path) Then
        For Each f In fso.GetFolder(path).Files
            ext = LCase(fso.GetExtensionName(f.Name))
            If ext = "mp3" Or ext = "wav" Or ext = "ogg" Or ext = "m4a" Or ext = "aac" Then
                html = html & "<option value=""" & PS_H(f.Name) & """></option>"
            End If
        Next
    End If
    Err.Clear
    On Error GoTo 0
    PS_AudioOptions = html
End Function

Sub PS_Flash(kind, msg)
    Session("PS_FlashType") = kind
    Session("PS_FlashMsg") = msg
End Sub

Sub PS_Back()
    Response.Redirect "mysongs.asp"
    Response.End
End Sub

Sub PS_Json(s)
    Response.Clear
    Response.ContentType = "application/json"
    Response.Write s
    Response.End
End Sub
%>
