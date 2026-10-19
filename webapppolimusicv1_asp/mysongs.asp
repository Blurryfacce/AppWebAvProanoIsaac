<!--#include file="conn.asp"-->
<!--#include file="mysongs_db.asp"-->
<%
' =====================================================================
'  mysongs.asp - Pagina propia de canciones de PoliMusic
'  - Roles: normal (solo escucha) y admin (CRUD de canciones)
'  - Usa Session("Username") y Session("Role") guardadas en login.asp
' =====================================================================
Response.CodePage = 65001
Response.Charset = "utf-8"
Response.Expires = -1

Dim psUser, psRole, psIsAdmin, psToken
psUser = Trim(Session("Username") & "")
If psUser = "" Then
    Response.Redirect "login.asp"
    Response.End
End If
psRole = LCase(Trim(Session("Role") & ""))
psIsAdmin = (psRole = "admin")
If psRole = "" Then psRole = "normal"

If Session("PS_Token") & "" = "" Then
    Randomize
    Session("PS_Token") = Hex(Int(Rnd() * 2147483647)) & Hex(Int(Rnd() * 2147483647))
End If
psToken = Session("PS_Token") & ""

Set psConn = PS_OpenConnection()
If Not psConn Is Nothing Then psConnError = ""

' ---------------------------------------------------------------------
'  Acciones (POST): play para todos, add / edit / delete solo admin
' ---------------------------------------------------------------------
Dim psAction, psId, psName, psFile, psPlays, psErr, psErr2, psCmd, psRs
If UCase(Request.ServerVariables("REQUEST_METHOD")) = "POST" Then
    psAction = LCase(Trim(Request.Form("action") & ""))

    If Request.Form("token") & "" <> psToken Then
        If psAction = "play" Then PS_Json "{""ok"":false}"
        PS_Flash "error", "Your session expired. Please try again."
        PS_Back
    End If

    If psConn Is Nothing Then
        If psAction = "play" Then PS_Json "{""ok"":false}"
        PS_Flash "error", "Database connection error: " & psConnError
        PS_Back
    End If

    ' Contador de reproducciones (cualquier usuario logueado)
    If psAction = "play" Then
        psId = PS_ToLong(Request.Form("id"))
        If psId < 0 Then PS_Json "{""ok"":false}"
        psErr = PS_Run("UPDATE TBL_SONG SET PLAYS = ISNULL(PLAYS, 0) + 1 WHERE ID_SONG = ?", Array(psId))
        If psErr <> "" Then PS_Json "{""ok"":false}"
        psPlays = 0
        Set psRs = PS_Cmd("SELECT PLAYS FROM TBL_SONG WHERE ID_SONG = ?", Array(psId)).Execute
        If Not psRs.EOF Then
            If Not IsNull(psRs(0)) Then psPlays = CLng(psRs(0))
        End If
        psRs.Close
        PS_Json "{""ok"":true,""count"":" & psPlays & "}"
    End If

    ' Control de acceso en el servidor: aunque alguien envie el formulario a mano,
    ' un usuario normal no puede modificar canciones.
    If Not psIsAdmin Then
        PS_Flash "error", "Only the admin role can create, edit or delete songs."
        PS_Back
    End If

    Select Case psAction
        Case "add", "edit"
            psId = PS_ToLong(Request.Form("id"))
            psName = Trim(Request.Form("song_name") & "")
            psFile = Trim(Request.Form("file_path") & "")
            psPlays = PS_ToLong(Request.Form("reproductions"))
            If psPlays < 0 Then psPlays = 0

            If psName = "" Or psFile = "" Then
                PS_Flash "error", "Song name and audio file are required."
            ElseIf Len(psName) > 255 Or Len(psFile) > 255 Then
                PS_Flash "error", "Song name and audio file must be 255 characters or less."
            ElseIf psAction = "add" Then
                psErr = PS_Run("INSERT INTO TBL_SONG (SONG_NAME, SONG_PATH, PLAYS) VALUES (?, ?, 0)", Array(psName, psFile))
                If psErr <> "" Then
                    ' Por si la columna id no es IDENTITY
                    psErr2 = PS_Run("INSERT INTO TBL_SONG (ID_SONG, SONG_NAME, SONG_PATH, PLAYS) " & _
                                    "SELECT ISNULL(MAX(ID_SONG), 0) + 1, ?, ?, 0 FROM TBL_SONG", Array(psName, psFile))
                    If psErr2 = "" Then psErr = ""
                End If
                If psErr = "" Then
                    PS_Flash "ok", "Song """ & psName & """ was added."
                Else
                    PS_Flash "error", "Could not add the song: " & psErr
                End If
            ElseIf psId < 0 Then
                PS_Flash "error", "Invalid song."
            Else
                psErr = PS_Run("UPDATE TBL_SONG SET SONG_NAME = ?, SONG_PATH = ?, PLAYS = ? WHERE ID_SONG = ?", _
                               Array(psName, psFile, psPlays, psId))
                If psErr <> "" Then
                    PS_Flash "error", "Could not update the song: " & psErr
                ElseIf psAffected = 0 Then
                    PS_Flash "error", "The song no longer exists."
                Else
                    PS_Flash "ok", "Song """ & psName & """ was updated."
                End If
            End If

        Case "delete"
            psId = PS_ToLong(Request.Form("id"))
            If psId < 0 Then
                PS_Flash "error", "Invalid song."
            Else
                ' Primero se quita la cancion de las tablas que la referencian
                ' (TBL_PLAYER, playlists, tracklists...) para no romper las llaves foraneas
                psErr = PS_Run("DECLARE @id INT = ?; DECLARE @sql NVARCHAR(MAX) = N''; " & _
                               "SELECT @sql = @sql + N'DELETE FROM ' + QUOTENAME(SCHEMA_NAME(t.schema_id)) + N'.' + QUOTENAME(t.name) + N' WHERE ID_SONG = @id; ' " & _
                               "FROM sys.columns c JOIN sys.tables t ON t.object_id = c.object_id " & _
                               "WHERE c.name = 'ID_SONG' AND t.name <> 'TBL_SONG'; " & _
                               "EXEC sp_executesql @sql, N'@id INT', @id = @id; " & _
                               "DELETE FROM TBL_SONG WHERE ID_SONG = @id;", Array(psId))
                If psErr <> "" Then
                    PS_Flash "error", "Could not delete the song: " & psErr
                Else
                    PS_Flash "ok", "The song was deleted."
                End If
            End If

        Case Else
            PS_Flash "error", "Unknown action."
    End Select
    PS_Back
End If

' ---------------------------------------------------------------------
'  Lectura de canciones
' ---------------------------------------------------------------------
Dim psRows, psCount, psTotalPlays, i, psFlashType, psFlashMsg
psCount = 0
psTotalPlays = 0
If Not psConn Is Nothing Then
    On Error Resume Next
    Set psRs = PS_Cmd("SELECT ID_SONG, SONG_NAME, SONG_PATH, PLAYS FROM TBL_SONG ORDER BY ID_SONG", Empty).Execute
    If Err.Number <> 0 Then
        psConnError = Err.Description
        Err.Clear
    ElseIf Not psRs.EOF Then
        psRows = psRs.GetRows()
        psCount = UBound(psRows, 2) + 1
        For i = 0 To psCount - 1
            If IsNull(psRows(3, i)) Then psRows(3, i) = 0
            psTotalPlays = psTotalPlays + CLng(psRows(3, i))
        Next
    End If
    If IsObject(psRs) Then psRs.Close
    On Error GoTo 0
End If

psFlashType = Session("PS_FlashType") & ""
psFlashMsg = Session("PS_FlashMsg") & ""
Session.Contents.Remove "PS_FlashType"
Session.Contents.Remove "PS_FlashMsg"
%>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>PoliMusic &middot; Song List</title>
    <link rel="preconnect" href="https://fonts.googleapis.com">
    <link href="https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,600;9..144,800&family=Plus+Jakarta+Sans:wght@400;500;700&display=swap" rel="stylesheet">
    <link rel="stylesheet" href="mysongs.css">
</head>
<body data-token="<%= PS_H(psToken) %>" class="<% If psIsAdmin Then Response.Write "is-admin" %>">

<header class="ps-topbar">
    <a class="ps-brand" href="mysongs.asp" aria-label="PoliMusic home">
        <span class="ps-brand-disc" aria-hidden="true"></span>
        <span>Poli<b>Music</b></span>
    </a>
    <div class="ps-user">
        <span class="ps-avatar" aria-hidden="true"><%= PS_H(UCase(Left(psUser, 1))) %></span>
        <span class="ps-welcome">Welcome <strong><%= PS_H(psUser) %></strong></span>
        <span class="ps-role ps-role-<%= PS_H(psRole) %>"><%= PS_H(psRole) %></span>
        <a class="ps-logout" href="logout.asp">
            <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M15 4h3a2 2 0 0 1 2 2v12a2 2 0 0 1-2 2h-3M10 17l5-5-5-5M15 12H3"/></svg>
            Logout
        </a>
    </div>
</header>

<main class="ps-main">
    <section class="ps-hero">
        <div>
            <p class="ps-eyebrow">Your library</p>
            <h1>Song List</h1>
            <p class="ps-lead">This is the list of songs available in PoliMusic. Press play to listen to any of them.</p>
        </div>
        <dl class="ps-stats">
            <div><dt>Songs</dt><dd id="psStatSongs"><%= psCount %></dd></div>
            <div><dt>Total plays</dt><dd id="psStatPlays"><%= psTotalPlays %></dd></div>
        </dl>
    </section>

    <% If psFlashMsg <> "" Then %>
    <div class="ps-flash ps-flash-<%= PS_H(psFlashType) %>" role="status">
        <span><%= PS_H(psFlashMsg) %></span>
        <button type="button" class="ps-flash-close" aria-label="Close">&times;</button>
    </div>
    <% End If %>

    <% If psConn Is Nothing Or psConnError <> "" Then %>
    <div class="ps-flash ps-flash-error" role="alert">
        <span>Database error: <%= PS_H(psConnError) %></span>
    </div>
    <% End If %>

    <section class="ps-card">
        <div class="ps-toolbar">
            <label class="ps-search">
                <svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5"/></svg>
                <input type="search" id="psSearch" placeholder="Search songs..." autocomplete="off">
            </label>
            <% If psIsAdmin Then %>
            <button type="button" class="ps-btn ps-btn-primary" id="psAddBtn">
                <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 5v14M5 12h14"/></svg>
                Add song
            </button>
            <% Else %>
            <span class="ps-readonly">Listening mode &middot; only admins can edit songs</span>
            <% End If %>
        </div>

        <div class="ps-table-wrap">
            <table class="ps-table" id="psSongs">
                <thead>
                    <tr>
                        <th class="ps-col-id">ID</th>
                        <th>Song Name</th>
                        <th class="ps-col-play">Play</th>
                        <th class="ps-col-num">Reproductions</th>
                        <% If psIsAdmin Then %><th class="ps-col-actions">Actions</th><% End If %>
                    </tr>
                </thead>
                <tbody>
                <% For i = 0 To psCount - 1 %>
                    <tr data-id="<%= PS_H(psRows(0, i)) %>"
                        data-name="<%= PS_H(psRows(1, i)) %>"
                        data-file="<%= PS_H(psRows(2, i)) %>"
                        data-src="<%= PS_H(PS_AudioUrl(psRows(2, i))) %>"
                        data-plays="<%= PS_H(psRows(3, i)) %>">
                        <td class="ps-col-id" data-label="ID"><%= PS_H(psRows(0, i)) %></td>
                        <td class="ps-col-name" data-label="Song Name">
                            <span class="ps-eq" aria-hidden="true"><i></i><i></i><i></i></span>
                            <span class="ps-song-title"><%= PS_H(psRows(1, i)) %></span>
                            <span class="ps-song-file"><%= PS_H(psRows(2, i)) %></span>
                        </td>
                        <td class="ps-col-play" data-label="Play">
                            <button type="button" class="ps-play" aria-label="Play <%= PS_H(psRows(1, i)) %>">
                                <svg class="ps-i-play" viewBox="0 0 24 24" aria-hidden="true"><path d="M8 5v14l11-7z"/></svg>
                                <svg class="ps-i-pause" viewBox="0 0 24 24" aria-hidden="true"><path d="M7 5h4v14H7zM13 5h4v14h-4z"/></svg>
                            </button>
                            <noscript><audio controls preload="none" src="<%= PS_H(PS_AudioUrl(psRows(2, i))) %>"></audio></noscript>
                        </td>
                        <td class="ps-col-num" data-label="Reproductions"><span class="ps-plays"><%= PS_H(psRows(3, i)) %></span></td>
                        <% If psIsAdmin Then %>
                        <td class="ps-col-actions" data-label="Actions">
                            <button type="button" class="ps-btn ps-btn-ghost ps-edit">Edit</button>
                            <button type="button" class="ps-btn ps-btn-danger ps-delete">Delete</button>
                        </td>
                        <% End If %>
                    </tr>
                <% Next %>
                    <tr class="ps-empty" id="psEmpty"<% If psCount > 0 Then Response.Write " hidden" %>>
                        <td colspan="<% If psIsAdmin Then Response.Write "5" Else Response.Write "4" %>">
                            <% If psCount > 0 Then %>No songs match your search.<% Else %>There are no songs yet.<% End If %>
                        </td>
                    </tr>
                </tbody>
            </table>
        </div>
    </section>
</main>

<!-- Reproductor fijo -->
<div class="ps-player" id="psPlayer">
    <audio id="psAudio" preload="none"></audio>
    <div class="ps-player-info">
        <span class="ps-player-disc" aria-hidden="true"></span>
        <div>
            <strong id="psNowTitle">Nothing playing</strong>
            <small id="psNowSub">Pick a song from the list</small>
        </div>
    </div>
    <div class="ps-player-controls">
        <button type="button" class="ps-ctrl" id="psPrev" aria-label="Previous">
            <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M6 5h2v14H6zM20 5v14L9 12z"/></svg>
        </button>
        <button type="button" class="ps-ctrl ps-ctrl-main" id="psToggle" aria-label="Play">
            <svg class="ps-i-play" viewBox="0 0 24 24" aria-hidden="true"><path d="M8 5v14l11-7z"/></svg>
            <svg class="ps-i-pause" viewBox="0 0 24 24" aria-hidden="true"><path d="M7 5h4v14H7zM13 5h4v14h-4z"/></svg>
        </button>
        <button type="button" class="ps-ctrl" id="psNext" aria-label="Next">
            <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M16 5h2v14h-2zM4 5v14l11-7z"/></svg>
        </button>
    </div>
    <div class="ps-player-progress">
        <span id="psCur">0:00</span>
        <input type="range" id="psSeek" min="0" max="100" step="0.1" value="0" aria-label="Seek">
        <span id="psDur">0:00</span>
    </div>
    <label class="ps-player-vol">
        <svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 9v6h4l5 4V5L8 9zM16 8.5a5 5 0 0 1 0 7"/></svg>
        <input type="range" id="psVol" min="0" max="1" step="0.05" value="0.8" aria-label="Volume">
    </label>
</div>

<div class="ps-toast" id="psToast" role="status" hidden></div>

<% If psIsAdmin Then %>
<!-- Formularios CRUD (solo se generan para el rol admin) -->
<dialog class="ps-dialog" id="psSongDialog">
    <form method="post" action="mysongs.asp" class="ps-form">
        <input type="hidden" name="token" value="<%= PS_H(psToken) %>">
        <input type="hidden" name="action" id="psFormAction" value="add">
        <input type="hidden" name="id" id="psFormId" value="">
        <h2 id="psFormTitle">Add song</h2>
        <label>Song name
            <input type="text" name="song_name" id="psFormName" maxlength="255" required>
        </label>
        <label>Audio file
            <input type="text" name="file_path" id="psFormFile" maxlength="255" list="psAudioFiles" placeholder="my-song.mp3" required>
            <small>File name inside the <%= PS_H(PS_AudioBase()) %> folder.</small>
        </label>
        <label id="psFormPlaysWrap">Reproductions
            <input type="number" name="reproductions" id="psFormPlays" min="0" value="0">
        </label>
        <div class="ps-form-actions">
            <button type="button" class="ps-btn ps-btn-ghost" data-close>Cancel</button>
            <button type="submit" class="ps-btn ps-btn-primary">Save</button>
        </div>
    </form>
</dialog>
<datalist id="psAudioFiles"><%= PS_AudioOptions() %></datalist>

<dialog class="ps-dialog" id="psDeleteDialog">
    <form method="post" action="mysongs.asp" class="ps-form">
        <input type="hidden" name="token" value="<%= PS_H(psToken) %>">
        <input type="hidden" name="action" value="delete">
        <input type="hidden" name="id" id="psDeleteId" value="">
        <h2>Delete song</h2>
        <p>Are you sure you want to delete <strong id="psDeleteName"></strong>? This cannot be undone.</p>
        <div class="ps-form-actions">
            <button type="button" class="ps-btn ps-btn-ghost" data-close>Cancel</button>
            <button type="submit" class="ps-btn ps-btn-danger">Delete</button>
        </div>
    </form>
</dialog>
<% End If %>

<script src="mysongs.js"></script>
</body>
</html>
<%
If Not psConn Is Nothing Then
    On Error Resume Next
    psConn.Close
    On Error GoTo 0
End If
%>
