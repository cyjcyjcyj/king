<%@EnableSessionState=False%>
<%
	Response.Expires = -1
	PID = Request("PID")
	TimeO = Request("to")
	Set UploadProgress = Server.CreateObject("Persits.UploadProgress")
	format = "<br><CENTER><b>正在上传，请耐心等待...</b></CENTER><br>%T%t%B3%T 速度：(%S/秒)  估计剩余时间：%R %r%U / %V(%P)%l%t"
	bar_content = UploadProgress.FormatProgress(PID, TimeO, "#00007F", format)
  If "" = bar_content Then
%>
<html>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<TITLE>Upload Finished</TITLE>
<SCRIPT LANGUAGE="JavaScript">
function CloseMe()
{
	window.parent.close();
	return true;
}
</SCRIPT>
</HEAD>
<BODY OnLoad="CloseMe()" BGCOLOR="menu">
</BODY>
</HTML>
<%
  Else    ' Not finished yet
%>
<html>
<HEAD>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta HTTP-EQUIV="Refresh" CONTENT="1;URL=<%=Request.ServerVariables("URL") & "?to=" & TimeO & "&PID=" & PID%>">
<TITLE>Uploading Files...</TITLE>
<style type="text/css">
body,td {font-family:Tahoma; font-size: 8pt }
td.spread {font-size: 6pt; line-height:6pt }
td.brick {font-size:6pt; height:12px}
</style>
</HEAD>
<BODY BGCOLOR="menu" topmargin=0>
<% = bar_content %>
</BODY>
</HTML>
<% End If %>