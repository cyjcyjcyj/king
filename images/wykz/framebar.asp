<%@EnableSessionState=False%>
<% Response.Expires = -1 %>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>正在上传............</title>
<style type='text/css'>
td {font-family:Tahoma; font-size: 12px;}
</style>
</head>
<BODY BGCOLOR="menu" scroll="NO" frameborder="NO" status="no" style="border:0px;">
<script language="JavaScript">
<!--
function Stopupload()
{
	if (typeof(window.opener)!="undefined"){
		try{
			window.opener.document.execCommand("stop");
		}
		catch(e){}
	}
	else if(typeof(window.dialogArguments)!="undefined"){
		try{
			window.dialogArguments.document.execCommand("stop");
		}
		catch(e){}
	}
	window.close();
}
//-->
</script>
<IFRAME src="bar.asp?PID=<%= Request("PID") & "&to=" & Request("to") %>" title="Uploading" noresize scrolling=no frameborder=0 framespacing=10 width=369 height=115></IFRAME>
<TABLE BORDER="0" WIDTH="100%" cellpadding="2" cellspacing="0">
  <TR><TD ALIGN="center"><button onclick="Stopupload()" style="font-size:12px;">取消上传</button>
</TD></TR>
</TABLE>
</BODY>
</HTML>
