<%@language="vbscript" enablesessionstate="false" lcid="1033"%>
<!-- #include file="incupload.asp" -->
<%
Response.CacheControl = "no-cache"
Response.Expires = -1

progressID = Request.QueryString("ProgressID")

Set upload = New DoteyUpload
Set Info = upload.GetProgressInfo(progressID)

If Info.ReadyState = "loaded" Then
	response.Write("<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">" & vbcrlf)
	response.Write("�ļ��ϴ�����<script language='javascript'>window.close();</script>")
Else
%>
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
	<meta http-equiv="refresh" content="1;url=Progress.asp?ProgressID=<%=ProgressID%>">
	<title>���ϴ�<%= Info.Percentage %></title>
</head>
<body>
<table border="0" cellpadding="0" cellspacing="0">
	<tr><td>�� �� С:</td><td><%= Info.TotalSize %></td></tr>
	<tr><td>�Ѿ��ϴ�:</td><td><%= Info.SizeCompleted %></td></tr>
	<tr><td>ƽ������:</td><td><%= Info.TransferRate %></td></tr>
	<tr><td>ʹ��ʱ��:</td><td><%= Info.ElapsedTime %></td></tr>
	<tr><td>ʣ��ʱ��:</td><td><%= Info.TimeLeft %></td></tr>
</table>
<table border="1" cellpadding="0" cellspacing="0" width="100%">
	<tr>
		<td>
			<table border="0" cellpadding="0" cellspacing="0" width="100%">
				<tr><td width="<%= Info.Percentage %>" bgcolor="blue"></td><td>&nbsp;</td></tr>
			</table>
		</td>
	</tr>
</table>
</body>
</html>
<%
End If

Set Info = Nothing
Set upload = Nothing
%>