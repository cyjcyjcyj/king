<!--#include file="conn.asp"---->
<%
id=cint(trim(request("id")))
set rsxm=server.CreateObject("adodb.recordset")
rsxm.open"delete * from sbcx where id="&id,cn,1,3
response.Redirect "sbcx_list.asp"
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
</head>

<body>

</body>
</html>
