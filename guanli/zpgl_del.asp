<!--#include file="conn.asp"----->
<% id=cint(trim(request("id")))

 set rsjd=server.CreateObject("adodb.recordset")
 rsjd.open"delete * from job where id="&id,cn,1,3


  response.Redirect "zpgl.asp"%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
</head>

<body>

</body>
</html>
