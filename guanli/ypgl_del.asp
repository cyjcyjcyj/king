<!--#include file="conn.asp"----->
<% id=cint(trim(request("id")))

 set rsjd=server.CreateObject("adodb.recordset")
 rsjd.open"delete * from jobbook where id="&id,cn,1,3


  response.Redirect "ypgl.asp"%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�ޱ����ĵ�</title>
</head>

<body>

</body>
</html>
