
<!--#include file="conn.asp" ------->
<%id=cint(trim(request("id")))
  set rshf=server.CreateObject("adodb.recordset")
  rshf.open"select * from hfbook where dy_id="&id,cn,1,3
%>

<html>

<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href=AdminStyle.css rel=stylesheet type=text/css>
<title>留言回复</title>
<style type="text/css">
<!--
body,td,th {
	font-size: 12px;
}
body {
	background-color: #6699CC;
}
-->
</style></head>

<body>
<%if not rshf.eof then%>
<table width="100%" border="0" cellpadding="4" cellspacing="0" bgcolor="#FFFFFF" id="table1">
<form action=hfxiugai.asp?id=<%=id%> method=post>
	<tr bgcolor="#CCCCCC">
		<td colspan="2">
	  <p align="center">留言回复</td>
	</tr>
	<tr>
		<td width="20%">
		<p align="right">回复内容：</td>
		<td width="78%"><textarea rows="14" name="nr" cols="48"><%=rshf("nr")%></textarea>
	  </td>
	</tr>
	<tr>
		<td colspan="2">
		<p align="center"><input type="submit" value="保存数据" name="B1"></td>
	</tr></form>
</table>
<%end if%>
<%if  rshf.eof then%>
<table width="100%" border="0" cellpadding="4" cellspacing="0" bgcolor="#FFFFFF" id="table1">
<form action=hf_add.asp?id=<%=id%> method=post>
	<tr bgcolor="#CCCCCC">
		<td colspan="2">
	  <p align="center">留言回复</td>
	</tr>
	<tr>
		<td width="20%">
		<p align="right">回复内容：</td>
		<td width="78%"><textarea rows="14" name="nr" cols="48"></textarea>
	  </td>
	</tr>
	<tr>
		<td colspan="2">
		<p align="center"><input type="submit" value="保存数据" name="B1"></td>
	</tr></form>
</table>
<%end if%>
</body>

</html>
