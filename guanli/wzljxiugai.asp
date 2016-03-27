<!--#include file="lianjie.asp"---->
<%
id=cint(trim(request("id")))
set rsx=server.CreateObject("adodb.recordset")
rsx.open"select * from yqlj where id="&id,cn,1,3
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
<style type="text/css">
<!--
body {
	background-color: #7B9AE7;
}
body,td,th {
	font-size: 12px;
}
.gg1 {
	border: 1px solid #FFFFFF;
}
.gg2 {
	border-top-width: 1px;
	border-right-width: 1px;
	border-bottom-width: 1px;
	border-left-width: 1px;
	border-top-style: solid;
	border-right-style: solid;
	border-bottom-style: none;
	border-left-style: none;
	border-top-color: #DBBF90;
	border-right-color: #DBBF90;
	border-bottom-color: #DBBF90;
	border-left-color: #DBBF90;
}
.gg3 {
	border-top-width: 1px;
	border-right-width: 1px;
	border-bottom-width: 1px;
	border-left-width: 1px;
	border-top-style: solid;
	border-right-style: none;
	border-bottom-style: none;
	border-left-style: none;
	border-top-color: #DBBF90;
	border-right-color: #DBBF90;
	border-bottom-color: #DBBF90;
	border-left-color: #DBBF90;
}
.gg4 {
	border: 1px solid #66CCFF;
}
.gg5 {
	border: 1px solid #000000;
	font-size: 12px;
}
.style3 {
	color: #DBBF90;
	font-weight: bold;
}
.style4 {color: #FFFFFF}
a:link {
	color: #0000FF;
	text-decoration: none;
}
a:visited {
	text-decoration: none;
}
a:hover {
	text-decoration: underline;
}
a:active {
	text-decoration: none;
}
-->
</style></head>

<body>
<form name="form" method="post" action="wzlj_xiugai.asp?id=<%=id%>">
  <table width="539" border="0" align="center" cellpadding="0" cellspacing="0" class="gg1">
    <tr align="center" bgcolor="#639AC6">
      <td height="17" colspan="2"><span class="style3">友情连接</span></td>
    </tr>
    <tr>
      <td width="77" height="32" align="center" class="gg2">网站名称：</td>
      <td width="460" align="left" class="gg3">&nbsp; <input name="wzmc" type="text" class="gg4" id="wzmc" value="<%=rsx("wzmc")%>"></td>
    </tr>
    <tr>
      <td height="26" align="center" class="gg2">网站网址：</td>
      <td class="gg3">&nbsp;&nbsp;<input name="wzwz" type="text" class="gg4" id="wzwz" value="<%=rsx("wzwz")%>"></td>
    </tr>
    <tr>
      <td height="25" align="center" class="gg2">链接时间：</td>
      <td class="gg3">&nbsp;&nbsp;<input name="jrsj" type="text" class="gg4" id="jrsj" value="<%=rsx("jrsj")%>"></td>
    </tr>
    <tr>
      <td height="25" colspan="2" align="center" class="gg2">
      <input name="ok" type="submit" class="gg5" id="ok" value="修改" >&nbsp;&nbsp;
      

<input name="ok2" type="reset" class="gg5" id="ok2" value="重置" >
</td>
    </tr>
  </table>
</form>
</body>
</html>
