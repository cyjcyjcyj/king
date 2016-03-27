<!--#include file="conn.asp"------->
<%
 
 set rszp=server.CreateObject("adodb.recordset")
 rszp.open"select * from job",cn,1,3%>
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
}
a:link {
	color: #000000;
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
.tableBorder {	BORDER-RIGHT: #183789 1px solid; BORDER-TOP: #183789 1px solid; BORDER-LEFT: #183789 1px solid; WIDTH: 97%; BORDER-BOTTOM: #183789 1px solid; BACKGROUND-COLOR: #ffffff
}
-->
</style>
 </head>

<body>
<table height="25" border="0" cellpadding="0" cellspacing="0"  class=tableBorder>
  <tr>
    <td  background=images/top_bg.gif height="25"><div align="center" style="color: #000000; font-weight: bold;"><strong>招聘管理</strong></div></td>
  </tr>
  <tr>
    <td>
      <table width="100%" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#FFFFFF" bgcolor="#f1f3f5">
        <tr >
          <td width="8%" height="25">
            <div align="center">ID</div></td>
          <td width="35%">
            <div align="center">招聘对象</div></td>
          <td width="11%"><div align="center">招聘人数</div></td>
          <td width="11%"><div align="center">发布时间</div></td>
          <td width="11%"><div align="center">有效期限</div></td>
          <td width="10%">
            <div align="center">操作</div></td>
          <td width="14%">
            <div align="center">操作</div></td>
        </tr>
        <%
for i=1 to rszp.recordcount
%>
        <tr>
          <td height="22">
            <div align="center"><%=rszp("id")%></div></td>
          <td >&nbsp;&nbsp;<%=rszp("zpdx")%></td>
          <td><div align="center"><%=rszp("zprs")%>人</div></td>
          <td><div align="center"><%=rszp("fbsj")%></div></td>
          <td><div align="center"><%=rszp("yxqx")%>天</div></td>
          <td>
            <div align="center"><a href="zpglxiugai.asp?id=<%=rszp("id")%>" >修改</a></div></td>
          <td>
            <div align="center"><a href="zpgl_del.asp?id=<%=rszp("id")%>">删除</a></div></td>
        </tr>
		<%rszp.movenext
		next%>
       
        <tr>
          <td height="25" colspan="7" align="center">&nbsp;&nbsp;
          </td>
        </tr>
      </table>
    </td>
  </tr>
</table>
</body>
</html>
