<!--#include file="lianjie.asp"----->
<%set rsgly=server.CreateObject("adodb.recordset")
rsgly.open"select * from yqlj",cn,1,3%>
<title>管理页面</title>
<META http-equiv=Content-Type content=text/html; charset=gb2312 charset=gb2312>

<style type="text/css">
<!--
.style1 {color: #DBBF90;
	font-weight: bold;
}
body,td,th {
	font-size: 12px;
}
a:link {
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
</style>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="5" marginheight="0" marginwidth="0" bgcolor="#DDEEFF">
<table class="tableBorder" width="96%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td  align="center" background="images/admin_bg_1.gif" height="25"><span class="style1">友情链接列表 </span></td>
  </tr>
  <tr>
    <td valign="top">    <table width="100%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
      <tr bgcolor="#E8F1FF">
        <td align="center">网站名称<span class="TableRow2"></span></td>
        <td align="center"><span class="BodyTitle">网站网址</span></td>
        <td align="center">是否图片链接</td>
        <td align="center"><span class="BodyTitle">修改</span></td>
        <td align="center"><span class="BodyTitle">删除</span></td>
      </tr>
      <%for i=1 to rsgly.recordcount%>
      <tr align="center" bgcolor="#E8F1FF">
        <td height="15"><span class="TableRow2"><%=rsgly("wzmc")%></span></td>
        <td height="15"><span class="TableRow2"><%=rsgly("wzwz")%></span></td>
        <td height="15"><span class="TableRow2"><%=rsgly("sfstplj")%></span></td>
        <td height="15"><span class="TableRow1"><a href="yqljxiugai.asp?id=<%=rsgly("id")%>">修改</a></span></td>
        <td height="15"><span class="TableRow1"><a href="yqlj_del.asp?id=<%=rsgly("id")%>">删除</a></span></td>
      </tr>
      <%rsgly.movenext
next%>
    </table></td>
  </tr>
</table>
