<!--#include file="lianjie.asp"----->
<%xwlb=request("xwlb")
set rsgly=server.CreateObject("adodb.recordset")
rsgly.open"select * from xw where xwlb='"&xwlb&"'",cn,1,3%>
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
    <td  align="center" background="images/admin_bg_1.gif" height="25"><span class="style1">新闻列表</span></td>
  </tr>
  <tr>
    <td height="20"  align="center" bgcolor="#E8F1FF"><span class="style1"><a href="xwlb.asp?xwlb=&#28909;&#28857;&#26032;&#38395;">热点新闻列表</a> <a href="xwlb.asp?xwlb=&#34892;&#19994;&#21160;&#24577;">行业动态列表</a></span></td>
  </tr>
  <tr>
    <td valign="top">    <table width="100%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
      <tr bgcolor="#E8F1FF">
        <td align="center"><span class="TableRow2">新闻标题</span><span class="TableRow2"> </span></td>
        <td align="center">新闻类别</td>
        <td align="center"><span class="BodyTitle">添加时间</span></td>
        <td align="center"><span class="BodyTitle">修改</span></td>
        <td align="center"><span class="BodyTitle">删除</span></td>
      </tr>
      <%for i=1 to rsgly.recordcount%>
      <tr align="center" bgcolor="#E8F1FF">
        <td height="15" align="left"><span class="TableRow2"><%=rsgly("xwbt")%></span></td>
        <td height="15"><span class="TableRow2"><%=rsgly("xwlb")%></span></td>
        <td height="15"><span class="TableRow1"><%=rsgly("tjsj")%></span></td>
        <td height="15"><span class="TableRow1"><a href="xwxiugai.asp?id=<%=rsgly("id")%>">修改</a></span></td>
        <td height="15"><span class="TableRow1"><a href="xw_del.asp?id=<%=rsgly("id")%>">删除</a></span></td>
      </tr>
      <%rsgly.movenext
next%>
    </table></td>
  </tr>
</table>
