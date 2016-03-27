<!--#include file="lianjie.asp"----->
<%set rsyjfl=server.CreateObject("adodb.recordset")
rsyjfl.open"select * from xwlb",cn,1,3%>
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
    <td  align="center" background="images/admin_bg_1.gif" height="25"><span class="style1">分类列表</span></td>
  </tr>
  <tr>
    <td valign="top">    <table width="100%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
      <tr bgcolor="#E8F1FF">
        <td align="center"><span class="BodyTitle">ＩＤ</span></td>
		<td align="center"><span class="BodyTitle">中文分类名称</span></td>
        <td align="center"><span class="BodyTitle"><strong>修改</strong></span></td>
        
        <td align="center"><span class="BodyTitle"><strong>删除</strong></span></td>
      </tr>
      <%for i=1 to rsyjfl.recordcount%>
      <tr align="center" bgcolor="#E8F1FF">
        <td height="15"><span class="TableRow2"><%=rsyjfl("id")%></span></td>
        
        <td height="15"><span class="TableRow2"><%=rsyjfl("xwlb")%></span></td>
        <td height="15"><span class="TableRow1"><span class="TableRow2"><span class="BodyTitle"><a href="xwlbxiugai.asp?id=<%=rsyjfl("id")%>">修改</a></span></span></span></td>
        
        <td height="15"><span class="TableRow2"><span class="BodyTitle"><a href="xwlbdel.asp?id=<%=rsyjfl("id")%>" onclick="{if(confirm('该操作不可恢复！\n\n确实删除选定的项目？\n注意！一旦删除此项目下的新闻也将被删除')){this.document.Prodlist.submit();return true;}return false;}">删除</a></span></span></td>
      </tr>
      <%rsyjfl.movenext
next
rsyjfl.close
set rsyjfl=nothing
cn.close
set cn=nothing%>
    </table></td>
  </tr>
</table>
<p>&nbsp;</p>
<p>&nbsp;</p>
