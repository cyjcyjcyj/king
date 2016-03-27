<!--#include file="lianjie.asp"----->
<%set rsdg=server.CreateObject("adodb.recordset")
rsdg.open"select * from dingdan",cn,1,3%>
<title>管理页面</title><script language="JavaScript">

function openWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}

</script>
<META http-equiv=Content-Type content=text/html; charset=gb2312>
<link rel="stylesheet" href="style.css" type="text/css">
<style type="text/css">
<!--
body,td,th {
	font-size: 12px;
}
-->
</style>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="5" marginheight="0" marginwidth="0" bgcolor="#DDEEFF">
<table cellpadding="2" cellspacing="1" border="0" width="96%" class="tableBorder" align=center>
<tr valign="middle">
<th height=26 colspan=7>
  申请列表</th>
</tr>
<tr align="center">
<td width="21%" height=25 class=BodyTitle>申请者名称</td>
<td width="23%" class=BodyTitle> 申请人地址:</td>
<td width="13%" height=25 class=BodyTitle> 联系人</td>
<td width="13%" class=BodyTitle> 电话：</td>
<td width="16%" height=25 class=BodyTitle>申请时间</td>
<td width="8%" class=BodyTitle>详细信息</td>
<td width="6%" class=BodyTitle>删除</td>
</tr><%for i=1 to rsdg.recordcount%>
<tr align="center">
<td height=23  class="TableRow2"><%=rsdg("sqrmc")%></td>
<td height=23  class="TableRow2"><%=rsdg("sqrdz")%></td>
<td class="TableRow1"><%=rsdg("lxr")%></td>
<td class="TableRow1"><%=rsdg("dh")%></td>
<td class="TableRow1"><%=rsdg("tjsj")%></td>
<td class="TableRow1"><a href='#' onclick="openWindow('dd_view.asp?id=<%=rsdg("id")%>','','width=600,height=200,scrollbars=yes')">详细信息</a></td>
<td class="TableRow1"><a href="dd_del.asp?id=<%=rsdg("id")%>">删除</a></td>
</tr>
<%rsdg.movenext
next%>

</table>


