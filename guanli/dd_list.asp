<!--#include file="lianjie.asp"----->
<%set rsdg=server.CreateObject("adodb.recordset")
rsdg.open"select * from dingdan",cn,1,3%>
<title>����ҳ��</title><script language="JavaScript">

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
  �����б�</th>
</tr>
<tr align="center">
<td width="21%" height=25 class=BodyTitle>����������</td>
<td width="23%" class=BodyTitle> �����˵�ַ:</td>
<td width="13%" height=25 class=BodyTitle> ��ϵ��</td>
<td width="13%" class=BodyTitle> �绰��</td>
<td width="16%" height=25 class=BodyTitle>����ʱ��</td>
<td width="8%" class=BodyTitle>��ϸ��Ϣ</td>
<td width="6%" class=BodyTitle>ɾ��</td>
</tr><%for i=1 to rsdg.recordcount%>
<tr align="center">
<td height=23  class="TableRow2"><%=rsdg("sqrmc")%></td>
<td height=23  class="TableRow2"><%=rsdg("sqrdz")%></td>
<td class="TableRow1"><%=rsdg("lxr")%></td>
<td class="TableRow1"><%=rsdg("dh")%></td>
<td class="TableRow1"><%=rsdg("tjsj")%></td>
<td class="TableRow1"><a href='#' onclick="openWindow('dd_view.asp?id=<%=rsdg("id")%>','','width=600,height=200,scrollbars=yes')">��ϸ��Ϣ</a></td>
<td class="TableRow1"><a href="dd_del.asp?id=<%=rsdg("id")%>">ɾ��</a></td>
</tr>
<%rsdg.movenext
next%>

</table>


