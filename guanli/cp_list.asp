<!--#include file="lianjie.asp"----->
<%set rscp=server.CreateObject("adodb.recordset")
rscp.open"select * from cp",cn,1,3%>
<title>����ҳ��</title>
<META http-equiv=Content-Type content=text/html; charset=gb2312 charset=gb2312>
<link rel="stylesheet" href="style.css" type="text/css">
<style type="text/css">
<!--
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
</style><BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="5" marginheight="0" marginwidth="0" bgcolor="#DDEEFF">
<table width="96%" border="0" align=center cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="tableBorder">
<tr valign="middle" background="images/admin_bg_1.gif">
<th height=26 colspan=5 background="images/admin_bg_1.gif">
  ��ʦ�б�</th>
</tr>
<tr align="center" bgcolor="#DEEFFF">
<td width="25%" height=25 class=BodyTitle>��ʦ����</td>

<td width="27%" height=25 class=BodyTitle>���ʱ��</td>
<td width="8%" class=BodyTitle>�޸�</td>
<td width="15%" class=BodyTitle>ɾ��</td>
</tr><%for i=1 to rscp.recordcount%>
<tr align="center" bgcolor="#DEEFFF">
<td height=23  class="TableRow2"><%=rscp("cpmc")%></td>
<td class="TableRow1"><%=rscp("tjsj")%></td>
<td class="TableRow1"><a href="cpxiugai.asp?id=<%=rscp("id")%>">�޸�</a></td>
<td class="TableRow1"><a href="cp_del.asp?id=<%=rscp("id")%>">ɾ��</a></td>
</tr>
<%rscp.movenext
next%>
</table>
<%rscp.close
set rscp=nothing
 cn.close
 set cn=nothing%>

