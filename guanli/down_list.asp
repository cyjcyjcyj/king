<!--#include file="lianjie.asp"----->
<%set rsyjfl=server.CreateObject("adodb.recordset")
rsyjfl.open"select * from down",cn,1,3%>
<title>����ҳ��</title>
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
    <td  align="center" background="images/admin_bg_1.gif" height="25"><span class="style1">�����б�</span></td>
  </tr>
  <tr>
    <td valign="top">    <table width="100%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
      <tr bgcolor="#E8F1FF">
        <td align="center"><span class="BodyTitle">�ɣ�</span></td>
		<td align="center"><span class="BodyTitle">��������</span></td>
        <td align="center">���ʱ��</td>
        <td align="center"><span class="BodyTitle"><strong>�޸�</strong></span></td>
        
        <td align="center"><span class="BodyTitle"><strong>ɾ��</strong></span></td>
      </tr>
      <%for i=1 to rsyjfl.recordcount%>
      <tr align="center" bgcolor="#E8F1FF">
        <td height="15"><span class="TableRow2"><%=rsyjfl("id")%></span></td>
        
        <td height="15"><span class="TableRow2"><%=rsyjfl("btmc")%></span></td>
        <td height="15"><span class="TableRow2"><%=rsyjfl("tjsj")%></span></td>
        <td height="15"><span class="TableRow1"><span class="TableRow2"><span class="BodyTitle"><a href="downxiugai.asp?id=<%=rsyjfl("id")%>">�޸�</a></span></span></span></td>
        
        <td height="15"><span class="TableRow2"><span class="BodyTitle"><a href="down_del.asp?id=<%=rsyjfl("id")%>" onclick="{if(confirm('�ò������ɻָ���\n\nȷʵɾ��ѡ������Ŀ��\nע�⣡һ��ɾ������Ŀ�µ�������ĿҲ����ɾ��')){this.document.Prodlist.submit();return true;}return false;}">ɾ��</a></span></span></td>
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
