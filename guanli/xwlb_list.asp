<!--#include file="lianjie.asp"----->
<%xwlb=request("xwlb")
set rsgly=server.CreateObject("adodb.recordset")
rsgly.open"select * from xw where xwlb='"&xwlb&"'",cn,1,3%>
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
    <td height="20"  align="center" bgcolor="#E8F1FF"><span class="style1"><a href="xwlb.asp?xwlb=&#28909;&#28857;&#26032;&#38395;">�ȵ������б�</a> <a href="xwlb.asp?xwlb=&#34892;&#19994;&#21160;&#24577;">��ҵ��̬�б�</a></span></td>
  </tr>
  <tr>
    <td valign="top">    <table width="100%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
      <tr bgcolor="#E8F1FF">
        <td align="center"><span class="TableRow2">���ű���</span><span class="TableRow2"> </span></td>
        <td align="center">�������</td>
        <td align="center"><span class="BodyTitle">���ʱ��</span></td>
        <td align="center"><span class="BodyTitle">�޸�</span></td>
        <td align="center"><span class="BodyTitle">ɾ��</span></td>
      </tr>
      <%for i=1 to rsgly.recordcount%>
      <tr align="center" bgcolor="#E8F1FF">
        <td height="15" align="left"><span class="TableRow2"><%=rsgly("xwbt")%></span></td>
        <td height="15"><span class="TableRow2"><%=rsgly("xwlb")%></span></td>
        <td height="15"><span class="TableRow1"><%=rsgly("tjsj")%></span></td>
        <td height="15"><span class="TableRow1"><a href="xwxiugai.asp?id=<%=rsgly("id")%>">�޸�</a></span></td>
        <td height="15"><span class="TableRow1"><a href="xw_del.asp?id=<%=rsgly("id")%>">ɾ��</a></span></td>
      </tr>
      <%rsgly.movenext
next%>
    </table></td>
  </tr>
</table>
