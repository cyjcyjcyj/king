<link rel="stylesheet" href="style.css" type="text/css">
<!--#include file="lianjie.asp"----->
<%id=cint(trim(request("id")))
set rsd=server.CreateObject("adodb.recordset")
rsd.open"select * from dingdan where id="&id,cn,1,3%>
<title>����ҳ��</title>
<META http-equiv=Content-Type content=text/html; charset=gb2312>

<style type="text/css">
<!--
body,td,th {
	font-size: 12px;
}
-->
</style>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="5" marginheight="0" marginwidth="0" bgcolor="#DDEEFF">
<form name="form1" method="post" action="xw_add.asp">
  <table cellpadding="2" cellspacing="1" border="0" width="96%" class="tableBorder" align=center>
    <tr valign="middle">
      <th height=26 colspan=5> ������ϸ��Ϣ</th>
    </tr>
    
    <tr align="center">
      <td width="29%" height=23  class="TableRow2"> ����������: </td>
      <td width="71%" height=23 align="left"  class="TableRow2"><%=rsd("sqrmc")%></td>
    </tr>
   <tr align="center">
      <td width="29%" height=23  class="TableRow2"> �����˵�ַ: </td>
      <td width="71%" height=23 align="left"  class="TableRow2"><%=rsd("sqrdz")%></td>
   </tr><tr align="center">
      <td width="29%" height=23  class="TableRow2"> �ʱ�: </td>
      <td width="71%" height=23 align="left"  class="TableRow2"><%=rsd("yb")%></td>
   </tr><tr align="center">
      <td width="29%" height=23  class="TableRow2"> Email: </td>
      <td width="71%" height=23 align="left"  class="TableRow2"><%=rsd("email")%></td>
   </tr><tr align="center">
      <td width="29%" height=23  class="TableRow2">��ϵ��:</td>
      <td width="71%" height=23 align="left"  class="TableRow2"><%=rsd("lxr")%></td>
   </tr><tr align="center">
      <td width="29%" height=23  class="TableRow2"> �绰:</td>
      <td width="71%" height=23 align="left"  class="TableRow2"><%=rsd("dh")%></td>
   </tr><tr align="center">
      <td width="29%" height=23  class="TableRow2"> ����:</td>
      <td width="71%" height=23 align="left"  class="TableRow2"><%=rsd("cz")%></td>
   </tr><tr align="center">
      <td width="29%" height=23  class="TableRow2"> ��Ʒ����:</td>
      <td width="71%" height=23 align="left"  class="TableRow2"><%=rsd("sbmc")%></td>
   </tr><tr align="center">
      <td width="29%" height=57  class="TableRow2"> ��Ʒ���˵��:</td>
      <td width="71%" height=90 align="left"  class="TableRow2"><%=rsd("sbsjsm")%></td>
   </tr>
  </table>
</form>

