<!--#include file="lianjie.asp"-->
<%set rsdlb=server.CreateObject("adodb.recordset")
rsdlb.open"select * from dlb ",cn,1,3%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��̨����</title>
<style type="text/css">
<!--
body,td,th {
	font-size: 12px;
}
.style1 {	color: #DBBF90;
	font-weight: bold;
}
-->
</style></head>

<body>
<form name="form" method="post" action="cpxlb_add.asp">
  <table class="tableBorder" width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
    <tr>
      <td  align="center" background="images/admin_bg_1.gif" height="25"><span class="style1">��Ʒ�������</span></td>
    </tr>
    <tr>
      <td valign="top">
        <table width="100%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
          <tr bgcolor="#E8F1FF">
            <td width="16%" align="right">��Ӧһ�����ࣺ</td>
            <td style="PADDING-LEFT: 10px"><span class="TableRow2">
              <select name="dlb_id">
			  <%for i=1 to rsdlb.recordcount%>
                <option value="<%=rsdlb("id")%>"><%=rsdlb("dlb")%></option>
				<%rsdlb.movenext
				next%>
              </select>
</span></td>
          </tr><tr bgcolor="#E8F1FF">
            <td width="16%" align="right">���ķ������ƣ�</td>
            <td style="PADDING-LEFT: 10px"><span class="TableRow2">
            <input name="xlb" type="text" id="xlb" size="40">
</span></td>
          </tr>
          <tr bgcolor="#E8F1FF">
            <td height="30"></td>
            <td height="30" style="PADDING-LEFT: 10px">
              <input type="submit" name="Submit" value="���">
              <input type="reset" name="Submit2" value="����">
            </td>
          </tr>
      </table></td>
    </tr>
  </table>
</form>
</body>
</html>
