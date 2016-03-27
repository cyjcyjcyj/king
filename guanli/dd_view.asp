<link rel="stylesheet" href="style.css" type="text/css">
<!--#include file="lianjie.asp"----->
<%id=cint(trim(request("id")))
set rsd=server.CreateObject("adodb.recordset")
rsd.open"select * from dingdan where id="&id,cn,1,3%>
<title>管理页面</title>
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
      <th height=26 colspan=5> 订单详细信息</th>
    </tr>
    
    <tr align="center">
      <td width="29%" height=23  class="TableRow2"> 申请人名称: </td>
      <td width="71%" height=23 align="left"  class="TableRow2"><%=rsd("sqrmc")%></td>
    </tr>
   <tr align="center">
      <td width="29%" height=23  class="TableRow2"> 申请人地址: </td>
      <td width="71%" height=23 align="left"  class="TableRow2"><%=rsd("sqrdz")%></td>
   </tr><tr align="center">
      <td width="29%" height=23  class="TableRow2"> 邮编: </td>
      <td width="71%" height=23 align="left"  class="TableRow2"><%=rsd("yb")%></td>
   </tr><tr align="center">
      <td width="29%" height=23  class="TableRow2"> Email: </td>
      <td width="71%" height=23 align="left"  class="TableRow2"><%=rsd("email")%></td>
   </tr><tr align="center">
      <td width="29%" height=23  class="TableRow2">联系人:</td>
      <td width="71%" height=23 align="left"  class="TableRow2"><%=rsd("lxr")%></td>
   </tr><tr align="center">
      <td width="29%" height=23  class="TableRow2"> 电话:</td>
      <td width="71%" height=23 align="left"  class="TableRow2"><%=rsd("dh")%></td>
   </tr><tr align="center">
      <td width="29%" height=23  class="TableRow2"> 传真:</td>
      <td width="71%" height=23 align="left"  class="TableRow2"><%=rsd("cz")%></td>
   </tr><tr align="center">
      <td width="29%" height=23  class="TableRow2"> 产品名称:</td>
      <td width="71%" height=23 align="left"  class="TableRow2"><%=rsd("sbmc")%></td>
   </tr><tr align="center">
      <td width="29%" height=57  class="TableRow2"> 产品设计说明:</td>
      <td width="71%" height=90 align="left"  class="TableRow2"><%=rsd("sbsjsm")%></td>
   </tr>
  </table>
</form>

