<link rel="stylesheet" href="style.css" type="text/css"><!--#include file="lianjie.asp"----->
<%id=cint(trim(request("id")))
set rsgly=server.CreateObject("adodb.recordset")
rsgly.open"select * from dlb where id="&id,cn,1,3
%>
<title>管理页面</title>
<META http-equiv=Content-Type content=text/html; charset=gb2312>

<style type="text/css">
<!--
.style1 {color: #DBBF90;
	font-weight: bold;
}
body,td,th {
	font-size: 12px;
}
-->
</style>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="5" marginheight="0" marginwidth="0" bgcolor="#DDEEFF">
<form name="form1" method="post" action="cpdlb_xiugai.asp?id=<%=id%>">
  <table class="tableBorder" width="96%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
    <tr>
      <td  align="center" background="images/admin_bg_1.gif" height="25"><span class="style1">新闻类别修改</span></td>
    </tr>
    <tr>
      <td valign="top">
        <table width="100%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
          <tr bgcolor="#E8F1FF">
            <td width="20%" align="right"><span class="TableRow2">中文产品类别</span>：</td>
            <td width="80%" style="PADDING-LEFT: 10px"><span class="TableRow2">
              <input name="dlb" type="text" class="kuang" id="yhm4" value="<%=rsgly("dlb")%>">
            </span></td>
          </tr>
          <tr bgcolor="#E8F1FF">
            <td height="30"></td>
            <td height="30" style="PADDING-LEFT: 10px">
              <input name="Submit" type="submit" id="Submit" value="修改">
              <input type="reset" name="Submit22" value="重置">
            </td>
          </tr>
      </table></td>
    </tr>
  </table>
</form>
<%rsgly.close
set rsgly=nothing
cn.close
set cn=nothing%>
