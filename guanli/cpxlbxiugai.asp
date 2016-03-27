<link rel="stylesheet" href="style.css" type="text/css">
<!--#include file="lianjie.asp"----->
<%id=cint(trim(request("id")))
set rsgly=server.CreateObject("adodb.recordset")
rsgly.open"select * from xlb where id="&id,cn,1,3

set rsdlb=server.CreateObject("adodb.recordset")
rsdlb.open"select * from dlb ",cn,1,3 

set rsdlb1=server.CreateObject("adodb.recordset")
rsdlb1.open"select * from dlb where id="&rsgly("dlb_id"),cn,1,3 
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
<form name="form1" method="post" action="cpxlb_xiugai.asp?id=<%=id%>">
  <table class="tableBorder" width="96%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
    <tr>
      <td  align="center" background="images/admin_bg_1.gif" height="25"><span class="style1">新闻类别修改</span></td>
    </tr>
    <tr>
      <td valign="top">
        <table width="100%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
          <tr bgcolor="#E8F1FF">
            <td width="20%" align="right">对应一级分类：</td>
            <td style="PADDING-LEFT: 10px"><span class="TableRow2">
              <select name="dlb_id">
			  <option value="<%=rsdlb1("id")%>"><%=rsdlb1("dlb")%></option>
			  <%for i=1 to rsdlb.recordcount%>
                <option value="<%=rsdlb("id")%>"><%=rsdlb("dlb")%></option>
				<%rsdlb.movenext
				next%>
              </select>
</span></td>
          </tr> <tr bgcolor="#E8F1FF">
            <td width="20%" align="right"><span class="TableRow2">中文新闻类别</span>：</td>
            <td width="80%" style="PADDING-LEFT: 10px"><span class="TableRow2">
              <input name="xlb" type="text" class="kuang" id="yhm4" value="<%=rsgly("xlb")%>">
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
