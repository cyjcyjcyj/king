<!--#include file="lianjie.asp"----->
<%set rsgly=server.CreateObject("adodb.recordset")
rsgly.open"select * from dlb",cn,1,3%>
<title>管理页面</title>
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
    <td  align="center" background="images/admin_bg_1.gif" height="25"><span class="style1">产品类别列表[<a href="cpdlbadd.asp"><font color="#FFFFFF">添加大类别</font></a>] </span></td>
  </tr>
  <tr>
    <td valign="top">    <table width="100%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
      <tr bgcolor="#E8F1FF">
        <td colspan="5" align="center"><table width="100%" height="61"  border="0" cellpadding="0" cellspacing="0">
          <%for i=1 to rsgly.recordcount%>
		  <tr>
            <td width="50%"><img src="images/d.gif" width="15" height="15"><%=rsgly("dlb")%></td>
            <td width="50%"> <a href="cpdlbxiugai.asp?id=<%=rsgly("id")%>">修改</a> <a href="cpdlb_del.asp?id=<%=rsgly("id")%>">删除</a></td>
		  </tr>
          <tr>
            <td height="43" colspan="2"><table width="76%"  border="0" cellspacing="0" cellpadding="0">
			<%set rsxlb=server.CreateObject("adodb.recordset")
			rsxlb.open"select * from xlb where dlb_id="&rsgly("id"),cn,1,3
			for j=1 to rsxlb.recordcount%>
              <tr>
                <td width="3%">&nbsp;</td>
                <td width="48%"><img src="images/x.gif" width="15" height="15"><%=rsxlb("xlb")%></td>
                <td width="49%"><a href="cpxlbxiugai.asp?id=<%=rsxlb("id")%>">修改</a> <a href="cpxlb_del.asp?id=<%=rsxlb("id")%>">删除</a></td>
              </tr>
			  <%rsxlb.movenext
			  next%>
            </table>
              <br></td>
          </tr>
		  <%rsgly.movenext
		  next%>
        </table></td>
        </tr>
    
    </table></td>
  </tr>
</table>
