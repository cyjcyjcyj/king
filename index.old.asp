<!--#include file="conn.asp"---><html>
<head>
<title>广州锦倍塔信息科技有限公司</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<script src="Scripts/AC_RunActiveContent.js" type="text/javascript"></script>
<style type="text/css">
<!--
body,td,th {
	font-size: 12px;
	color: #DBBF90;
}
body {
	background-image: url(images/index_22.jpg);
}
a:link {
	color: #DBBF90;
	text-decoration: none;
}
a:visited {
	text-decoration: none;
	color: #DBBF90;
}
a:hover {
	text-decoration: underline;
	color: #00CC33;
}
a:active {
	text-decoration: none;
	color: #00CF31;
}
-->
</style></head>
<body bgcolor="#FFFFFF" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="1002" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td colspan="2"><!--#include file="top.asp"---></td>
  </tr>
  <tr>
    <td width="758"><table id="__01" width="758" height="552" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td><img src="images/index_21_01_01.jpg" alt="" width="758" height="110" border="0" usemap="#Maps1"></td>
      </tr>
      <tr>
        <td height="189"><%
id=cint(trim(request("id")))
set rsgsxm=server.CreateObject("adodb.recordset")
rsgsxm.open"select * from gsxmnr where id=74",cn,1,3
response.Write rsgsxm("gsxmnr")
rsgsxm.close
set rsgsxm=nothing%></td>
      </tr>
      <tr>
        <td><img src="images/index_21_01_03.jpg" alt="" width="758" height="61" border="0" usemap="#Maps2"></td>
      </tr>
      <tr>
        <td height="118"><table width="758" height="151" border="0" cellpadding="0" cellspacing="0">
          <tr>
            <td width="26" height="151">&nbsp;</td>
            <td height="151" valign="top" style="line-height:2"><table width="705" cellpadding="0" cellspacing="0">
              <%set rsxwok=server.CreateObject("adodb.recordset")
rsxwok.open"select top 6 * from (select *  from xw order by tjsj desc)",cn,1,3
	  for i=1 to rsxwok.recordcount%>
              <TR>
                <TD width="28" height="25" align="center"><img src="images/xwbt.jpg" width="4" height="4"></TD>
                <TD width="559"><A  href="xw_view.asp?id=<%=rsxwok("id")%>&dbid=d2"><%=rsxwok("xwbt")%></A></TD>
                <TD width="116"><%=rsxwok("tjsj")%></TD>
              </TR>
              <%rsxwok.movenext
	  next
	  rsxwok.close
	  set rsxwok=nothing%>
            </table></td>
            </tr>
        </table></td>
      </tr>
      <tr>
        <td><img src="images/index_21_01_05.jpg" alt="" width="758" height="74" border="0" usemap="#Maps3"></td>
      </tr>
    </table></td>
    <td width="244" align="center" valign="top"><!--#include file="rightsy.asp"---></td>
  </tr>
  <tr>
    <td height="109" colspan="2" align="center" background="images/index_22.jpg"><!--#include file="gd.asp"---></td>
  </tr>
  <tr>
    <td height="72" colspan="2"><!--#include file="bottom.asp"---></td>
  </tr>
</table>
<!-- ImageReady Slices (okk1.psd) -->
<!-- End ImageReady Slices -->

<map name="Maps1"><area shape="rect" coords="685,81,738,104" href="gsjj.asp?id=75">
</map>
<map name="Maps2"><area shape="rect" coords="686,28,742,52" href="xw_list.asp">
</map>
<map name="Maps3"><area shape="rect" coords="678,43,746,68" href="cpzs.asp">
</map></body>
</html>
