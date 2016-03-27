<!--#include file="conn.asp"-->
<%
id=cint(trim(request("id")))
set rs=server.CreateObject("adodb.recordset")
rs.open"select * from jhcp where id="&id,cn,1,3 
%>
<html>
<head>
<title>广州锦倍塔信息科技有限公司</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<script src="Scripts/AC_RunActiveContent.js" type="text/javascript"></script>
<style type="text/css">
<!--
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
.body{width:1050px;margin:0 auto;}
.context{width:760px;height:1000px;margin:0 20px 50px 0; }
.right{width:260px;float:right;}
-->
</style></head>
<body>
<!--#include file="top.asp"--->
<div class="body">

<div class="right"><!--#include file="rightsy.asp"---></div>

<div class="context">
<font size="2"><b><%=rs("cpmc")%></b></font>
<td width="735" align="center"><br>
<img src="guanli/<%=rs("tpdz1")%>" ><br>
                  <br></td>


</div>
</div>

<!--#include file="bottom.asp"--->

<!-- ImageReady Slices (okk1.psd) -->
<!-- End ImageReady Slices -->
</body>
</html>
