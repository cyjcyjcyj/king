<!--#include file="conn.asp"--->
<%
if id=75 then
bg="images/gsjj_bt.jpg"
end if

if id=76 then
bg="images/xswl_bt.jpg"
end if

if id=77 then
bg="images/lxwm_bt.jpg"
end if
%>
<html>
<head>
<title>广州锦倍塔信息科技有限公司</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<script src="Scripts/AC_RunActiveContent.js" type="text/javascript"></script>
<style type="text/css">
<!--
.body{width:1090px;margin:0 auto;}

.right{width:260px;float:right;}
.context{width:760px;min-height:800px;margin:0 20px 50px 0;}
.toppic{text-align:center;padding-left:20px;}
-->
</style>
</head>

<body>

<!--#include file="top.asp"--->

<div class="body">
<div class="right"><!--#include file="rightsy.asp"---></div>
<%
id=cint(trim(request("id")))
set rsgsxm=server.CreateObject("adodb.recordset")
rsgsxm.open"select * from gsxmnr where id="&id,cn,1,3
%>
<div class="context">
<div class="toppic"><img src="images/toppic.png" alt=""></div>
<%=rsgsxm("gsxmnr")%>

</div>  
</div>


<!--#include file="bottom.asp"--->

<!-- ImageReady Slices (okk1.psd) -->
<!-- End ImageReady Slices -->
</body>
</html>
