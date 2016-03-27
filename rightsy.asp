<%set rscp=server.CreateObject("adodb.recordset")
rscp.open"select * from dlb order by endlb asc",cn,1,3%>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
*{margin:0;padding:0;}
.product{width:260px;height:360px;margin-bottom:10px;}
.in1{padding:20px;border:4px solid #0068b7;margin-top:5px;height:260px;}
.in2{padding:22px;border:4px solid #0068b7;margin-top:5px;height:100px;}
.in3{padding:22px;border:4px solid #0068b7;margin-top:5px;height:60px;}
.mail{width:260px;height:180px;margin-bottom:10px;text-align:center}
.connect{margin-top:35px}
-->
</style>
</head>
<body>
<!-- ImageReady Slices (index_21_02.jpg) -->

<div class="product">
    <img src="images/product_list.jpg" alt="">
    <%for i=1 to rscp.recordcount%>
    <div class="in1">
           <img src="images/xwbt.jpg">
           <a  href="cpzs.asp?dlb_id=<%=rscp("id")%>&dbid=d4"><%=rscp("dlb")%></a>
    </div> 
	<%rscp.movenext
	next%>
</div>
 
<div class=mail>
<img src="images/mail.jpg" alt="">
<div class="in2">
<form name="checkform" action="" method="POST" id="checkform">
          <p>
            用户名：<input name="name" type="txt" id="password" value="" size="10">
          </p>
          <p>&nbsp;&nbsp;&nbsp;</p>
          <p>
            密码：<input name="password" type="password" id="password" value="" size="10">
          </p> 
     
<input  style="height:30px;width:60px;margin-top:20px;margin-right:20px" type="button" name="log" value="登录" onclick="#"/>
<input  style="height:30px;width:60px;margin-top:20px" type="button" name="regi" value="注册" onclick="#"/>
</form>
          
 </div>
 </div>
 
<div class="connect" align="center">
<img src="images/connect.jpg" alt="">
<div class="in3">
<p>24小时服务热线</p>
<p>400-622-1350</p>
</div>
</div>

<!--<td><%
set rsgsxm=server.CreateObject("adodb.recordset")
rsgsxm.open"select * from gsxmnr where id=78",cn,1,3
response.Write rsgsxm("gsxmnr")
rsgsxm.close
set rsgsxm=nothing%>
</td>-->


