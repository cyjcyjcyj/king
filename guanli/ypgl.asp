<!--#include file="conn.asp"------->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>无标题文档</title>
<style type="text/css">
<!--
body {
	background-color: #7B9AE7;
}
body,td,th {
	font-size: 12px;
}
.gg1 {
	border: 1px solid #FFFFFF;
}
.gg2 {
	border-top-width: 1px;
	border-right-width: 1px;
	border-bottom-width: 1px;
	border-left-width: 1px;
	border-top-style: solid;
	border-right-style: solid;
	border-bottom-style: none;
	border-left-style: none;
	border-top-color: #DBBF90;
	border-right-color: #DBBF90;
	border-bottom-color: #DBBF90;
	border-left-color: #DBBF90;
}
.gg3 {
	border-top-width: 1px;
	border-right-width: 1px;
	border-bottom-width: 1px;
	border-left-width: 1px;
	border-top-style: solid;
	border-right-style: none;
	border-bottom-style: none;
	border-left-style: none;
	border-top-color: #DBBF90;
	border-right-color: #DBBF90;
	border-bottom-color: #DBBF90;
	border-left-color: #DBBF90;
}
.gg4 {
	border: 1px solid #66CCFF;
}
.gg5 {
	border: 1px solid #000000;
}
.style3 {
	color: #DBBF90;
	font-weight: bold;
}
a:link {
	color: #000000;
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
</head>

<body><%set rszp=server.CreateObject("adodb.recordset")
 rszp.open"select * from jobbook",cn,1,3
 for i=1 to rszp.recordcount%>

<table border="0" align="center" cellpadding="0" cellspacing="0"  class=tableBorder>
  <tr>
    <td width="796" height="25"  background=images/top_bg.gif><strong>应聘管理 <a href="ypgl_del.asp?id=<%=rszp("id")%>">删除</a></strong></td>
  </tr>
  <tr>
    <td>
    
        <table width="100%" height="249" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#FFFFFF" bgcolor="#f1f3f5">
          <tr>
            <td width="17%" height="25">
              应聘岗位</td>
            <td colspan="2">&nbsp;&nbsp;<b><%=rszp("zy")%></b> &nbsp; </td>
            <td>&nbsp;</td>
          </tr>
         
          <tr>
            <td height="10">
              姓 名</td>
            <td width="34%">&nbsp;&nbsp;<%=rszp("xm")%></td>
            <td width="13%">性 别</td>
            <td width="36%"><%=rszp("xb")%></td>
          </tr>
          <tr>
            <td height="22">
              出生日期</td>
            <td>&nbsp;&nbsp;<%=rszp("csrq")%></td>
            <td>婚姻状况</td>
            <td><%=rszp("hyqk")%></td>
          </tr>
          <tr>
            <td height="22">毕业院校</td>
            <td colspan="3">&nbsp;&nbsp;<%=rszp("byyx")%></td>
          </tr>
          <tr>
            <td height="22">学 历</td>
            <td>&nbsp;&nbsp;<%=rszp("xl")%></td>
            <td>专 业</td>
            <td><%=rszp("zy")%></td>
          </tr>
          <tr>
            <td height="22">毕业时间</td>
            <td>&nbsp;&nbsp;<%=rszp("bysj")%></td>
            <td><font color="#FF0000">应聘日期</font></td>
            <td><font color="#FF0000"><%=rszp("sqsj")%></font></td>
          </tr>
          <tr>
            <td height="22">电 话</td>
            <td colspan="3">&nbsp;&nbsp;<%=rszp("dh")%></td>
          </tr>
          <tr>
            <td height="22">E-mail</td>
            <td colspan="3">&nbsp;&nbsp;<a href="Sendmail.asp?email=<%=rszp("email")%>"><%=rszp("email")%></a></td>
          </tr>
          <tr>
            <td height="22">联系地址</td>
            <td colspan="3">&nbsp;&nbsp;<%=rszp("lxdz")%></td>
          </tr>
       
          <tr>
            <td height="25">水平与能力</td>
            <td height="25" colspan="3">&nbsp;&nbsp;<%=rszp("nlzc")%></td>
          </tr>
          <tr>
            <td height="25">
              个人简历</td>
            <td height="25" colspan="3">&nbsp;&nbsp;<%=rszp("grjl")%></td>
          </tr>
        </table>
    </td>
  </tr>
</table><br><br><%rszp.movenext
next
%>

</body>
</html>
