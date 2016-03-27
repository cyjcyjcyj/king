<!--#include file="conn.asp"------->
<%
 id=cint(trim(request("id")))
 set rszp=server.CreateObject("adodb.recordset")
 rszp.open"select * from job where id="&id,cn,1,3%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charszpet=gb2312">
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
.tableBorder {	BORDER-RIGHT: #183789 1px solid; BORDER-TOP: #183789 1px solid; BORDER-LEFT: #183789 1px solid; WIDTH: 97%; BORDER-BOTTOM: #183789 1px solid; BACKGROUND-COLOR: #ffffff
}
.style5 {color: #000000}
-->
</style>
</head>

<body>
<table height="314" border="0" cellpadding="0" cellspacing="0"  class=tableBorder>
  <tr>
    <td  background=images/top_bg.gif height="25"><div align="center" style="color: #DBBF90; font-weight: bold;"><strong>修改招聘信息</strong></div></td>
  </tr>
  <tr>
    <form method="post" action="zpgl_xiugai.asp?id=<%=id%>">
    
      <td><div align="center">
          <table width="100%" height="321" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#FFFFFF" bgcolor="#f1f3f5">
            <tr>
              <td width="19%" height="25">
                <div align="center">招聘对象</div></td>
              <td width="81%">              <input name="zpdx" type="text" class="inputtext" id="zpdx" value="<%=rszp("zpdx")%>" size="30" maxlength="30"></td>
            </tr>
            <tr>
              <td height="25">
                <div align="center">招聘人数</div></td>
              <td ><input name="zprs" type="text" class="inputtext" id="zprs" value="<%=rszp("zprs")%>" size="5" maxlength="30">
                人</td>
            </tr>
            <tr>
              <td height="25">
                <div align="center">工作地点</div></td>
              <td >
                <input name="gzdd" type="text" class="inputtext" id="gzdd" value="<%=rszp("gzdd")%>" size="30" maxlength="50">
              </td>
            </tr>
            <tr>
              <td height="25"><div align="center">工资待遇</div></td>
              <td>
                <input name="gzdy" type="text" class="inputtext" id="gzdy" value="<%=rszp("gzdy")%>" size="30" maxlength="50"></td>
            </tr>
            <tr>
              <td height="25">
                <div align="center">发布日期</div></td>
              <td><input name="fbsj" type="text" class="inputtext" id="fbsj" value="<%=rszp("fbsj")%>" size="30" maxlength="50"> </td>
            </tr>
            <tr>
              <td height="25"><div align="center">有效期限</div></td>
              <td ><input name="yxqx" type="text" class="inputtext" id="yxqx" value="<%=rszp("yxqx")%>" size="5" maxlength="30">
                天</td>
            </tr>
            <tr>
              <td height="25"><div align="center">招聘要求</div></td>
              <td>
                <textarea name="zpyq" cols="50" rows="5" class="inputtext" id="zpyq"><%=rszp("zpyq")%></textarea></td>
            </tr>
            <tr>
              <td height="25" colspan="2">
                <div align="center">
                  <input name="submit" type="submit" value=" 确 定 ">
&nbsp;
                  <input name="reset" type="reset" value=" 取 消 ">
              </div></td>
            </tr>
          </table>
      </div></td>
    </form>
  </tr>
</table>
</body>
</html>
