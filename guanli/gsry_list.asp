<!--#include file="lianjie.asp"----->
<%leixing=request("leixing")
set rscp=server.CreateObject("adodb.recordset")
rscp.open"select * from gsry where leixing='"&leixing&"'",cn,1,3%>
<title>管理页面</title>
<script language="javascript">
var flag=false;
function DrawImage(ImgD,w,h){
var image=new Image();
image.src=ImgD.src;
if(image.width>0 && image.height>0){
flag=true;
if(image.width/image.height>= w/h){
if(image.width>w){
ImgD.width=w;
ImgD.height=(image.height*w)/image.width;
}else{
ImgD.width=image.width;
ImgD.height=image.height;
}
//ImgD.alt=image.width+"×"+image.height;
}
else{
if(image.height>h){
ImgD.height=h;
ImgD.width=(image.width*h)/image.height;
}else{
ImgD.width=image.width;
ImgD.height=image.height;
}
//ImgD.alt=image.width+"×"+image.height;
}
}
}
</script> 
<META http-equiv=Content-Type content=text/html; charset=gb2312 charset=gb2312>
<link rel="stylesheet" href="style.css" type="text/css">
<style type="text/css">
<!--
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
</style><BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="5" marginheight="0" marginwidth="0" bgcolor="#DDEEFF">
<table width="96%" border="0" align=center cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="tableBorder">
<tr valign="middle" background="images/admin_bg_1.gif">
<th height=26 colspan=6 background="images/admin_bg_1.gif">
  <%=leixing%>列表</th>
</tr>
<tr align="center" bgcolor="#DEEFFF">
<td width="25%" height=25 class=BodyTitle>图片说明</td>


<td width="14%" class=BodyTitle>对应图片</td>
<td width="8%" class=BodyTitle>修改</td>
<td width="15%" class=BodyTitle>删除</td>
</tr><%for i=1 to rscp.recordcount%>
<tr align="center" bgcolor="#DEEFFF">
<td height=23  class="TableRow2"><%if rscp("tpsm")="" then response.Write "暂无" else response.Write rscp("tpsm") end if%></td>
<td class="TableRow1"><img src="<%=rscp("tpdz")%>" onload="DrawImage(this,216,143);"></td>
<td class="TableRow1"><a href="gsryxiugai.asp?id=<%=rscp("id")%>&leixing=<%=rscp("leixing")%>">修改</a></td>
<td class="TableRow1"><a href="gsry_del.asp?id=<%=rscp("id")%>&leixing=<%=rscp("leixing")%>">删除</a></td>
</tr>
<%rscp.movenext
next%>
</table>
<%rscp.close
set rscp=nothing
 cn.close
 set cn=nothing%>
