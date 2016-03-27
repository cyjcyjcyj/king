<!--#include file="lianjie.asp"----->
<%dlb_id=request("dlb_id")
 xlb_id=request("xlb_id")
   set rsdlb=server.CreateObject("adodb.recordset")
  rsdlb.open"select * from dlb",cn,1,3
if dlb_id="" and xlb_id="" then
set rscp=server.CreateObject("adodb.recordset")
rscp.open"select * from jhcp",cn,1,3
else
  if dlb_id<>"" and xlb_id="" then
  set rscp=server.CreateObject("adodb.recordset")
  rscp.open"select * from jhcp where dlb_id="&dlb_id,cn,1,3
   set rsxlb=server.CreateObject("adodb.recordset")
  rsxlb.open"select * from jhxlb where dlb_id="&dlb_id,cn,1,3
  set rsdlbmc=server.CreateObject("adodb.recordset")
  rsdlbmc.open"select * from dlb where id="&dlb_id,cn,1,3
  end if
  if dlb_id<>"" and xlb_id<>"" then
  set rscp=server.CreateObject("adodb.recordset")
  rscp.open"select * from jhcp where dlb_id="&dlb_id&" and xlb_id="&xlb_id,cn,1,3
  set rsxlb=server.CreateObject("adodb.recordset")
  rsxlb.open"select * from jhxlb where dlb_id="&dlb_id,cn,1,3
  set rsdlbmc=server.CreateObject("adodb.recordset")
  rsdlbmc.open"select * from dlb where id="&dlb_id,cn,1,3
  end if
 end if
dim pages,cur_recordcounts
pages=clng(request.QueryString("pages"))
rscp.pagesize=15
if not rscp.eof then
 if pages=0 then
 pages=1
  else
  rscp.move(pages-1)*rscp.pagesize
  end if
 if pages<rscp.pagecount then
  cur_recordcounts=rscp.pagesize
  else
   cur_recordcounts=rscp.recordcount-(pages-1)*rscp.pagesize
   end if
end if
%>

<title>管理页面</title><script language="JavaScript" type="text/JavaScript">
function CheckAll(form)  {
  for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.name != 'chkall')       e.checked = form.chkall.checked; 
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
<form action="dckt_list.asp?dlb_id=<%=dlb_id%>&xlb_id=<%=xlb_id%>" method="post"><table width="96%" border="0" align=center cellpadding="2" cellspacing="1" bgcolor="#FFFFFF" class="tableBorder">
<tr valign="middle" background="images/admin_bg_1.gif">
<th height=22 colspan=6 background="images/admin_bg_1.gif">
<%for i=1 to rsdlb.recordcount %>
<a href="dckt_list.asp?dlb_id=<%=rsdlb("id")%>"><%=rsdlb("dlb")%></a>&nbsp;
 <%rsdlb.movenext 
 next%> </th>
</tr>
<tr align="center" valign="middle" background="images/admin_bg_1.gif">
  <td height=24 colspan=6 background="images/admin_bg_1.gif"><%if dlb_id="" then %>
  所有项目列表
    <%else%>
 [ <b><%=rsdlbmc("dlb")%></b> ]&nbsp;&nbsp;&nbsp;
  <%for j=1 to rsxlb.recordcount 
%><a href="dckt_list.asp?xlb_id=<%=rsxlb("id")%>&dlb_id=<%=dlb_id%>"><%=rsxlb("xlb")%></a>
<%
  rsxlb.movenext
  next%>
  <%end if%></th>
</tr>
<tr align="center" bgcolor="#DEEFFF">
<td width="25%" height=25 class=BodyTitle>项目标题</td>

<td width="27%" height=25 class=BodyTitle>添加时间</td>
<td width="8%" class=BodyTitle>修改</td>
<td width="7%" class=BodyTitle>删除</td>
<td width="8%" class=BodyTitle>选</td>
</tr><%for i=1 to cur_recordcounts%>
<tr align="center" bgcolor="#DEEFFF">
<td height=23  class="TableRow2"><%=rscp("cpmc")%></td>
<td class="TableRow1"><%=rscp("tjsj")%></td>
<td class="TableRow1"><a href="dcktxiugai.asp?id=<%=rscp("id")%>">修改</a></td>
<td class="TableRow1"><a href="dckt_del.asp?id=<%=rscp("id")%>">删除</a></td>
<td class="TableRow1">  <input type="checkbox" value="<%=rscp("id")%>" name=id></td>
</tr>
<%rscp.movenext
next%>
</table> <table width="96%"  border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td align="right"><input type='checkbox' name=chkall onclick='CheckAll(this.form)'>
      全选
        <input type="submit" name="action" value="删除" onClick="{if(confirm('该操作不可恢复！\n\n确实删除选定的项目？')){this.document.Prodlist.submit();return true;}return false;}"></td>
        </tr>
      </table>
</form>
<table width="96%"  border="0" align="center" cellpadding="0" cellspacing="8">
  <tr>
    <td align="center"><% 
	if pages>1 then
	response.Write"<a href=dckt_list.asp?pages="&pages-1&"&dlb_id="&dlb_id&">前一页</a>&nbsp;"
	else
	response.Write"前一页&nbsp;"
	end if
	if pages<=rscp.pagecount-5 then
	start=pages-5
	else
	start=rscp.pagecount-10
	end if
	if start<=0 then start=1
	for i=start to rscp.pagecount 
	 if i=pages then
	 response.Write"&nbsp;"&i 
	 else 
	  if i<start+11 then
	  response.Write"&nbsp;<a href=dckt_list.asp?pages="&i&"&dlb_id="&dlb_id&">"&i&"</a>"
	  else
	  exit for
	 end if
	 end if
	 next 
	 if pages<rscp.pagecount then
	 response.Write"&nbsp;<a href=dckt_list.asp?pages="&pages+1&"&dlb_id="&dlb_id&">下一页</a>"
	 else
	 response.Write"&nbsp;下一页"
	 end if
	 %></td>
  </tr>
</table>


<%

if request("action")="删除" then
	delid=replace(request("id"),"'","")
	call delfeedback()
end if
sub delfeedback()
	if delid="" or isnull(delid) then

	response.write "<script language='javascript'>"
	response.write "alert('操作失败，没有选择合适参数，请单击“确定”返回！');"
	response.write "location.href='dckt_list.asp?dlb_id="&dlb_id&"&xlb_id="&xlb_id&"';"
	response.write "</script>"
	response.end

	else
		cn.execute("delete from jhcp where ID in ("&delid&")")
		cn.close
		set cn=nothing
    response.write "<script language='javascript'>"
	response.write "alert('项目删除成功，请单击“确定”返回！');"
	response.write "location.href='dckt_list.asp?dlb_id="&dlb_id&"&xlb_id="&xlb_id&"';"
	response.write "</script>"
	response.end

	end if
end sub

sub detailfeedback()
if id="" then
	response.write "<script language='javascript'>"
	response.write "alert('无此留言编号，请单击“确定”返回！');"
	response.write "location.href='dckt_list.asp?dlb_id="&dlb_id&"&xlb_id="&xlb_id&"';"
	response.write "</script>"
	response.end
end if
end sub
%>

<%rscp.close
set rscp=nothing
 cn.close
 set cn=nothing%>


