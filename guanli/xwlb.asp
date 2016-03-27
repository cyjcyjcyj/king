<!--#include file="lianjie.asp"----->
<%xwlb_id=request("xwlb_id")
set rsgly=server.CreateObject("adodb.recordset")
rsgly.open"select * from xw where xwlb_id="&xwlb_id,cn,1,3
set rsall=server.CreateObject("adodb.recordset")
rsall.open"select * from xwlb",cn,1,3
dim pages,cur_recordcounts
pages=clng(request.QueryString("pages"))
rsgly.pagesize=15
if not rsgly.eof then
 if pages=0 then
 pages=1
  else
  rsgly.move(pages-1)*rsgly.pagesize
  end if
 if pages<rsgly.pagecount then
  cur_recordcounts=rsgly.pagesize
  else
   cur_recordcounts=rsgly.recordcount-(pages-1)*rsgly.pagesize
   end if
end if
%>
<title>管理页面</title>
<META http-equiv=Content-Type content=text/html;charset=gb2312>

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
<script language="JavaScript" type="text/JavaScript">
function CheckAll(form)  {
  for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.name != 'chkall')       e.checked = form.chkall.checked; 
   }
  }
</script>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="5" marginheight="0" marginwidth="0" bgcolor="#DDEEFF">
<form action="xwlb.asp">
<input name="xwlb_id" type="hidden" value="<%=xwlb_id%>">
<table class="tableBorder" width="96%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td  align="center" background="images/admin_bg_1.gif" height="25"><span class="style1">新闻列表</span></td>
  </tr>
  <tr>
    <td height="20"  align="center" bgcolor="#E8F1FF"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td align="center"><%for k=1 to rsall.recordcount%>
            <span class="style1"><a href=xwlb.asp?xwlb_id=<%=rsall("id")%>><%=rsall("xwlb")%></a></span>&nbsp;&nbsp;&nbsp;&nbsp;
            <%rsall.movenext
		  next%></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td valign="top">    <table width="100%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
      <tr bgcolor="#E8F1FF">
        <td align="center"><span class="TableRow2">新闻标题</span><span class="TableRow2"> </span></td>
        <td align="center">新闻类别</td>
        <td align="center"><span class="BodyTitle">添加时间</span></td>
        <td align="center"><span class="BodyTitle">修改</span></td>
        <td align="center"><span class="BodyTitle">选</span></td>
      </tr>
      <%for i=1 to cur_recordcounts
	  set rsxwlb=server.CreateObject("adodb.recordset")
	  rsxwlb.open"select * from xwlb where id="&rsgly("xwlb_id"),cn,1,3%>
      <tr align="center" bgcolor="#E8F1FF">
        <td height="15" align="left"><span class="TableRow2"><%=rsgly("xwbt")%></span></td>
        <td height="15"><span class="TableRow2"><%=rsxwlb("xwlb")%></span></td>
        <td height="15"><span class="TableRow1"><%=rsgly("tjsj")%></span></td>
        <td height="15"><span class="TableRow1"><a href="xwxiugai.asp?id=<%=rsgly("id")%>">修改</a></span></td>
        <td height="15"><span class="TableRow1">
          <input type="checkbox" value="<%=rsgly("id")%>" name=id>
</span></td>
      </tr>
      <%rsgly.movenext
next%>
    </table>
      <table width="96%"  border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td align="right"><input type='checkbox' name=chkall onclick='CheckAll(this.form)'>
      全选
        <input type="submit" name="action" value="删除" onclick="{if(confirm('该操作不可恢复！\n\n确实删除选定的项目？')){this.document.Prodlist.submit();return true;}return false;}"></td>
        </tr>
      </table></td>
  </tr>
</table>
</form>
<table width="96%"  border="0" align="center" cellpadding="0" cellspacing="8">
  <tr>
    <td align="center"><% 
	if pages>1 then
	response.Write"<a href=xwlb.asp?pages="&pages-1&"&xwlb_id="&xwlb_id&">前一页</a>&nbsp;"
	else
	response.Write"前一页&nbsp;"
	end if
	if pages<=rsgly.pagecount-5 then
	start=pages-5
	else
	start=rsgly.pagecount-10
	end if
	if start<=0 then start=1
	for i=start to rsgly.pagecount 
	 if i=pages then
	 response.Write"&nbsp;"&i 
	 else 
	  if i<start+11 then
	  response.Write"&nbsp;<a href=xwlb.asp?pages="&i&"&xwlb_id="&xwlb_id&">"&i&"</a>"
	  else
	  exit for
	 end if
	 end if
	 next 
	 if pages<rsgly.pagecount then
	 response.Write"&nbsp;<a href=xwlb.asp?pages="&pages+1&"&xwlb_id="&xwlb_id&">下一页</a>"
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
	response.write "location.href='xwlb.asp?xwlb_id="&xwlb_id&"';"
	response.write "</script>"
	response.end

	else
		cn.execute("delete from xw where ID in ("&delid&")")
		cn.close
		set cn=nothing

	response.write "<script language='javascript'>"
	response.write "alert('项目删除成功，请单击“确定”返回！');"
	response.write "location.href='xwlb.asp?xwlb_id="&xwlb_id&"';"
	response.write "</script>"
	response.end

	end if
end sub

sub detailfeedback()
if id="" then
	response.write "<script language='javascript'>"
	response.write "alert('无此留言编号，请单击“确定”返回！');"
	response.write "location.href='xwlb.asp?xwlb_id="&xwlb_id&"';"
	response.write "</script>"
	response.end
end if
end sub
%>
