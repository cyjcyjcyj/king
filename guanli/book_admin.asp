<!--#include file="conn.asp"-->
<style type="text/css">
<!--
body,td,th {
	font-size: 12px;
}
.style1 {	color: #DBBF90;
	font-weight: bold;
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
<%
function lleft(content,lef)
for le=1 to len(content)
if asc(mid(content,le,1))<0 then
lef=lef-2
else
lef=lef-1
end if
if lef<=0 then exit for
next
lleft=left(content,le)
end function
%>

<HTML>
<center>

<script language="JavaScript" type="text/JavaScript">
function CheckAll(form)  {
  for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.name != 'chkall')       e.checked = form.chkall.checked; 
   }
  }
</script>
</head>

<table width="760" border=0 cellspacing=0 cellpadding=0 align=center class="grayline">
<tr><td align=center bgcolor="#E8F1FF">
<br>


<%

action=request("action")

'管理首页
if action="" then%>

<table width="100%" border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse;">
	<form name=book action=book_admin.asp method=post><tr>	
	<td align=center width=5%>选</td>
	<td align=center width=10% height=15>姓名</td> 
	<td align=center width=46%>内容（编辑与回复）</td>
	<td align=center width=19%>日期</td>
	<td align=center width=11%>状态</td>
	<td align=center width=9%>审核</td>
	</tr>
<%
dim rs,msg_per_page
dim sql
msg_per_page = 10 '每页显示记录数
Set rs=Server.CreateObject("ADODB.Recordset")
sql="select * from Feedback where del=false order by PostDate desc"

rs.cursorlocation = 3 
rs.pagesize = msg_per_page '每页显示记录数
rs.open sql,cn,1,1 

	if rs.eof and rs.bof then
	response.write "<tr><td colspan=6 align=center><BR>暂时没有留言<BR><BR></td></tr>"
	end if

	if not (rs.eof and rs.bof) then '检测记录集是否为空
		totalrec = RS.RecordCount '总记录条数
		if rs.recordcount mod msg_per_page = 0 then '计算总页数,recordcount:数据的总记录数
		n = rs.recordcount\msg_per_page 'n:总页数
		else 
		n = rs.recordcount\msg_per_page+1 
		end if 

		currentpage = request("page") 'currentpage:当前页
		If currentpage <> "" then
			currentpage = cint(currentpage)
			if currentpage < 1 then 
				currentpage = 1
			end if 
			if err.number <> 0 then 
				err.clear
				currentpage = 1
			end if
		else
			currentpage = 1
		End if 
		if currentpage*msg_per_page > totalrec and not((currentpage-1)*msg_per_page < totalrec)then 
			currentPage=1
		end if
		rs.absolutepage = currentpage 'absolutepage：设置指针指向某页开头
		rowcount = rs.pagesize 'pagesize：设置每一页的数据记录数
		dim i
		dim k

		Do while not rs.eof and rowcount>0
	content=rs("Comments")	
	replay=rs("replay")
	UserName=rs("UserName")

			Response.write "<tr><td><input type='checkbox' value='"&rs("ID")&"' name=id></td><td>"&UserName&"</td><td><a href='book_admin.asp?action=replay&id="&rs("ID")&"'>"
			response.write lleft(server.htmlencode(content),50)
			response.write "</a></td><td  align=center>"&rs("Postdate")&"</td><td  align=center>"
			if Isnull(Replay) then
				response.write "<font color=red>新留言</font>"
			else				
				response.write "已回复"
			end if
				response.write "</td><td align=center>"
				if rs("Online")="0" then response.write "<font color=red>隐藏</font>" else response.write "公开" 	end if
				response.write "</td></tr>"
			rowcount=rowcount-1
			rs.movenext		
		loop
	end if

	rs.close
	cn.close
	set rs=nothing
	set cn=nothing
%>
<tr><td colspan=6><input type='checkbox' name=chkall onclick='CheckAll(this.form)'> 全选 
	<input type="submit" name="action" value="删除" onclick="{if(confirm('该操作不可恢复！\n\n确实删除选定的留言？')){this.document.Prodlist.submit();return true;}return false;}"> 	
	</td></tr></form></table>
<%
call listPages()
end if

if request("action")="删除" then
	delid=replace(request("id"),"'","")
	call delfeedback()
end if

if request("action")="replay" then
	id=cint(trim(request("id")))
	call detailfeedback()
end if

if request("action")="setup" then
	call setup()
end if




%>                     
</td></tr>
<tr><td>&nbsp;</td></tr>
</table>   
</td></tr>
</table>
</body>

<%
sub setup()
Set rs=Server.CreateObject("ADODB.Recordset")
sql="select * from book_setup"
rs.open sql,cn,1,3

if request("save")="ok" then

	if trim(request.form("sitename"))="" or trim(request.form("admin"))="" or trim(request.form("maxlength"))="" or trim(request.form("pages"))="" then
	response.write "<script language='javascript'>"
	response.write "alert('资料填写不完整或有错误，请检查后重新提交。');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if

	if (not isNumeric(request.form("maxlength"))) or (not isNumeric(request.form("pages"))) then
	response.write "<script language='javascript'>"
	response.write "alert('在要求填写数字的地方，您填写了非数字内容，请检查您填写的内容是否符合要求。');"
	response.write "location.href='javascript:history.go(-1)';"
	response.write "</script>"
	response.end
	end if


rs("sitename")=request("sitename")
rs("admin")=request("admin")
if trim(request("password"))<>"" then rs("password")=md5(trim(request("password")))
rs("view")=request("view")
rs("maxlength")=request("maxlength")
rs("pages")=request("pages")
rs("html")=request("html")
rs.update

	response.write "<script language='javascript'>"
	response.write "alert('您提交的设置已保存。');"
	response.write "location.href='book_admin.asp?action=setup';"
	response.write "</script>"
	response.end

else
%>
<%
end if
rs.close
set rs=nothing
end sub



sub delfeedback()
	if delid="" or isnull(delid) then

	response.write "<script language='javascript'>"
	response.write "alert('操作失败，没有选择合适参数，请单击“确定”返回！');"
	response.write "location.href='book_admin.asp';"
	response.write "</script>"
	response.end

	else
		cn.execute("delete from Feedback where ID in ("&delid&")")
		cn.close
		set cn=nothing

	response.write "<script language='javascript'>"
	response.write "alert('留言删除成功，请单击“确定”返回！');"
	response.write "location.href='book_admin.asp';"
	response.write "</script>"
	response.end

	end if
end sub

sub detailfeedback()
if id="" then
	response.write "<script language='javascript'>"
	response.write "alert('无此留言编号，请单击“确定”返回！');"
	response.write "location.href='book_admin.asp';"
	response.write "</script>"
	response.end
end if

	'修改留言资料
if request("send")="ok" then
	set rs=server.createobject("adodb.recordset")
	sql = " select * from feedback where del=false and ID="&id
	rs.open sql,cn,1,3

		if not (rs.eof and rs.bof) then
		rs("comments")=request.form("comments")
		rs("Replay")=replace(request.form("Replay"),vbCRLF,"<BR>")
		rs("ReplayDate") = Now()
		rs("Online")=request("Online")
		rs.update
		end if

	rs.close

	response.write "<script language='javascript'>"
	response.write "alert('留言已经修改或回复成功，请单击“确定”返回！');"
	response.write "location.href='book_admin.asp';"
	response.write "</script>"
	response.end
end if

	'显示详细资料
	set rs = server.createobject("adodb.recordset")
	sql = "select * from feedback where ID="&id
	rs.open sql,cn,1,1

		if rs.eof and rs.bof then 
		response.write "<script language='javascript'>"
		response.write "alert('无此留言，请单击“确定”返回！');"
		response.write "location.href='book_admin.asp';"
		response.write "</script>"
		response.end
		end if

		if not (rs.eof and rs.bof) then 
			Comments=replace(rs("Comments"),"<BR>",vbCRLF)
			if rs("replay")<>"" then replay=replace(rs("Replay"),"<BR>",vbCRLF) else repley=""  end if

		%>

   <table width="600" border="1" align="center" cellpadding="3" bordercolor="#333333" style="border-collapse: collapse;">
		 <form name="repl" method="post" action='book_admin.asp?action=replay&id=<%=id%>'>
		 <tr><TD align="right" width=20% height=15>留言者IP地址</TD><td><%=rs("IP")%></td></tr>
		 <tr><TD align="right" width=20%>留言日期</TD><td><%=rs("PostDate")%></td></tr>		 
		 <tr><TD align="right" width=20%>留言人姓名</TD><td><%=rs("UserName")%>&nbsp;</td></tr>
		<tr><TD align="right" width=20%>留言邮箱</TD><td><%=rs("UserMail")%>&nbsp;</td></tr>
		<tr><TD align="right" width=20%>留言网址</TD><td><%=rs("url")%>&nbsp;</td></tr>
		<tr><TD align="right" width=20%>其它联系方式</TD><td><%=rs("qq")%>&nbsp;</td></tr>
		 <tr><TD align="right" width=20%>内容</TD><td><textarea style="overflow:auto" name="comments" cols="60" rows="8"><%=Comments%></textarea></td></tr>
		 <tr><TD align="right" width=20% valign=top>回复内容</TD><td><textarea style="overflow:auto" name="Replay" cols="60" rows="8"><%=replay%></textarea>&nbsp;</td></tr>
		<tr><TD align="right" width=20%>是否隐藏</TD><td><input type="radio" name="Online" value="0" <%if rs("Online")="0" then%>checked<%end if%>>
			  隐藏<input type="radio" name="Online" value="1" <%if rs("Online")="1" then%>checked<%end if%>>
			  公开 </td></tr>
			<TR><TD align="right" width=20%>&nbsp;<INPUT TYPE="hidden" name=send value=ok></TD><TD>
				<input type="submit" name="action" value=" 提 交 "></TD></TR>
</form></TABLE>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<%
		end if	

	rs.close
	set rs=nothing

end sub


'分页
sub listPages() 
if n <= 1 then exit sub 
%>
共<%=totalrec%>条留言 
<%if currentpage = 1 then%>
<font color=darkgray>首页 前页</font>
<%else%> 
<a href="<%=request.ServerVariables("script_name")%>?page=1">
首页</font></a> <a href="<%=request.ServerVariables("script_name")%>?page=<%=currentpage-1%>">前页</a>
<%end if%>
<%if currentpage = n then%> 
<font color=darkgray >下页 末页</font>
<%else%> 
<a href="<%=request.ServerVariables("script_name")%>?page=<%=currentpage+1%>">下页</a> <a href="<%=request.ServerVariables("script_name")%>?page=<%=n%>">末页</a>
<%end if%>
  第<%=currentpage%>页 共<%=n%>页
<%end sub%>
