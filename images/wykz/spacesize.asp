<!--#include file="checklogin.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<title><%=SystemName%> - 空间总计系统</title>
<style type="text/css">
a{color:#000080;text-decoration:none;}
a:hover{color:#ff3333;text-decoration:underline}
body{
	text-align:center;
	font-size:12px;
	color:#000090;
	background:#ffffff url(images/bgbrick.gif);
}
td{
	height:25px;
	vertical-align:middle;
	word-break:break-all;
	border-top:1px #000000 solid;
	border-right:1px #000000 solid;
}
div{
	margin-right: auto;
	margin-left: auto;
}
.ltd{
	border-left:1px #000000 solid;
}
#footer{
	background-color:#FFFFFF;
	width:758px;
	border:1px #000000 solid;
	text-align:center;
}
</style>
</head>
<body topmargin="0" leftmargin="0">
<center><br><b><%=SystemName%> - 空间总计系统</b><br><br></center>
<table width="760" border="0" cellspacing="0" cellpadding=0 align="center" bgcolor="#FFFFFF">
  <tr>
    <td nowrap="nowrap" class="ltd"><b>目录</b></td>
    <td nowrap="nowrap"><b>大小</b></td>
  </tr>
<%
	dim obj_fso,obj_folder,folders,s_size,s_tsize,s_folder
	if trim(PathCanModify)<>"" then
		PathCanModify=replace(PathCanModify,"\","/")
	end if
	s_tsize=0
	PathCanModify=Server.MapPath(PathCanModify)
	set obj_fso=server.createobject("scripting.filesystemobject")
	if obj_fso.FolderExists(PathCanModify) then
		set obj_folder=obj_fso.getfolder(PathCanModify)
		for each s_folder in obj_folder.subfolders
			folders=s_folder.name
			folders=PathCanModify&"\"&folders
			s_size=spacesize(folders)
%>
  <tr align="left" onMouseOver="javascript:this.bgColor='#EFEFEF';" onMouseOut="javascript:this.bgColor='';">
    <td nowrap="nowrap" class="ltd">&nbsp;<a href="explorer.asp?path=<%=Server.URLEncode(replace(replace(folders&"\",Server.MapPath("/"),""),"\","/"))%>" target="_blank"><%=folders%></a></td>
    <td nowrap="nowrap">&nbsp;<%
			if s_size=-1 then
				response.Write("不存在")
			else
				s_tsize=(s_tsize+s_size)
				response.Write(formatnum(s_size))
			end if
%></td>
  </tr>
<%
		next
		set obj_folder=nothing
	end if
	set obj_fso=nothing
	chksystem()
%>
  <tr align="left" onMouseOver="javascript:this.bgColor='#EFEFEF';" onMouseOut="javascript:this.bgColor='';">
    <td nowrap="nowrap" class="ltd">&nbsp;<a href="explorer.asp?path=<%=Server.URLEncode(replace(replace(PathCanModify&"\",Server.MapPath("/"),""),"\","/"))%>" target="_blank"><%=PathCanModify%>(不含子目录)</a></td>
    <td nowrap="nowrap">&nbsp;<%
	dim total
	total=spacesize(PathCanModify)
	if total=-1 then
		response.Write("不存在")
	else
		response.Write(formatnum(total-s_tsize))
	end if
%></td>
  </tr>
  <tr align="left" onMouseOver="javascript:this.bgColor='#EFEFEF';" onMouseOut="javascript:this.bgColor='';">
    <td nowrap="nowrap" class="ltd">&nbsp;<a href="explorer.asp?path=<%=Server.URLEncode(replace(replace(PathCanModify&"\",Server.MapPath("/"),""),"\","/"))%>" target="_blank"><%=PathCanModify%>(总大小)</a></td>
    <td nowrap="nowrap">&nbsp;<%=formatnum(total)%></td>
  </tr>
</table>
<div id="footer"><br><%
endtime=timer()
if endtime<starttime then
	endtime=endtime+24*3600
end if
%><span style='position:relative;top:4px; text-align:center;line-height:120%;'>
<%=copyright%><br>Processed in <%=(endtime-starttime)*1000%> MSEL
</span><br><br></div>
</body>
</html>