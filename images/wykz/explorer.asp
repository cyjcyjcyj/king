<!--#include file="checklogin.asp"-->
<%
dim page,act,pathurl,obj_fso,IsMSIE
dim obj_folder,s_folder,startnum,s_file
dim totalpage,currentsize,object_num,folder_sum

IsMSIE=msie()
totalpage=1
object_num=0
folder_sum=0
currentsize=0
page=trim(request.querystring("page"))
act=trim(request.querystring("act"))

if IsInteger(page) then
	page=cint(page)
	if page<1 then page=1
else
	page=1
end if

pathurl=Replace(Server.URLEncode(path),"+","%20")
set obj_fso=Server.CreateObject("Scripting.FileSystemObject")
if lcase(act)="paste" then
	call Paste(obj_fso)
end if
chksystem()
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=SystemName%> - 资源管理器</title>
<link rel="stylesheet" type="text/css" href="images/explorer.css" />
<script language="javascript">
<!--
var PathEncode="<%=pathurl%>";
var path="<%=path%>";
var page="<%=page%>";
var selfname="<%=selfname%>";
var LoginOneIP="<%=LoginOneIP%>";
var item_count=0;
setTimeout("window.location.reload()",1000*900);
//-->
</script>
<script type="text/javascript" src="images/common.js"></script>
<script type="text/javascript" src="images/function.js"></script>
</head>
<body>
<div id="bodymain">
<div id="header" class="LTRBorder">
	<div id="top1"><%call top1()%></div>
	<div id="top2"><%call address()%></div>
	<div id="top3"><%call topmsg()%></div>
</div>
<div id="content" class="LTRBorder">
<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
<%if obj_fso.folderexists(Server.MapPath(path)) then%>
    <td width="230" style="border-right:1px #CCCCCC solid;" valign="top">
	<div id="leftlan"><%call folders()%></div>
	</td>
    <td valign="top"><div id="rightlan"><%call files()%></div></td>
<%else%>
	<td align="center" style="color:#FF0000; height:100px;"><b>错误：目录</b>&nbsp;“<%=left(path,len(path)-1)%>”&nbsp;<b>不存在！</b></td>
<%end if%>
  </tr>
</table>
</div>
<div id="footer" class="LTRBorder">
<table id="foot_table" border="1" cellspacing="2" cellpadding="0">
  <tr>
    <td align="left">
	<span style="position:relative;top:1px; left:1px;"><%=object_num%>&nbsp;个对象（<%=folder_sum%>&nbsp;个文件夹&nbsp;&nbsp;&nbsp;<%=object_num-folder_sum%>&nbsp;个文件）</span>
	</td>
    <td width="80" align="left" valign="middle">
		<span style="position:relative;top:1px; left:1px;" title="当页文件总大小"><%=file_sizes(currentsize)%></span>
	</td>
    <td width="170" align="left">
	<span style="position:relative;top:0px; left:4px; height:16px;"><img align="absmiddle" src="images/mycom.gif"></span>
	<span style="position:relative;top:1px; left:1px;">我的站点</span>
	</td>
  </tr>
</table><%
endtime=timer()
if endtime<starttime then
	endtime=endtime+24*3600
end if
%><span style='position:relative;top:4px; text-align:center;line-height:120%;'>
<%=copyright%><br>Processed in <%=(endtime-starttime)*1000%> MSEL　<%=time()%>
</span></div><script language="javascript">setTimeout("remind()",1000);</script></div>
</body>
</html>
<%
sub top1()
%>
		<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
		  <tr>
			<td nowrap="nowrap" align="left" width="30%">&nbsp;<span style="position:relative"><a href="http://www.skymean.com" target="_blank" title="点击查看新版本"><%=SystemName &"&nbsp;"& SystemVersion%></a></span></td>
			<td nowrap="nowrap" align="center" width="31%">
				<img src="images/del.gif" class="imgbutton" onMouseOver="this.className='imgbt'" onMouseOut="this.className='imgbutton'" onClick="if(window.confirm('你真的要删除选定的文件或目录吗？')){document.explorer.submit();}" alt="删除选中文件">&nbsp;&nbsp;
				<img src="images/copy.gif" class="imgbutton" onMouseOver="this.className='imgbt'" onMouseOut="this.className='imgbutton'" onClick="copy('copy')" alt="复制选中文件">&nbsp;&nbsp;
				<img src="images/cut.gif" class="imgbutton" onMouseOver="this.className='imgbt'" onMouseOut="this.className='imgbutton'" onClick="copy('move')" alt="剪切选中文件">&nbsp;&nbsp;
				<%if session(SessionPrefix&"act")="copy" or session(SessionPrefix&"act")="move" then%>
				<img src="images/paste.gif" class="imgbutton" onMouseOver="this.className='imgbt'" onMouseOut="this.className='imgbutton'" onClick="window.location.href='<%=selfname%>?act=paste&path=<%=pathurl%>&page=<%=page%>';" alt="粘贴文件到当前目录">&nbsp;&nbsp;
				<%end if%>
				<img src="images/refresh.gif" class="imgbutton" onMouseOver="this.className='imgbt'" onMouseOut="this.className='imgbutton'" onClick="window.location.reload()" alt="刷新">
			</td>
			<td nowrap="nowrap" align="right" width="30%">
<%
	response.Write("<a target='_blank' href='spacesize.asp'>空间总计</a>&nbsp;")
	if loginstatus and popedom then
		response.write "<a target='_blank' href='adminuser.asp'>用户管理</a>&nbsp;"
		response.write "<a target='_blank' href='adminconfig.asp'>系统配置</a>&nbsp;"
		response.write "<a target='_blank' href='asptools.asp'>ASP工具</a>&nbsp;"
	else
		response.write "<a target='_blank' href='modifypwd.asp'>修改密码</a>&nbsp;"
	end if
	response.write "<a href='logout.asp'>注销登陆</a>"
%>
			&nbsp;</td>
		  </tr>
		</table>
<%
end sub


sub address()
%>
<form method="get" action="<%=selfname%>">
	&nbsp;<span style="position:relative;top:-7px;">位置：</span><span style="position:relative;top:<%if IsMSIE then response.Write("-6"):else response.Write("-9"):end if%>px;"><input type="text" name="path" style="width:665px;border:solid 1px #000099;" value="<%=path%>"></span>
	<input type="image" src="images/goaddress.gif" onMouseOver="this.src='images/goaddress2.gif'" onMouseOut="this.src='images/goaddress.gif'"><img src="images/goaddress2.gif" style="display:none;" width="0" height="0" />
</form>
<span style="display:none">
<iframe id="hideframe" name="hideframe" src="about:blank" frameborder=0 scrolling=no width="0" height="0" style="display:none"></iframe>
<form name="cmd_aff" action="" method="post" target="hideframe">
<span id="edit_item" style="position:relative"></span>
</form>
<input type="hidden" id="f_file" value="" /><input type="hidden" id="f_file_v" value="" />
<input type="hidden" id="alert_save" value="" />
</span>
<%
end sub

sub topmsg()
	response.Write("&nbsp;<span style='position:relative;top:2px;color:#FF0000;'>")
	if showlasttime=1 and len(trim(session(SessionPrefix&"lastdata"))&"")>0 then
		dim aler
		aler=split(trim(session(SessionPrefix&"lastdata")),"|")
		if lcase(aler(0))="first" then
			response.write "第一次登陆管理，欢迎！：）"
		else
			response.write "上次登陆IP：<a href='http://ip.coralqq.com/?q="&aler(0)&"' target=_blank title='点击在新窗口中打开IP查询'><font color='#FF0000'>"&aler(0)&"</font></a>　登陆时间："&aler(1)
		end if
	else
		response.Write("欢迎登陆管理！现在时间：" & now())
	end if
	response.Write("</span>")
end sub


sub folders()
	dim folder_name
%>
<table border="0" height="100%" cellspacing="0" cellpadding="0" style="width:230px;">
  <tr>
    <td colspan="2" height="30">
		<div style="position:relative;top:0px; width:230px;">
		<input type="button" value="全选目录" class="button" onClick="selectall(document.explorer.folders,true)"> 
		<input type="button" value="清除选择" class="button" onClick="selectall(document.explorer.folders,false)"> 
		<input type="button" value="新建目录" class="button" onClick="$('create_folder_s').style.display=$('create_folder_s').style.display==''?'none':'';">
		</div>
	</td>
  </tr>
	<tr id="create_folder_s" style="display:none">
	  <td align="right" colspan="2" style="border-bottom:1px #999999 solid;">
		<table width="100%" border="0" cellspacing="0" cellpadding="0">
		  <tr>
			<td align="right">新建文件夹名：</td>
			<td align="left"><input type="text" id="_new_folder_name" size="13" onkeypress="if(event.keyCode==13){create_file('md',$('_new_folder_name').value);}" /></td>
			<td align="left"><input type="button" value="确定" onclick="create_file('md',$('_new_folder_name').value)" class="secbutton" /></td>
		  </tr>
		</table>
	  </td>
	</tr>
	<tr>
	  <td align="left" colspan="2" height="23" style="border-bottom:1px #999999 solid;">&nbsp;<%=explorerop(path)%></td>
	</tr>
	<form name='explorer' method='post' action="cmd.asp?path=<%=pathurl%>&page=<%=page%>">
  	<input type='hidden' name='act' value='deltree'>
	<input type='hidden' name='frompath' value="<%=path%>">
<%
	set obj_folder=obj_fso.getfolder(Server.MapPath(path))
	object_num=obj_folder.files.count
	if object_num mod PerPageSize=0 then
		totalpage=object_num\PerPageSize
	else
		totalpage=object_num\PerPageSize+1
	end if
	if page<1 then
		page=1
	end if
	if page>totalpage then
		page=totalpage
	end if

	response.Write "<script language=""javascript"">" & vbcrlf
	for each s_folder in obj_folder.subfolders
		object_num=object_num+1
		folder_sum=folder_sum+1
		folder_name=s_folder.name

		response.write "sfolder("""&replace(folder_name,"'","\'")&""","""
		if not IsMSIE and len(folder_name)>22 then
			response.write replace(replace(cuted(folder_name,13),"'","\'"),chr(34),"\"&chr(34))
		end if
		response.write ""","""&replace(Server.URLEncode(folder_name),"+","%20")&""");"
	next
	response.write vbcrlf & "</script>"
%>
</table>
<%
end sub


sub files()
	dim file_name,file_type,sfile_size,i
%>
<table border="0" height="100%" cellspacing="0" cellpadding="0" style="width:540px">
  <tr>
    <td colspan="4" height="30">
		<div style="position:relative;top:0px; width:540px; text-align:right;">
		<input type="button" value="全选文件" class="button" onClick="selectall(document.explorer.files,true)">　
		<input type="button" value="清除选择" class="button" onClick="selectall(document.explorer.files,false)">　
		<input type="button" value="新建文件" class="button" onClick="$('create_file_s').style.display=$('create_file_s').style.display==''?'none':'';">　
		<input type="button" value="上传文件" class="button" onClick="window.open('<%=replace("upload.asp?path="&pathurl,"'","\'")%>','','')">&nbsp;
		</div>
	</td>
  </tr>
  <tr id="create_file_s" style="display:none">
    <td colspan="4" style="border-bottom:1px #999999 solid;">
		<table border="0" cellspacing="0" cellpadding="0" align="right">
		  <tr>
			<td align="right">新建文件名：</td>
			<td align="left"><input type="text" id="_new_file_name" size="20" onkeypress="if(event.keyCode==13){create_file('copycon',$('_new_file_name').value);}" />&nbsp;</td>
			<td align="left"><input type="button" value="确定" onclick="create_file('copycon',$('_new_file_name').value)" class="secbutton" />&nbsp;</td>
		  </tr>
		</table>
	</td>
  </tr>
  <tr>
    <td width="245" height="23" style="font-weight:bold" class="filetdb">文件名</td>
    <td width="62" style="font-weight:bold" class="filetdb">大小</td>
    <td width="115" style="font-weight:bold" class="filetdb">修改时间</td>
    <td width='118' style="font-weight:bold" class="filetdb2">操作</td>
  </tr>
<%
	i=1
	startnum=(page-1)*PerPageSize
	response.Write "<script language=""javascript"">" & vbcrlf
	response.Write("item_count=0;" & vbcrlf)
	for each s_file in obj_folder.files
		file_name=s_file.name
		if i>startnum then
			file_type=getselfname(s_file,".")
			sfile_size=s_file.size
			currentsize=currentsize+sfile_size

			response.write "files("""&replace(file_name,"'","\'")&""","""
			if IsMSIE then
				if len(file_name)>72 then
					response.Write replace(replace(cuted(file_name,50),"'","\'"),chr(34),"\"&chr(34))
				end if
			else
				if len(file_name)>32 then
					response.Write replace(replace(cuted(file_name,20),"'","\'"),chr(34),"\"&chr(34))
				end if
			end if
			response.write ""","""&replace(Server.URLEncode(file_name),"+","%20")&""","""&replace(icon(file_type),"'","\'")&""","""&file_sizes(sfile_size)&""","""&s_file.DateLastModified&""","
			if instr("|"&lcase(trim(CanNotEditFileType))&"|","|"&file_type&"|")>0 then
				response.write "0"
			elseif instr("|"&lcase(trim(CanEditFileType))&"|","|"&file_type&"|")>0 then
				response.write "1"
			else
				response.write "2"
			end if
			response.write ");"
		end if
		if i>startnum+PerPageSize then
			exit for
		end if
		i=i+1
	next
	response.write vbcrlf & "</script>" & vbcrlf

	response.Write "<tr><td colspan='4' height='10'></td></tr>" & vbcrlf
	Response.Write "<td colspan='4' align='center' style='font-weight:bold;border-top:1px #999999 solid;'>" & VbCrlf
	if page>1 then
		response.write "<a href='"&selfname&"?path="&pathurl&"&page="&(page-1)&"'>上一页</a>&nbsp;&nbsp;"
	end if
	response.write "共"&obj_folder.files.count&"个文件&nbsp;&nbsp;当前第&nbsp;"
	response.Write "<span style='position:relative;top:3px;'>"
	response.write "<select id='selectpage' onchange=""window.location.href='"&selfname&"?page='+(this.options.selectedIndex+1)+'&path="&replace(pathurl,"'","\'")&"';this.disabled=true;"">" & vbcrlf
	for i=1 to totalpage
		if i=page then
			response.write "<option selected>"&i&"</option>" & vbcrlf
		else
			response.write "<option>"&i&"</option>" & vbcrlf
		end if
	next
	response.write "</select></span>&nbsp;页" & vbcrlf
	response.write "&nbsp;&nbsp;共&nbsp;"&totalpage&"&nbsp;页&nbsp;&nbsp;每页"&PerPageSize&"个文件"
	if page<totalpage then
		response.write "&nbsp;&nbsp;<a href='"&selfname&"?path="&pathurl&"&page="&(page+1)&"'>下一页</a>"
	end if
	response.write "</td>" & vbcrlf
	response.write "</tr>" & vbcrlf
	response.Write "<tr><td colspan='4' height='10'></td></tr>" & vbcrlf
%>
	</form>
</table>
<%
end sub


sub Paste(fsoobj)
	dim SplitFolders,SplitFiles,i,SourceF,TargetF,CurrentFolder,cmd
	CurrentFolder=trim(Server.MapPath(path))
	cmd=lcase(trim(session(SessionPrefix&"act")))
	if not IsObject(fsoobj) then
		call showmsg("错误：FSO对象参数传送出错！无法执行粘贴操作！","")
		response.end
	end if

	if lcase(trim(session(SessionPrefix&"frompath")))<>lcase(CurrentFolder) and fsoobj.FolderExists(CurrentFolder) then
		if cmd="copy" or cmd="move" then
			if session(SessionPrefix&"folders")<>"" then
				SplitFolders=split(session(SessionPrefix&"folders"),"||")
				for i=0 to ubound(SplitFolders)
					SourceF=session(SessionPrefix&"frompath")&"\"&SplitFolders(i)
					TargetF=CurrentFolder&"\"&SplitFolders(i)
					if fsoobj.FolderExists(SourceF) then
						fsoobj.CopyFolder SourceF,TargetF,true
						if cmd="move" then
							fsoobj.DeleteFolder SourceF,true
						end if
					end if
				next
			end if
			if session(SessionPrefix&"files")<>"" then
				SplitFiles=split(session(SessionPrefix&"files"),"||")
				for i=0 to ubound(SplitFiles)
					SourceF=session(SessionPrefix&"frompath")&"\"&SplitFiles(i)
					TargetF=CurrentFolder&"\"&SplitFiles(i)
					if fsoobj.FileExists(SourceF) then
						fsoobj.CopyFile SourceF,TargetF,true
						if cmd="move" then
							fsoobj.DeleteFile SourceF,true
						end if
					end if
				next
			end if
			if cmd="move" then
				session(SessionPrefix&"act")=""
				session(SessionPrefix&"frompath")=""
				session(SessionPrefix&"folders")=""
				session(SessionPrefix&"files")=""
			end if
		end if
	end if
end sub

function icon(ext)
	dim iconimg
	select case ext
		case "htm","html"
			iconimg="html.gif"
		case "css","ini","inf"
			iconimg="css.gif"
		case "js","vbs","vbe"
			iconimg="js.gif"
		case "exe"
			iconimg="exe.gif"
		case "bat","cmd"
			iconimg="bat.gif"
		case "pdf"
			iconimg="pdf.gif"
		case "ppt"
			iconimg="ppt.gif"
		case "swf"
			iconimg="swf.gif"
		case "xls"
			iconimg="xls.gif"
		case "asp","asa"
			iconimg="asp.gif"
		case "aspx"
			iconimg="aspx.gif"
		case "mht","mhtml"
			iconimg="mht.gif"
		case "txt","log"
			iconimg="text.gif"
		case "jpg","png"
			iconimg="jpg.gif"
		case "bmp"
			iconimg="bmp.gif"
		case "gif"
			iconimg="gif.gif"
		case "mdb"
			iconimg="mdb.gif"
		case "doc"
			iconimg="word.gif"
		case "mid","midi"
			iconimg="midi.gif"
		case "wav","ram"
			iconimg="wav.gif"
		case "mp3"
			iconimg="mp3.gif"
		case "avi","rm","mp","mpg","mpeg","mpe","rmvb"
			iconimg="wmp.gif"
		case "zip"
			iconimg="zip.gif"
		case "rar"
			iconimg="rar.gif"
		case "dll","sys"
			iconimg="dll.gif"
		case "hlp"
			iconimg="hlp.gif"
		case "reg","key"
			iconimg="reg.gif"
		case "chm"
			iconimg="chm.gif"
		case "htc"
			iconimg="htc.gif"
		case "url"
			iconimg="url.gif"
		case "lnk"
			iconimg="lnk.gif"
		case "sql"
			iconimg="sql.gif"
		case "xml"
			iconimg="xml.gif"
		case "xslt"
			iconimg="xslt.gif"
		case "config"
			iconimg="vs.gif"
		case "cs"
			iconimg="cs.gif"
		case "inc","lst"
			iconimg="inc.gif"
		case "iso","bin","mds"
			iconimg="iso.gif"
		case "mdf"
			iconimg="mdf.gif"
		case else
			iconimg="unknown.gif"
	end select
	icon=iconimg
end function

function file_sizes(filesizes)
	if filesizes<1024 then
		file_sizes=filesizes&" Bytes"
	elseif filesizes<1048576 then
		file_sizes=round(filesizes/1024,2)&" KB"
	else
		file_sizes=round(filesizes/1048576,3)&" MB"
	end if
end function

function explorerop(folderpath)
	explorerop="<a href='javascript:history.back()' onfocus='if(this.blur) this.blur()'><img align='absmiddle' src='images/back.gif' title='后退'></a>&nbsp;&nbsp;" & vbcrlf
	explorerop=explorerop&"<a href='javascript:history.forward()' onfocus='if(this.blur) this.blur()'><img align='absmiddle' src='images/forward.gif' title='前进'></a>&nbsp;&nbsp;" & vbcrlf
		
	if folderpath="/" then
		explorerop=explorerop&"<img align='absmiddle' src='images/noup.gif' title='此乃根目录'>"
	elseif lcase(folderpath)=lcase(session(SessionPrefix&"adminpath")) then
		explorerop=explorerop&"<img align='absmiddle' src='images/noup.gif' title='只能管理到此目录'>"
	else
		explorerop=explorerop&"<a href='"&selfname&"?path="&server.urlencode(left(folderpath,instrrev(folderpath,"/",len(folderpath)-1)))&"'><img align='absmiddle' src='images/up.gif' title='向上'></a>"
	end if
end function
%>