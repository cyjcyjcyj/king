<!--#include file="checklogin.asp"-->
<%
dim filepath,obj_fso,obj_file,s_content,fcontent,code,adodb,edit_filename
code=request("code")
adodb=lcase(trim(request("adodb")))
if trim(code)="" then
	code="GB2312"
end if
chksystem()
path=trim(Request.Form("path"))&""
if path="" then
	path=trim(Request.QueryString("path"))&""
end if
if right(path,1)<>"/" then
	path=path&"/"
end if
if left(path,1)<>"/" then
	path="/"&path
end if
if InvalidChar(path,0) then
	call showmsg("错误：指定的路径含有非法字符！","")
	response.end
end if
Call CheckPath(PathCanModify,path)
if right(path,1)="/" then
	path=left(path,len(path)-1)
end if

if request.form("act")="editfile" then
	fcontent=request.form("fcontent")
	filepath=Server.MapPath(path)
	set obj_fso=Server.CreateObject("Scripting.FileSystemObject")
	if not obj_fso.FileExists(filepath) then
		obj_fso.CreateTextFile(filepath)
	end if
	if trim(request.Form("code"))<>"" then
		if not SaveToFile(fcontent,filepath,code) then
			call showmsg("ADODB.Stream保存文件失败！","")
			set obj_fso=nothing
			response.end
		end if
	else
		set obj_file=obj_fso.OpenTextFile(filepath,2,false)
		obj_file.write fcontent
		obj_file.close
		set obj_file=nothing
	end if
	set obj_fso=nothing
end if
edit_filename=right(path,len(path)-instrrev(path,"/"))
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=SystemName%> - 文件编辑</title>
<link rel="stylesheet" type="text/css" href="images/edit.css" />
<script type="text/javascript" src="images/common.js"></script>
<script type="text/javascript" src="images/edit.js"></script>
</head>
<body>
<div id="bodyposition" onkeypress="hiddencopy()">
<div id="notepadmain">
<div id="notepadr1c1" onclick="closeabout(1)"></div>
<div id="notepadtitle" onclick="closeabout(1)"><%=edit_filename%> - 记事本<span id="content_change"></span></div>
<div id="notepadbutton"><a href="javascript:window.close()" onfocus='if(this.blur) this.blur()' title="关闭"><img src="images/notepad/blank.gif" width="15" height="13"></a></div>
<div id="notepadtool">
<table border="0" cellspacing="0" cellpadding="0">
<form name="filepos" method="get" action="<%=selfname%>">
  <tr>
    <td><a href="javascript:savefile()" title="保存"><img src="images/notepad/save.gif" align="absmiddle" /></a>
	<a href='javascript:recover("<%=edit_filename%>")' title="还原初始内容"><img src="images/notepad/revser.gif" align="absmiddle" /></a>
	</td>
    <td>&nbsp;&nbsp;位置：</td>
    <td><input type="text" id="filepath" name="path" value="<%=path%>" onkeypress="if(event.keyCode==13) document.filepos.submit()" title="文件路径，可以更改变成“另存为…”"></td>
    <td><input type="button" id="readfile" value='读文件' onfocus="if(this.blur) this.blur()" onclick="document.filepos.submit()"></td>
    <td>&nbsp;&nbsp;用ADO组件</td>
    <td><input type="checkbox" name="adodb" value="yes"<%if adodb="yes" then response.Write(" checked")%> onClick="useado()" onfocus="if(this.blur) this.blur()" title="读取文件乱码时请尝试使用ADODB.Stream"></td>
    <td id="toolbarright" style="display:none">
		<table border="0" cellspacing="0" cellpadding="0">
		  <tr>
			<td>&nbsp;&nbsp;编码</td>
			<td><input type="text" id="code" name="code" style="margin:0px;width:60px;border:solid 1px #000099;" value="<%=code%>" title="填写ADO组件读取文件的编码">&nbsp;</td>
			<td>
			<select id="encoding" size="1" onChange="$('code').value=$('encoding').value">
			<option value="GB2312">Encoding</option>
			<option value="GB2312">GB2312</option>
			<option value="Unicode">Unicode</option>
			<option value="UTF-8">UTF-8</option>
			</select></td>
		  </tr>
		</table>
	</td>
  </tr>
</form>
</table>
<script language="javascript">useado();</script>
</div>
<div id="notepadr2c1"><img src="images/notepad/notepad_r2_c1.gif" width="743" height="22"></div>
<div id="notepadr3c1"></div>
<div id="notepadcontent">
<%
filepath=Server.MapPath(path)
set obj_fso=Server.CreateObject("Scripting.FileSystemObject")
if obj_fso.FileExists(filepath) then
	if adodb<>"yes" then
		set obj_file=obj_fso.OpenTextFile(filepath,1,false)
		if not obj_file.atendofstream then
			s_content=htmlcode(obj_file.readall())
		end if
		obj_file.close
		set obj_file=nothing
	else
		call AdodbRead(filepath,code)
	end if
%>
<form name="savefile" action="<%=selfname&"?path="&server.urlencode(path)%>" method="post">
<input type="hidden" name="act" value="editfile">
<input type="hidden" name="path" value="<%=path%>">
<input type="hidden" name="adodb" id="useado" value="<%=adodb%>">
<input type="hidden" name="code" id="charset" value="<%if lcase(trim(code))<>"gb2312" then response.Write(code)%>">
<textarea name="fcontent" cols="102" rows="26" onkeydown="$('content_change').innerHTML='&nbsp;*'"><%=s_content%></textarea>
</form>
<script language="javascript">store("<%=edit_filename%>");</script>
<%
else
	response.write "<br><br><br><br><br><p align=center><font color=#ff3333>你要编辑的文件不存在！</font></p>"
end if
set obj_fso=nothing
%>
</div>
<div id="notepadr3c3"></div>
<div id="notepadr4c1"><img src="images/notepad/notepad_r4_c1.gif" width="743" height="22"></div>
</div>
<div id="aboutmain" style="display:none">
<div id="okbutton"><a href="./" onClick="return closeabout();" onfocus='if(this.blur) this.blur()'><img src="images/notepad/blank.gif" width="75" height="23"></a></div>
<div id="closebutton"><a href="./" onClick="return closeabout();" onfocus='if(this.blur) this.blur()'><img src="images/notepad/blank.gif" width="16" height="14"></a></div>
<div id="aboutr1c1">
<span style="position:relative;top:6px; left:6px;">关于<%=SystemName%> - 文件编辑器</span>
</div>
<div id="aboutr3c1">
<span style="position:relative;top:8px;"><%
endtime=timer()
if endtime<starttime then
	endtime=endtime+24*3600
end if
%><span style='position:relative;top:4px; text-align:center;line-height:120%;'>
<%=copyright%><br>Processed in <%=(endtime-starttime)*1000%> MSEL
</span>
</span>
</div>
<div id="aboutr4c1"></div>
<div id="aboutr5c1"></div>
<div id="aboutr6c1"></div>
</div>
</div>
</body>
</html>
<%
Sub AdodbRead(filepath1,code1)
	on error resume next
	dim objStream
	Set objStream = Server.CreateObject("adodb.stream")
	if Err then
		Err.Clear
		call showmsg("服务器不支持ADO组件","")
		response.End
	end if
	objStream.Type = 2
	objStream.Open
	objStream.Charset = code1
	if Err then
		Err.Clear
		call showmsg("无效编码","")
		response.End
	end if
	objStream.LoadFromFile(filepath1)
	if Err then
		Err.Clear
		call showmsg("文件路径错误","")
		response.End
	end if
	s_content = htmlcode(objStream.ReadText)
	objStream.Close
	Set objStream = Nothing
End Sub


Function SaveToFile(fileContent,thePath,code1)
	Dim objStream
    On Error Resume Next
    Set objStream = Server.CreateObject("ADODB.Stream")
	With objStream
		.Type=2
		.Mode=3
		.Open
		.Charset=code1
		.WriteText fileContent
		if Err then
			Set objStream = Nothing
			Err.Clear
			call showmsg("ADODB.Stream初始化错误","")
			response.End
		end if
		.saveToFile thePath, 2
		.Close
	End With
	Set objStream = Nothing
	if Err then
		Err.Clear
		SaveToFile = false
	else
		SaveToFile = true
	end if
End Function
%>