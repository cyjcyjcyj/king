<!--#include file="checklogin.asp"-->
<%
if not popedom then
	call showmsg("你没有权限进入此页面！","index.asp")
	response.end
end if
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=SystemName%> - ASP工具</title>
<style type="text/css">
<!--
body{font-family: "宋体","Times New Roman", Times, serif; font-size:12px; text-align:center;}
td,select{font-size:12px;}
.table{border-left:1px #999999 solid;}
.trtrb{border-top:1px #999999 solid;border-right:1px #999999 solid; border-bottom:1px #999999 solid;}
.trtr{border-top:1px #999999 solid;border-right:1px #999999 solid;}
.tx{font-family: "宋体";font-size:12px;border:1px solid;border-color:black black #000000;color: #0000FF;}
.button{border:1px #666666 solid; background-color:#FFFFFF; height:18px;}
-->
</style>
</head>
<body leftmargin="0">
<%
dim act,thePath
act=lcase(trim(request("action")))
if act="combine" then
   '用ASP将文件分割器分割的文件合并
   dim fname,f,newname
   newname=request("newname")
   set f=request("f")
   for i=1 to f.count
	if f(i)<>"" then
		fname=fname&"|"&f(i)
	end if
   next
   if newname="" then
	call back("新文件名不能为空！")
   end if
   if fname="" then
	call back("需合并文件名不能全为空！")
   end if
   call combine(fname,newname)
elseif act="addtomdb" or act="releasefrommdb" then
	thePath = Request("thePath")
	Script_TimeOut = trim(request("timeout"))
	if IsInteger(Script_TimeOut) then
		Script_TimeOut = round(Script_TimeOut*60,0)
	else
		Script_TimeOut = 3600
	end if
	Server.ScriptTimeOut = Script_TimeOut
	if act="addtomdb" then
		addToMdb(thePath)
		response.write "<script language=javascript>alert('操作完成！');window.close();</script>"
	elseif act="releasefrommdb" then
		unPack(thePath)
		response.write "<script language=javascript>alert('操作完成！');window.close();</script>"
	end if
end if
%>
<form name="form1" method="post" action="<%=selfname%>">
  <table width="542" border="0" cellspacing="0" cellpadding="0" align="center" class="table">
    <tr bgcolor="#CCCCCC"> 
      <td class="trtr" height="22" align="center" valign="middle" bgcolor="#CCCCCC" width="540"><B>ASP文件合并器 v1.2 by 秋忆</B></td>
    </tr>
	<tr valign="middle"> 
      <td class="trtr" align="left" height="120" width="540" style="line-height:140%;"><b>使用说明：</b>先用任意一款文件分割器将要上传的文件分割为若干小文件(IIS6.0默认上传不能大于200K)，然后上传到服务器，再将上传后的小文件名填入列表(相对路径)，填入新文件名，马上就可以生成和原来一样的文件了。<br>
	  <b>使用范围：</b>请确认服务器的Adodb.Stream可用。有完善的上传功能，相信这个没有多少人用的了，附带着吧，呵呵~~ ^_^<br>
	  <b>严重声明：</b>请勿使用该文件从事非法活动，本人不承担任何责任。
      </td>
    </tr>
    <tr>
	<td class="trtr" height="70">
	<div style="position:relative; top:0px; left:15px;">
        <li>合并后的新文件名：
          <input type="text" name="newname" class="tx" style="width:200px;">
        </li>
		<br><%if msie() then response.Write("<br>")%>
        <li>需要合并的文件数： 
          <input type="text" name="upcount" class="tx" value="3" style="width:50px" onkeydown="if(window.event.keyCode==13) {setid();return false}">
          <input type="button" class="button" onclick="setid();" value="设 定">
        </li>
	</div>
	</td>
    </tr>
    <tr valign="middle"> 
      <td class="trtr" align="left" id="upid" height="122" width="540"> 需合并文件: 
        <input type="text" name="f" style="width:350px" class="tx">
      </td>
    </tr>
    <tr valign="middle" bgcolor="#eeeeee" align="center"> 
      <td class="trtrb" bgcolor="#eeeeee" height="24" width="540">
        <input type="submit" value="合 并" class="button">&nbsp;&nbsp;&nbsp;&nbsp;
        <input type="reset" value="重 设" class="button">
		<input type="hidden" name="action" value="combine">
      </td>
    </tr>
  </table>
<script language="javascript">
function setid()
{
	  str='<br>';
	  if(!window.form1.upcount.value)
	   window.form1.upcount.value=1;
 	  for(i=1;i<=window.form1.upcount.value;i++)
	     str+='需合并文件'+i+'：<input type="text" name="f" style="width:350px" class="tx"><br><br>';
	  window.upid.innerHTML=str+'<br>';
}
	setid();
</script>
</form>
<br><br>

<table width="542" border="0" cellspacing="0" cellpadding="0" align="center" class="table">
    <tr bgcolor="#CCCCCC"> 
      <td class="trtr" height="22" align="center" valign="middle" bgcolor="#CCCCCC"><B>ASP文件打包/解包器 v1.0 by 秋忆</B></td>
    </tr>
	<tr><td>
<table width="542" border="0" cellspacing="0" cellpadding="0" align="center">
<form method=post target=_blank action="<%=selfname%>">
  <tr height="30">
    <td class="trtr">&nbsp;文件夹打包：</td>
    <td class="trtr">&nbsp;
	<input type="text" name="thePath" value="<%=Server.MapPath(".")%>" class="tx" style="width:300px">
	<input type="hidden" value="addToMdb" name="action">
	<select name="theMethod">
	<option value="fso">FSO</option>
	<option value="app">无FSO</option>
	</select>
	</td>
  </tr>
  <tr>
    <td class="trtr" colspan="2" height="25" align="center">
	脚本超时：<input type="text" name="timeout" value="60" class="tx" style="width:40px" />分钟　　
	<input type="submit" value="开始打包" class="button">
	</td>
  </tr>
  <tr>
    <td class="trtr" colspan="2" height="30">&nbsp;注：打包生成Qiuyi.mdb文件，位于当前页面目录<%=Server.MapPath(".")%>下。</td>
  </tr>
  <tr>
    <td class="trtr" colspan="2" height="40">&nbsp;</td>
  </tr>
  </form>
</table>
</td></tr>
<tr><td>
<table width="542" border="0" cellspacing="0" cellpadding="0" align="center">
<form method=post target=_blank action="<%=selfname%>">
  <tr>
    <td class="trtr" nowrap="nowrap" height="30">&nbsp;文件夹解包（需FSO支持）：</td>
    <td class="trtr" nowrap="nowrap">&nbsp;
	<input type="text" name="thePath" value="<%=Server.MapPath(".")%>\Qiuyi.mdb" class="tx" style="width:300px">
	<input type="hidden" value="releaseFromMdb" name="action">
	</td>
  </tr>
  <tr>
    <td class="trtr" colspan="2" height="25" align="center">
	脚本超时：<input type="text" name="timeout" value="60" class="tx" style="width:40px" />分钟　　
	<input type="submit" value="开始解包" class="button">
	</td>
  </tr>
  <tr>
    <td class="trtrb" colspan="2" height="30">&nbsp;注：解开的所有文件都位于当前页面目录<%=Server.MapPath(".")%>下。也可以亲自使用本系统附带的undo.vbs文件解开压缩包。</td>
  </tr>
</form>
</table>
</td></tr>
</table>
<table width="542" border="0" cellspacing="0" cellpadding="0" align="center">
<tr><td align="center">
<span style='position:relative;top:4px; text-align:center;line-height:120%;'>
<%
endtime=timer()
if endtime<starttime then
	endtime=endtime+24*3600
end if
response.Write(copyright)
%><br>Processed in <%=(endtime-starttime)*1000%> MSEL
</span>
</td></tr>
</table>
</body>
</html>
<%

sub back(str)
	response.write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">" & vbcrlf
	response.write "<script language=javascript>alert('"& str &"');history.back();</script>"
	response.end
end sub

sub combine(Filename,newname)
	on error resume next
	dim n,i,fso,dr
	newname=server.MapPath(newname)
	Filename=split(Filename,"|")
	i=ubound(Filename)
	redim fstr(i)
	
	if Err then Err.Clear
	set fso = Server.CreateObject("Scripting.FileSystemObject")
	if not Err then
		for n=1 to i
		   fname(n)=server.MapPath(Filename(n))
		   if not fso.FileExists(fname(n)) then
			set fso=nothing
			call back("文件“"&replace(Filename(n),"\","\\")&"”找不到！")
		   end if
		next
		set fso=nothing
	else
		Err.Clear
	end if
	
	if Err then Err.Clear
	set dr=Server.CreateObject("Adodb.Stream")
	if Err then
		Err.Clear
		call back("服务器不支持Adodb.Stream，无法使用合并功能！")
	end if
	for n=1 to i
	   dr.Mode=3
	   dr.Type=1
	   dr.Open
	   dr.LoadFromFile(fname(n))
	   fstr(n)=dr.read
	next
	
	dr.Mode=3
	dr.Type=1
	dr.Open
	for n=1 to i
	   dr.write=fstr(n)
	next
	dr.SaveToFile newname,2
	dr.Close
	set dr=nothing 
	response.write "新文件<b>"&newname&"</b>成功生成！"
	if Err then
		Err.Clear
		Response.Write("<h1>Error: </h1>" & Err.Description & "<p>")
	end if
end sub

Sub addToMdb(thePath)
	On Error Resume Next
	Dim rs, conn, stream, connStr, adoCatalog
	set rs = Server.CreateObject("Scripting.FileSystemObject")
	if not rs.FolderExists(thePath) then
		set rs = nothing
		response.Write("目录"&thePath&"不存在！")
		response.end
	end if
	set rs = nothing
	
	Set rs = Server.CreateObject("ADODB.RecordSet")
	Set stream = Server.CreateObject("ADODB.Stream")
	Set conn = Server.CreateObject("ADODB.Connection")
	Set adoCatalog = Server.CreateObject("ADOX.Catalog")
	connStr = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & Server.MapPath("Qiuyi.mdb")
	
	adoCatalog.Create connStr
	conn.Open connStr
	conn.Execute("Create Table FileData(Id int IDENTITY(0,1) PRIMARY KEY CLUSTERED, thePath VarChar, fileContent Image)")
	
	stream.Open
	stream.Type = 1
	rs.Open "FileData", conn, 3, 3
	
	If lcase(trim(Request("theMethod"))) = "fso" Then
		fsoTreeForMdb thePath, rs, stream
	 Else
		saTreeForMdb thePath, rs, stream
	End If
	
	rs.Close
	Conn.Close
	stream.Close
	Set rs = Nothing
	Set conn = Nothing
	Set stream = Nothing
	Set adoCatalog = Nothing
	if Err then
		Err.Clear
		Response.Write("<h1>Error: </h1>" & Err.Description & "<p>")
	end if
End Sub

Function fsoTreeForMdb(thePath, rs, stream)
	Dim item, theFolder, folders, files, sysFileList,fsoX
	sysFileList = "$Qiuyi.mdb$Qiuyi.ldb$"
	set fsoX = Server.CreateObject("Scripting.FileSystemObject")
	If fsoX.FolderExists(thePath) = False Then
		call back(thePath & " 目录不存在或者不允许访问!")
	End If
	Set theFolder = fsoX.GetFolder(thePath)
	Set files = theFolder.Files
	Set folders = theFolder.SubFolders
	
	For Each item In folders
		fsoTreeForMdb item.Path, rs, stream
	Next
	
	For Each item In files
		If InStr(sysFileList, "$" & item.Name & "$") <= 0 Then
			rs.AddNew
			rs("thePath") = Mid(item.Path, 4)
			stream.LoadFromFile(item.Path)
			rs("fileContent") = stream.Read()
			rs.Update
		End If
	Next
	
	set fsoX = Nothing
	Set files = Nothing
	Set folders = Nothing
	Set theFolder = Nothing
End Function

Sub saTreeForMdb(thePath, rs, stream)
		on error resume next
		Dim item, theFolder, sysFileList,saX
		sysFileList = "$Qiuyi.mdb$Qiuyi.ldb$"
		Set saX = Server.CreateObject("Shell.Application")
		Set theFolder = saX.NameSpace(thePath)
		
		For Each item In theFolder.Items
			If item.IsFolder = True Then
				saTreeForMdb item.Path, rs, stream
			 Else
				If InStr(sysFileList, "$" & item.Name & "$") <= 0 Then
					rs.AddNew
					rs("thePath") = Mid(item.Path, 4)
					stream.LoadFromFile(item.Path)
					rs("fileContent") = stream.Read()
					rs.Update
				End If
			End If
		Next

		Set saX = Nothing
		Set theFolder = Nothing
		if Err then
			Err.Clear
			Response.Write("<h1>Error: </h1>" & Err.Description & "<p>")
		end if
End Sub

Sub unPack(thePath)
		On Error Resume Next
		'Server.ScriptTimeOut = 5000
		Dim rs, ws, str, conn, stream, connStr, theFolder,fsoX
		set rs = Server.CreateObject("Scripting.FileSystemObject")
		if not rs.FileExists(thePath) then
			set rs = nothing
			response.Write("文件"&thePath&"不存在！")
			response.end
		end if
		set rs = nothing

		str = Server.MapPath(".") & "\"
		Set rs = CreateObject("ADODB.RecordSet")
		Set stream = CreateObject("ADODB.Stream")
		Set conn = CreateObject("ADODB.Connection")
		connStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & thePath & ";"

		conn.Open connStr
		rs.Open "FileData", conn, 1, 1
		stream.Open
		stream.Type = 1

		set fsoX = Server.CreateObject("Scripting.FileSystemObject")
		Do Until rs.Eof
			theFolder = Left(rs("thePath"), InStrRev(rs("thePath"), "\"))
			If fsoX.FolderExists(str & theFolder) = False Then
				createFolder(str & theFolder)
			End If
			stream.SetEos()
			stream.Write rs("fileContent")
			stream.SaveToFile str & rs("thePath"), 2
			rs.MoveNext
		Loop

		rs.Close
		conn.Close
		stream.Close
		set fsoX = Nothing
		Set ws = Nothing
		Set rs = Nothing
		Set stream = Nothing
		Set conn = Nothing
		if Err then
			Err.Clear
			Response.Write("<h1>Error: </h1>" & Err.Description & "<p>")
		end if
End Sub

Sub createFolder(thePath)
		on error resume next
		Dim i,fsoX
		i = Instr(thePath, "\")
		set fsoX = Server.CreateObject("Scripting.FileSystemObject")
		Do While i > 0
			If fsoX.FolderExists(Left(thePath, i)) = False Then
				fsoX.CreateFolder(Left(thePath, i - 1))
			End If
			If InStr(Mid(thePath, i + 1), "\") Then
				i = i + Instr(Mid(thePath, i + 1), "\")
			 Else
				i = 0
			End If
		Loop
		set fsoX = Nothing
		if Err then
			Err.Clear
			Response.Write("<h1>Error: </h1>" & Err.Description & "<p>")
		end if
End Sub
%>