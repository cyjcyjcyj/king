<!--#include file="checklogin.asp"-->
<%
dim filepathname,s,fso,f,intFilelength
filepathname = Server.MapPath(path)

Response.Buffer = True
Response.Clear

on error resume next
if Err then Err.Clear
Set s = Server.CreateObject("ADODB.Stream")
if Err then
	Err.Clear
	call dirdown(filepathname,"ADODB.Stream")
end if
s.Open
s.Type = 1

if Err then Err.Clear
Set fso = Server.CreateObject("Scripting.FileSystemObject")
if Err then
	Err.Clear
	call dirdown(filepathname,"FSO")
end if
if not fso.FileExists(filepathname) then
	Response.Write("<h1>Error:</h1>" & filepathname & " does not exist<p>")
	Response.End
end if

Set f = fso.GetFile(filepathname)
intFilelength = f.size

s.LoadFromFile(filepathname)
if Err then
	Err.Clear
	Response.Write("<h1>Error: </h1>" & Err.Description & "<p>")
	Response.End
end if

Response.AddHeader "Content-Disposition", "attachment; filename=" & f.name
Response.AddHeader "Content-Length", intFilelength
Response.CharSet = "UTF-8"
Response.ContentType = "application/octet-stream"

Response.BinaryWrite s.Read
Response.Flush

s.Close
Set s = Nothing

sub dirdown(FileRealPath,msg)
	Response.Expires = -1
	Response.AddHeader "Pragma", "no-cache"
	Response.AddHeader "cache-ctrol", "no-cache"
	FileRealPath=replace(FileRealPath,Server.MapPath("/")&"\","")
	FileRealPath=replace(FileRealPath,"\","/")
	FileRealPath=Server.URLEncode(FileRealPath)
	FileRealPath=Replace(FileRealPath,"+","%20")
%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=SystemName%> - 下载文件</title>
</head>
<body>
<script language="javascript">
<!--
if(window.confirm("出错：服务器不支持<%=msg%>！无法使用<%=msg%>下载文件。\n\n要直接打开文件地址下载吗？"))
	window.location.href="/<%=Replace(FileRealPath,"%2F","/")%>";
else
	window.close();
//-->
</script>
</body>
</html>
<%
	response.end
end sub
%>