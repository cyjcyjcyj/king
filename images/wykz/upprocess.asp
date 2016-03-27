<!--#include file="checklogin.asp"-->
<!--#include file="incupload.asp"-->
<%
Response.Expires = -1
Response.ExpiresAbsolute = Now() - 1
Response.Expires = 0
Response.CacheControl = "no-cache"

Server.ScriptTimeOut = Script_Timeout
dim upload,filepath,frompath,obj_fso,i,formName,file,act,msg

if popedom then		'管理员不限制上传
	FileMaxSize=0
	NotAllowFileType=""
	AllowFileType=""
end if

CreateUploadObj(UploadObject)

if act="uploadfile" then
	if filepath<>"" then
		filepath=replace(filepath,"\","/")
		filepath=replace(filepath,"//","/")
	end if
	if right(filepath,1)<>"/" then
		filepath=filepath&"/"
	end if
	if InvalidChar(filepath,0) then
		msg = msg & "<font color='#ff0000'>抱歉，没有填写路径或路径中含有非法字符！</font><br><br>"
		msg = msg & "<input type='button' style='width:65px;height:20px;font-size:12px' value='返回' onclick='window.history.go(-1)'>　　　<input type='button' style='width:65px;height:20px;font-size:12px' value='关闭' onclick='window.close()'>"&vbcrlf
		call main("出错信息",msg)
		response.end
	end if
	Call CheckPath(PathCanModify,filepath)

	frompath=Server.mappath(filepath)
	set obj_fso=server.createobject("scripting.filesystemobject")
	if not obj_fso.folderexists(frompath) then
		set obj_fso=nothing
		set upload=nothing
		msg = msg & "<font color='#ff0000'>抱歉，目录“"&filepath&"”不存在，请先创建该目录！</font><br>"&vbcrlf
		msg = msg & "<br><input type='button' style='width:65px;height:20px;font-size:12px' value='返回' onclick='window.history.go(-1)'>　　　<input type='button' style='width:65px;height:20px;font-size:12px' value='关闭' onclick='window.close()'>"
		call main("出错信息",msg)
		response.end
	end if
	
	i=0
	select case UploadObject
		case 1		'Aspupload 组件上传
			aspupload()
		case 2		'SA-FileUp 组件上传
			saupload()
		case 3		'LyfUpload 组件上传
			lyfupload()
		case else	'无组件上传
			nogroupware()
	end select
	
	set upload=nothing
	set obj_fso=nothing
	msg = msg & "<br><br><b>共&nbsp;<font color='#ff0000'>"&i&"</font>&nbsp;个文件上传成功！</b><br>"
	msg = msg & "<br><input type='button' style='width:65px;height:20px;font-size:12px' value='返回' onclick='window.history.go(-1)'>　　　<input type='button' style='width:65px;height:20px;font-size:12px' value='关闭' onclick='window.close()'>"&vbcrlf
	call main("上传成功",msg)
else
	call main("提交参数不正确","<br>act=" & act)
	set upload=nothing
end if



function CanUpload(Fileurl)
	Fileurl = lcase("|"& Mid(Fileurl, InstrRev(Fileurl, ".") + 1)& "|")
	NotAllowFileType = "|"&NotAllowFileType&"|"
	if instr(NotAllowFileType,Fileurl)>0 then
		CanUpload = false
	else
		CanUpload = true
	end if
end function



function CreateUploadObj(Upload_Object)
	on error resume next
	If Err Then Err.Clear
	select case Upload_Object
		case 1		'Aspupload 组件上传
			set upload=Server.CreateObject("Persits.Upload")
		case 2		'SA-FileUp 组件上传
			set upload=Server.CreateObject("SoftArtisans.FileUp")
		case 3		'LyfUpload 组件上传
			set upload=Server.CreateObject("LyfUpload.UploadFile")
		case else	'无组件上传
			set upload=new DoteyUpload
	end select
	if Err then
		Err.Clear
		msg = msg & "抱歉，服务器不支持 "
		select case Upload_Object
			case 1
				msg = msg & "AspUpload组件"
			case 2
				msg = msg & "SA-FileUp组件"
			case 3
				msg = msg & "LyfUpload组件"
			case else
				msg = msg & "无组件"
		end select
		msg = msg & " 上传，请在config.asp文件里设置服务器支持的组件"
		call main("出错信息","<font color='#ff0000'>"&msg&"</font>")
		response.end
	end if
	
	select case Upload_Object
		case 1		'Aspupload 组件上传
			act=lcase(trim(request.QueryString("action")))
			filepath=lcase(trim(request.QueryString("path")))
		case 2		'SA-FileUp 组件上传
			act=lcase(trim(upload.form("act")))
			filepath=trim(upload.form("filepath"))
		case 3		'LyfUpload 组件上传
			act=lcase(trim(upload.request("act")))
			filepath=trim(upload.request("filepath"))
		case else	'无组件上传
			upload.ProgressID = Request.QueryString("ProgressID")
			upload.Upload()
			act=lcase(trim(upload.Form("act")))
			filepath=trim(upload.Form("filepath"))
	end select
end function


function aspupload()
	On Error Resume Next
	if Request.QueryString("PID") = "" then
		upload.ProgressID="010D60EB00C5AA4B"
	else
		upload.ProgressID=Request.QueryString("PID")
	end if
	if FileMaxSize>0 then
		upload.SetMaxSize FileMaxSize, True
	end if
	upload.save(frompath)
	If Err.Number = 8 Then
		Err.Clear
		msg = msg & "<font color=red>上传的部分文件大小超过限制（不超过"& formatnumber(FileMaxSize/1024,2,-1) &"K）</font><br>"
	end if
	For Each file in upload.Files
		if not CanUpload(file.ext) then
			if trim(file.filename)&""<>"" then
				File.Delete
				msg = msg & "<font color=red>文件"&file.filename &"("&formatnumber(file.size/1024,2,-1)&" K) 格式不允许上传！</font><br>"
			end if
		else
			i=i+1
			msg = msg & file.filename & " ("&formatnumber(file.size/1024,2,-1)&" K)&nbsp;上传至&nbsp;"
			msg = msg & Server.MapPath(filepath & file.filename) & "&nbsp;成功!<br><br>"
		end if
	next
	If Err Then
		Err.Clear
		msg = msg & "出现错误: " & Err.Number & "<br>" & Err.Description
	End If
end function


function saupload()
	dim temp_N,Filesize,thisfiletype,canup
	for each FormName in upload.Form
		canup=true
		if IsObject(upload.Form(FormName)) Then
			If Not upload.Form(FormName).IsEmpty Then
				if FileMaxSize>0 then
					upload.Form(FormName).Maxbytes = FileMaxSize
				end if
				Filesize = upload.Form(FormName).TotalBytes
				temp_N = mid(upload.Form(FormName).UserFileName,InStrRev(upload.Form(FormName).UserFileName,"\")+1)
				thisfiletype= Mid(temp_N, InStrRev(temp_N, ".")+1)
				if FileMaxSize>0 then
					If Filesize>FileMaxSize then
						canup=false
						msg = msg & "<font color=red>文件"&upload.Form(FormName).UserFilename&"（"&formatnumber(Filesize/1024,2,-1)&" K）超过大小限制（不超过"& formatnumber(FileMaxSize/1024,2,-1) &"K）</font><br>"
					end if
				end if
				if not CanUpload(thisfiletype) then
					canup=false
					if trim(temp_N)<>"" then
						msg = msg & "<font color=red>文件"&temp_N&"("&Filesize&") 格式不允许上传！</font><br>"
					end if
				end if
				if canup=true then
					upload.Form(FormName).SaveAs frompath & "\" & temp_N
					msg = msg & upload.Form(FormName).UserFilename&" ("&formatnumber(Filesize/1024,2,-1)&" K)&nbsp;上传至&nbsp;"
					msg = msg & upload.Form(FormName).ServerName&"&nbsp;成功!<br><br>"
					i=i+1
				end if
			end if
		end if
	next
end function


function lyfupload()
	if FileMaxSize>0 then
		upload.maxsize=FileMaxSize
	end if
	if AllowFileType<>"" then
		upload.extname=AllowFileType
	end if
	
	dim filename(4),filesize(4),j,canup
	filename(0)=upload.SaveFile("file1",frompath,true)
	filesize(0)=upload.FileSize
	filename(1)=upload.SaveFile("file2",frompath,true)
	filesize(1)=upload.FileSize
	filename(2)=upload.SaveFile("file3",frompath,true)
	filesize(2)=upload.FileSize
	filename(3)=upload.SaveFile("file4",frompath,true)
	filesize(3)=upload.FileSize
	filename(4)=upload.SaveFile("file5",frompath,true)
	filesize(4)=upload.FileSize
	for j=0 to 4
		canup=true
		if filename(j)<>"" then
			if FileMaxSize>0 then
				if filename(j)="0" then
					canup=false
					msg = msg & "<font color=red>第"&j+1&"个文件大小超过限制（不超过"& formatnumber(FileMaxSize/1024,2,-1) &"K）</font><br>"
				end if
			end if
			if AllowFileType<>"" then
				if filename(j)="1" then
					canup=false
					msg = msg & "<font color=red>第"&j+1&"个文件格式不允许上传</font><br>"
				end if
			end if
			if canup=true then
				i=i+1
				msg = msg & filename(j)&" ("&formatnumber(filesize(j)/1024,2,-1)&" K)&nbsp;上传至&nbsp;"
				msg = msg & Server.MapPath(filepath&filename(j))&"&nbsp;成功!<br>"
			end if
		end if
	next
	erase filename
	erase filesize
end function


function nogroupware()
	dim canup,fileobj,clientpath
	for each fileobj in upload.Files.Items
		canup=true
		clientpath=fileobj.FilePath
		if len(clientpath)>80 then
			clientpath="..."&right(clientpath,77)
		end if
		if FileMaxSize>0 then
			if fileobj.FileSize>FileMaxSize then
				canup=false
				msg = msg & "<font color=red>文件"&clientpath&"（"&formatnumber(fileobj.FileSize/1024,2,-1)&"K）超过大小限制（不超过"& formatnumber(FileMaxSize/1024,2,-1) &"K）</font><br>"
			end if
		end if
		if trim(fileobj.FileExt)&""<>"" then
			if not CanUpload(fileobj.FileExt) then
				canup=false
				msg = msg & "<font color=red>文件"&clientpath&"（"&formatnumber(fileobj.FileSize/1024,2,-1)&"K）格式不允许上传</font><br>"
			end if
		end if
		if canup=true and trim(fileobj.FileName)&""<>"" then
			fileobj.SaveAs frompath&"\"&fileobj.FileName
			msg = msg & clientpath&" ("&formatnumber(fileobj.FileSize/1024,2,-1)&" K)&nbsp;上传至&nbsp;"
			msg = msg & Server.MapPath(filepath&fileobj.FileName)&"&nbsp;成功!<br>"
			i=i+1
		end if
	next
end function


sub main(title,body)
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />
<title><%=SystemName&" - "&title%></title>
<link rel="stylesheet" type="text/css" href="images/upresult.css" />
</head>
<body<%if title="上传成功" then response.write " onload='opener.window.location.reload()'"%>>
<div id="bodyposition">
	<div id="content">
	  <div id="title"><%=SystemName&" - 文件上传结果报告"%></div>
	  <div id="main"><br><%=body%><br><br><br></div>
	  <div id="footer"><br><%
endtime=timer()
if endtime<starttime then
	endtime=endtime+24*3600
end if
%><span style='position:relative;top:4px; text-align:center;line-height:120%;'>
<%=copyright%><br>Processed in <%=(endtime-starttime)*1000%> MSEL
</span><br><br></div>
	</div>
</div>
</body>
</html>
<%end sub%>

