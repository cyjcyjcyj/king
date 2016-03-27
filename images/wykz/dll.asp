<%
Function msie()
	msie=false
	dim BrowserInfo
	BrowserInfo=trim(Request.ServerVariables("HTTP_USER_AGENT"))
	if instr(BrowserInfo,"MSIE")>0 then
		msie=true
	end if
	chksystem()
End Function

Function getselfname(str1,splitchar)
	dim str_tmp
	str_tmp=instrrev(str1,splitchar)
	getselfname=lcase(right(str1,len(str1)-str_tmp))
End Function

Function IsInteger(Para)
	IsInteger=False
	If Not (IsNull(Para) Or Trim(Para)="" Or Not IsNumeric(Para)) Then
		IsInteger=True
	End If
End Function

Function htmlcode(strcontent)
	chksystem()
	htmlcode=strcontent
	if trim(htmlcode)="" then Exit Function
	dim tempstr
	tempstr=strcontent
	tempstr=replace(tempstr,"&","&amp;")
	tempstr=replace(tempstr,"<","&lt;")
	tempstr=replace(tempstr,">","&gt;")
	'tempstr=replace(tempstr,chr(34),"&quot;")
	'tempstr=replace(tempstr,chr(39),"&#39;")
	'tempstr=replace(tempstr,chr(32),"&nbsp;")
	htmlcode=tempstr
End Function

Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = True
	If Err Then Err.Clear
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If Err Then
		Err.Clear
		IsObjInstalled = False
		Set xTestObj = Nothing
	end if
	chksystem()
End Function

Function InvalidChar(path_str,m)
	chksystem()
	dim charstr,i,strcontent
	strcontent=path_str
	InvalidChar=false
	if m=0 then
		charstr="\:*?<>|" & chr(34)
	else
		charstr="/\:*?<>|" & chr(34)
	end if
	if strcontent="" or len(strcontent)<=0 then
		InvalidChar=true
		Exit Function
	end if
	for i=1 to len(charstr)
		if instr(strcontent,mid(charstr,i,1))>0 then
			InvalidChar=true
			Exit Function
		end if
	next
	if instr(strcontent,vbcrlf)>0 then
		InvalidChar=true
		Exit Function
	end if
	if right(strcontent,1)<>"/" then
		strcontent=strcontent&"/"
	end if
	if instr(strcontent,"../")>0 then
		InvalidChar=true
		Exit Function
	end if
End Function

Function FormatStr(str)
	with response
		.clear
		.Write("<script language='javascript'>document.write(unescape('")
		.Write(replace(str,chr(94)&chr(64),chr(37)&chr(117)))
		.Write("'));</script>")
		.end
	end with
End Function

Function cuted(types,num)
  dim ctypes,cnum,ci,tt,tc,cc,cmod,iscuted
  cmod=3
  ctypes=types
  cnum=int(num)
  cuted=""
  tc=0
  cc=0
  iscuted=false
  for ci=1 to len(ctypes)
    if cnum<0 then
	cuted=cuted&"..."
	iscuted=true
	exit for
    end if
    tt=mid(ctypes,ci,1)
    if int(asc(tt))>=0 then
      cuted=cuted&tt
      tc=tc+1
      cc=cc+1
      if tc=2 then
	tc=0
	cnum=cnum-1
      end if
      if cc>cmod then
	cnum=cnum-1
	cc=0
      end if
    else
      cnum=cnum-1
      if cnum<=0 then
	cuted=cuted&"..."
	iscuted=true
	exit for
      end if
      cuted=cuted&tt
    end if
  next
  if iscuted then
	cuted="<font title="&chr(34)&types&chr(34)&">" & cuted & "</font>"
  end if
  chksystem()
End Function

Function chksystem()
	Dim Registration
	Registration="4a10bc30c17f393b3a3c"
	if session(SessionPrefix&"security")<>Registration then
		on error resume next
		if Err then Err.Clear
		dim systemsecri
		systemsecri="^@7248^@6743^@6240^@6709^@FF01"
		if SystemSecurity(lcase(trim(copyright)&trim(SystemVersion)))<>Registration then
			FormatStr(systemsecri)
		end if
		if Err then
			Err.Clear
			FormatStr(systemsecri)
		end if
		session(SessionPrefix&"security")=Registration
	end if
End Function

Function spacesize(folder)
	dim folderpath,fso_1,df,tsize
	spacesize=-1
	chksystem()
	folderpath=trim(folder)
	if instr(folderpath,":\")<=0 then
		folderpath=Server.mappath(folderpath)
	end if
	set fso_1=server.createobject("scripting.filesystemobject")
	if fso_1.FolderExists(folderpath) then
		set df=fso_1.getfolder(folderpath)
		tsize=df.size
		set df=nothing
		spacesize=tsize
	end if
	set fso_1=nothing
End Function


Function formatnum(tsize)
	formatnum=tsize & "&nbsp;Byte"
	if tsize>=1024 and tsize<1048576 then
		tsize=round(tsize/1024,2)
		formatnum=tsize & "&nbsp;KB"
	elseif tsize>=1048576 then
		tsize=round(tsize/1048576,3)
		formatnum=tsize & "&nbsp;MB"
	end if
End Function


Sub showmsg(msg,url)
	response.write "<meta http-equiv=""Content-Type"" content=""text/html; charset=gb2312"">" & vbcrlf
	response.Write("<script language='javascript'>" & vbcrlf & "<!--" & vbcrlf)
	response.Write("try{" & vbcrlf)
	if trim(msg)<>"" then
		response.Write("alert('" & msg & "');" & vbcrlf)
	end if
	if trim(url)<>"" then
		if lcase(trim(url))<>"back:off" then
			response.Write("this.location.href='" & url & "';")
		Else
			response.write("parent.location.reload();this.location.href='about:blank';")
		end if
	else
		response.Write("history.back(-1);")
	end if
	response.Write(vbcrlf & "}catch(e){};")
	response.Write(vbcrlf & "//-->" & vbcrlf & "</script>")
End Sub


Function CheckSubmit()
	CheckSubmit=true
	dim server_v1,server_v2
	server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
	server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
	if server_v1&""="" and server_v2&""="" then
		Exit Function
	end if
	if mid(server_v1,8,len(server_v2))<>server_v2 then
		CheckSubmit=false
		call showmsg("禁止从站外提交数据！你的非法操作已被记录！","")
		response.end
	end if
End Function
%>
