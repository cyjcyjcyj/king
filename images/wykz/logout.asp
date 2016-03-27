<!--#include file="config.asp"-->
<!--#include file="md5.asp"-->
<%
'if request.QueryString("iplock")="empty" then
'	Application.Lock
'	Application(OnlineValue)=Empty
'	Application.unLock
'end if
if LoginOneIP=1 then
	Application.Lock
	if Application(OnlineValue)<>Empty and Application(OnlineValue)<>"" then
		Application(OnlineValue)=replace(Application(OnlineValue),"|" & md5s(session(SessionPrefix&"logonusername")) & ";" & session(SessionPrefix&"loginstatus"),"")
	end if
	Application.unLock
end if
session.timeout=1
session.abandon
response.redirect "index.asp"
%>