<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
'注意，这个配置文件的内容已经不可以更改了，只能在管理里设置参数
Option Explicit
Response.Buffer=True
Server.ScriptTimeOut=90
Session.CodePage=936

dim starttime,SystemName,SystemVersion,selfname,userip,PerPageSize,CanEditFileType,CanNotEditFileType
dim showlasttime,copyright,OnlineValue,UploadObject,LoginOneIP,FileMaxSize,Script_Timeout
dim NotAllowFileType,AllowFileType,ShowValidate,SessionPrefix,endtime
starttime=timer()

Script_Timeout=3600
UploadObject=0
FileMaxSize=52428800
NotAllowFileType="asp|asa|aspx|cgi|php|php4|cer|cdx|htw|ida|idq|shtm|shtml|stm|printer|cfm|htm"
AllowFileType="gif,png,jpg,jpeg,bmp,asp,asa,aspx,cgi,php,htm,html,doc,xls,ppt,rar,zip,cab,swf,chm,pdf,mp3,wav,wma,wmv,mid,midi,avi,rm,ra,ram,rmvb,mov,torrent"
CanEditFileType="txt|bat|c|htm|html|css|js|vbs|inc|stm|shtm|shtml|asp|asa|aspx|asax|ascx|asmx|aspa|php|php3|php4|jsp|cgi|ini|inf|htt|reg|cpp|cmd|htc"
CanNotEditFileType="zip|rar|tar|exe|gif|jpg|jpe|jpeg|png|bmp|psd|mdb|ppt|xls|mid|mp3|avi|chm|swf"
PerPageSize=40
showlasttime=1
LoginOneIP=1
ShowValidate=0
SessionPrefix="skymean.com"

copyright="Copyright <font face='Times New Roman'>&copy;</font> 2005-2006 <a href='http://www.skymean.com' target=_blank><font color='#000080'><b>www.skymean.com</b></font></a> All Rights Reserved<br>Designed by 秋忆"
SystemName="秋忆工作室在线文件管理器"
SystemVersion="v4.4"
selfname=Request.ServerVariables("Script_Name")
selfname=right(selfname,len(selfname)-instrrev(selfname,"/"))
userip=trim(Request.ServerVariables("HTTP_X_FORWARDED_FOR"))
if userip="" then
	userip=trim(Request.ServerVariables("Remote_Addr"))
end if
OnlineValue=replace(cstr(date()),"-","") & "OnlineUsers"
%>
<!--#include file="dll.asp"-->
