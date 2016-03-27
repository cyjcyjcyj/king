<% 
lujing="shuju/woai#asp#shuju.mdb"
set cn=server.CreateObject("adodb.connection")
cn.open"provider=microsoft.jet.oledb.4.0;data source="&server.MapPath(""&lujing&"")

%>
<!--#include file="md5.asp"------>
<%yhm=replace(trim(request("yhm")),"'","")
mm=replace(trim(request("mm")),"'","")
if yhm="" or mm="" then
response.Write"<script>alert('对不起，各项不能为空');history.back();</script>"
response.End()
end if
if not isnumeric(request.form("yzm")) then
response.Write "<script LANGUAGE='javascript'>alert('登录失败！验证码必须是数字，请正确填写！');history.go(-1);</script>"
response.end
end if
on error resume next
yzm=Cint(request.form("yzm"))


if yzm<>Session("GetCode") then
response.Write "<script LANGUAGE='javascript'>alert('登录失败！验证码错误！');history.go(-1);</script>"
response.end
end if

set rs=server.CreateObject("adodb.recordset")
rs.open"select * from guanliyuan where yhm='"&yhm&"' and mm='"&md5(mm)&"'",cn,1,3
if rs.eof then
response.Write "<script LANGUAGE='javascript'>alert('用户名或密码错误！');history.go(-1);</script>"
response.end
else
response.cookies("CookieName")=rs("yhm")
response.cookies("CookieName").Expires=date+1 
response.Redirect "guanli.asp"
end if
%>



