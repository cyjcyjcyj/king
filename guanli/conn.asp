<% 
adminuser=Request.Cookies("CookieName")
if adminuser="" then
response.Write "<script LANGUAGE='javascript'>alert('������˼����Ȩ��½��');history.go(-1);</script>"
response.end
else
 lujing="shuju/woai#asp#shuju.mdb"
set cn=server.CreateObject("adodb.connection")
cn.open"provider=microsoft.jet.oledb.4.0;data source="&server.MapPath(""&lujing&"")
end if
%>

