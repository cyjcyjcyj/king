<!--#include file="lianjie.asp"---->
<%
id=cint(trim(request("id")))
set rsn=server.CreateObject("adodb.recordset")
rsn.open"delete * from dd where id="&id,cn,1,3
response.Redirect "dd.asp"

%>
