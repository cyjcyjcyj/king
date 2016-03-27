<!--#include file="lianjie.asp"---->
<%id=cint(trim(request("id")))
set rsn=server.CreateObject("adodb.recordset")
rsn.open"delete * from xwlb where id="&id,cn,1,3
response.Redirect "xwlblist.asp"
%>


