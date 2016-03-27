<!--#include file="lianjie.asp"---->
<%id=cint(trim(request("id")))
set rsn=server.CreateObject("adodb.recordset")
rsn.open"delete * from down where id="&id,cn,1,3
response.Redirect "down_list.asp"
rsn.close
set rsn=nothing
cn.close
set cn=nothing
%>


