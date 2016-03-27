<!--#include file="lianjie.asp"---->
<%id=cint(trim(request("id")))

set rsn=server.CreateObject("adodb.recordset")
rsn.open"delete * from yqlj where id="&id,cn,1,3
response.Redirect "yqlj_list.asp"
rsn.close
set rsn=nothing
cn.close
set cn=nothing
%>


