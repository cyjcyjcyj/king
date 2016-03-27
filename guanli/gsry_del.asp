<!--#include file="lianjie.asp"---->
<%id=cint(trim(request("id")))
 leixing=trim(request("leixing"))

set rsn=server.CreateObject("adodb.recordset")
rsn.open"delete * from gsry where id="&id,cn,1,3
response.Redirect "gsry_list.asp?leixing="&leixing
rsn.close
set rsn=nothing
cn.close
set cn=nothing
%>



