<!--#include file="lianjie.asp"---->
<%id=cint(trim(request("id")))
xwlb=request("xwlb")
set rsn=server.CreateObject("adodb.recordset")
rsn.open"delete * from xw where id="&id,cn,1,3
response.Redirect "xwlb.asp?xwlb="&xwlb
%>


