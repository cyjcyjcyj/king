<!--#include file="lianjie.asp"-->

<%id=cint(trim(request("id")))
xwlb=replace(trim(request("yjfl")),"'","")
enxwlb=replace(trim(request("enyjfl")),"'","")
set rs=server.CreateObject("adodb.recordset")
rs.open"select * from xwlb where id="&id,cn,1,3
rs("xwlb")=xwlb
rs("enxwlb")=enxwlb
rs.update
url="xwlbxiugai.asp?id="&id
     response.Redirect url
     response.end
rs.close
set rs=nothing
cn.close
set cn=nothing



%>


<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
