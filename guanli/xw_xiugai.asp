<!--#include file="lianjie.asp"-->
<style type="text/css">
<!--
body,td,th {
	font-size: 12px;
}
-->
</style>
<%id=cint(trim(request("id")))
xwbt=request("xwbt")

xwlb_id=request("xwlb_id")
xwnr=replace(trim(request("xwnr")),"'","")

tjsj=date()
if xwnr="" or xwbt="" then
response.Write"<script>alert('对不起，各项不能为空');history.back();</script>"
response.End()
end if

set rs=server.CreateObject("adodb.recordset")
rs.open"select * from xw where id="&id,cn,1,3
rs("xwnr")=xwnr

rs("xwbt")=xwbt

rs("xwlb_id")=xwlb_id
rs("tjsj")=tjsj
rs.update
rs.close
set rs=nothing
cn.close
set cn=nothing

url="xwxiugai.asp?id="&id
  response.write "<br><br><br><center>添加成功,<a href="&url&">返回</a><br><br>谢谢您使用本系统</font></center>"
response.end





%>

<meta http-equiv="Content-Type" content="text/html; charset=gb2312">

