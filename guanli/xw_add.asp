<!--#include file="lianjie.asp"-->

<%
xwbt=request("xwbt")
enxwbt=request("enxwbt")
xwlb_id=request("xwlb_id")
xwnr=replace(trim(request("xwnr")),"'","")
enxwnr=replace(trim(request("enxwnr")),"'","")
tjsj=date()

set rs=server.CreateObject("adodb.recordset")
rs.open"select * from xw ",cn,1,3

rs.addnew
rs("xwnr")=xwnr
rs("enxwnr")=enxwnr
rs("xwbt")=xwbt
rs("enxwbt")=enxwbt
rs("xwlb_id")=xwlb_id
rs("tjsj")=tjsj
rs.update
  url="xw_list.asp"
  response.write "<br><br><br><center>��ӳɹ�,<a href="&url&">����</a><br><br>лл��ʹ�ñ�ϵͳ</font></center>"
response.end




%>

<meta http-equiv="Content-Type" content="text/html; charset=gb2312">

<style type="text/css">
<!--
body,td,th {
	font-size: 12px;
}
-->
</style>