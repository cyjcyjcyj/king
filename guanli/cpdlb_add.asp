<style type="text/css">
<!--
body,td,th {
	font-size: 12px;
}
-->
</style><!--#include file="lianjie.asp"-->

<%
dlb=request("dlb")
endlb=request("endlb")

set rs=server.CreateObject("adodb.recordset")
rs.open"select * from dlb where dlb='"&dlb&"'",cn,1,3
if not rs.eof then
response.Write"<script>alert('�Բ���������ظ�');history.back();</script>"
else 
rs.addnew
rs("dlb")=dlb
rs("endlb")=endlb
rs.update
end if
rs.close
set rs=nothing
cn.close
set cn=nothing
response.write "<br><br><br><center>��ӳɹ�,<a href=cplb_list.asp>����</a><br><br>лл��ʹ�ñ�ϵͳ</font></center>"
response.end

%>
