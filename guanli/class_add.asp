<style type="text/css">
<!--
body,td,th {
	font-size: 12px;
}
-->
</style><!--#include file="lianjie.asp"-->

<%
xxlb=request("yjfl")
enxxlb=request("enyjfl")
if xxlb="" or enxxlb="" then
response.Write"<script>alert('�Բ��𣬸����Ϊ��');history.back();</script>"
response.End()
end if

set rs=server.CreateObject("adodb.recordset")
rs.open"select * from class where xxlb='"&xxlb&"'",cn,1,3
if not rs.eof then
response.Write"<script>alert('�Բ���������ظ�');history.back();</script>"
else 
rs.addnew
rs("enxxlb")=enxxlb
rs("xxlb")=xxlb
rs.update
end if
rs.close
set rs=nothing
cn.close
set cn=nothing
response.write "<br><br><br><center>��ӳɹ�,<a href=class_list.asp>����</a><br><br>лл��ʹ�ñ�ϵͳ</font></center>"
response.end

%>
