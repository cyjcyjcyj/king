<!--#include file="lianjie.asp"----->
<%xwlb_id=request("xwlb_id")
set rsgly=server.CreateObject("adodb.recordset")
rsgly.open"select * from xw where xwlb_id="&xwlb_id,cn,1,3
set rsall=server.CreateObject("adodb.recordset")
rsall.open"select * from xwlb",cn,1,3
dim pages,cur_recordcounts
pages=clng(request.QueryString("pages"))
rsgly.pagesize=15
if not rsgly.eof then
 if pages=0 then
 pages=1
  else
  rsgly.move(pages-1)*rsgly.pagesize
  end if
 if pages<rsgly.pagecount then
  cur_recordcounts=rsgly.pagesize
  else
   cur_recordcounts=rsgly.recordcount-(pages-1)*rsgly.pagesize
   end if
end if
%>
<title>����ҳ��</title>
<META http-equiv=Content-Type content=text/html;charset=gb2312>

<style type="text/css">
<!--
.style1 {color: #DBBF90;
	font-weight: bold;
}
body,td,th {
	font-size: 12px;
}
a:link {
	text-decoration: none;
}
a:visited {
	text-decoration: none;
}
a:hover {
	text-decoration: underline;
}
a:active {
	text-decoration: none;
}
-->
</style>
<script language="JavaScript" type="text/JavaScript">
function CheckAll(form)  {
  for (var i=0;i<form.elements.length;i++)    {
    var e = form.elements[i];
    if (e.name != 'chkall')       e.checked = form.chkall.checked; 
   }
  }
</script>
<BODY leftmargin="0" bottommargin="0" rightmargin="0" topmargin="5" marginheight="0" marginwidth="0" bgcolor="#DDEEFF">
<form action="xwlb.asp">
<input name="xwlb_id" type="hidden" value="<%=xwlb_id%>">
<table class="tableBorder" width="96%" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF">
  <tr>
    <td  align="center" background="images/admin_bg_1.gif" height="25"><span class="style1">�����б�</span></td>
  </tr>
  <tr>
    <td height="20"  align="center" bgcolor="#E8F1FF"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td align="center"><%for k=1 to rsall.recordcount%>
            <span class="style1"><a href=xwlb.asp?xwlb_id=<%=rsall("id")%>><%=rsall("xwlb")%></a></span>&nbsp;&nbsp;&nbsp;&nbsp;
            <%rsall.movenext
		  next%></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td valign="top">    <table width="100%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
      <tr bgcolor="#E8F1FF">
        <td align="center"><span class="TableRow2">���ű���</span><span class="TableRow2"> </span></td>
        <td align="center">�������</td>
        <td align="center"><span class="BodyTitle">���ʱ��</span></td>
        <td align="center"><span class="BodyTitle">�޸�</span></td>
        <td align="center"><span class="BodyTitle">ѡ</span></td>
      </tr>
      <%for i=1 to cur_recordcounts
	  set rsxwlb=server.CreateObject("adodb.recordset")
	  rsxwlb.open"select * from xwlb where id="&rsgly("xwlb_id"),cn,1,3%>
      <tr align="center" bgcolor="#E8F1FF">
        <td height="15" align="left"><span class="TableRow2"><%=rsgly("xwbt")%></span></td>
        <td height="15"><span class="TableRow2"><%=rsxwlb("xwlb")%></span></td>
        <td height="15"><span class="TableRow1"><%=rsgly("tjsj")%></span></td>
        <td height="15"><span class="TableRow1"><a href="xwxiugai.asp?id=<%=rsgly("id")%>">�޸�</a></span></td>
        <td height="15"><span class="TableRow1">
          <input type="checkbox" value="<%=rsgly("id")%>" name=id>
</span></td>
      </tr>
      <%rsgly.movenext
next%>
    </table>
      <table width="96%"  border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td align="right"><input type='checkbox' name=chkall onclick='CheckAll(this.form)'>
      ȫѡ
        <input type="submit" name="action" value="ɾ��" onclick="{if(confirm('�ò������ɻָ���\n\nȷʵɾ��ѡ������Ŀ��')){this.document.Prodlist.submit();return true;}return false;}"></td>
        </tr>
      </table></td>
  </tr>
</table>
</form>
<table width="96%"  border="0" align="center" cellpadding="0" cellspacing="8">
  <tr>
    <td align="center"><% 
	if pages>1 then
	response.Write"<a href=xwlb.asp?pages="&pages-1&"&xwlb_id="&xwlb_id&">ǰһҳ</a>&nbsp;"
	else
	response.Write"ǰһҳ&nbsp;"
	end if
	if pages<=rsgly.pagecount-5 then
	start=pages-5
	else
	start=rsgly.pagecount-10
	end if
	if start<=0 then start=1
	for i=start to rsgly.pagecount 
	 if i=pages then
	 response.Write"&nbsp;"&i 
	 else 
	  if i<start+11 then
	  response.Write"&nbsp;<a href=xwlb.asp?pages="&i&"&xwlb_id="&xwlb_id&">"&i&"</a>"
	  else
	  exit for
	 end if
	 end if
	 next 
	 if pages<rsgly.pagecount then
	 response.Write"&nbsp;<a href=xwlb.asp?pages="&pages+1&"&xwlb_id="&xwlb_id&">��һҳ</a>"
	 else
	 response.Write"&nbsp;��һҳ"
	 end if
	 %></td>
  </tr>
</table>
<%

if request("action")="ɾ��" then
	delid=replace(request("id"),"'","")
	call delfeedback()
end if
sub delfeedback()
	if delid="" or isnull(delid) then
    response.write "<script language='javascript'>"
	response.write "alert('����ʧ�ܣ�û��ѡ����ʲ������뵥����ȷ�������أ�');"
	response.write "location.href='xwlb.asp?xwlb_id="&xwlb_id&"';"
	response.write "</script>"
	response.end

	else
		cn.execute("delete from xw where ID in ("&delid&")")
		cn.close
		set cn=nothing

	response.write "<script language='javascript'>"
	response.write "alert('��Ŀɾ���ɹ����뵥����ȷ�������أ�');"
	response.write "location.href='xwlb.asp?xwlb_id="&xwlb_id&"';"
	response.write "</script>"
	response.end

	end if
end sub

sub detailfeedback()
if id="" then
	response.write "<script language='javascript'>"
	response.write "alert('�޴����Ա�ţ��뵥����ȷ�������أ�');"
	response.write "location.href='xwlb.asp?xwlb_id="&xwlb_id&"';"
	response.write "</script>"
	response.end
end if
end sub
%>
