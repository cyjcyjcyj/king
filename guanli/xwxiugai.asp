<!--#include file="lianjie.asp"----->
<%id=cint(trim(request("id")))
set rsgly=server.CreateObject("adodb.recordset")
rsgly.open"select * from xw where id="&id,cn,1,3
set rsxwlb=server.CreateObject("adodb.recordset")
rsxwlb.open"select * from xwlb where id="&rsgly("xwlb_id"),cn,1,3
set rsall=server.CreateObject("adodb.recordset")
rsall.open"select * from xwlb",cn,1,3
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>��̨����</title>
<style type="text/css">
<!--
body,td,th {
	font-size: 12px;
}
.style1 {	color: #DBBF90;
	font-weight: bold;
}
-->
</style></head>

<body>
<form name="form" method="post" action="xw_xiugai.asp?id=<%=id%>">
  <table class="tableBorder" width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
    <tr>
      <td  align="center" background="images/admin_bg_1.gif" height="25"><span class="style1">���������޸�</span></td>
    </tr>
    <tr>
      <td valign="top">
        <table width="100%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF"><tr bgcolor="#E8F1FF">
            <td width="16%" align="right"><span class="TableRow2">���ű���</span>��</td>
            <td style="PADDING-LEFT: 10px"><span class="TableRow2">
            <input name="xwbt" type="text" id="xwbt" size="60" value="<%=rsgly("xwbt")%>">
</span></td>
            </tr><tr bgcolor="#E8F1FF">
            <td width="16%" align="right"><span class="TableRow2">�������</span>��</td>
            <td style="PADDING-LEFT: 10px"><span class="TableRow2">
            <select name="xwlb_id" id="xwlb_id">
              <option value="<%=rsxwlb("id")%>"><%=rsxwlb("xwlb")%></option>
			  <%for i=1 to rsall.recordcount%>
               <option value="<%=rsall("id")%>"><%=rsall("xwlb")%></option>
            <%rsall.movenext
			next%>
            </select>
            </span></td>
            </tr>
          <tr bgcolor="#E8F1FF">
            <td width="16%" align="right"><span class="TableRow2">��������</span>��</td>
            <td style="PADDING-LEFT: 10px"><span class="TableRow2"><textarea name="xwnr" style="display:none" cols="1" rows="1"><%=rsgly("xwnr")%></textarea><IFRAME ID="eWebEditor1" SRC="edit/ewebeditor.htm?id=xwnr&style=coolblue" FRAMEBORDER="0" SCROLLING="no" WIDTH="550" HEIGHT="350"></IFRAME>
            </span></td>
          </tr> 
          <tr bgcolor="#E8F1FF">
            <td height="30"></td>
            <td height="30" style="PADDING-LEFT: 10px">
              <input type="submit" name="Submit" value="�޸�">
              <input type="reset" name="Submit2" value="����">
            </td>
          </tr>
      </table></td>
    </tr>
  </table>
</form>
</body>
</html>
<%rsgly.close
set rsgly=nothing
rsxwlb.close
set rsxwlb=nothing
rsall.close
set rsall=nothing%>
