<!--#include file="lianjie.asp"----->
<%set rsxwlb=server.CreateObject("adodb.recordset")
rsxwlb.open"select * from xwlb",cn,1,3%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>后台管理</title>
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
<form name="form" method="post" action="xw_add.asp">
  <table class="tableBorder" width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
    <tr>
      <td  align="center" background="images/admin_bg_1.gif" height="25"><span class="style1">新闻内容添加</span></td>
    </tr>
    <tr>
      <td valign="top">
        <table width="100%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF"><tr bgcolor="#E8F1FF">
            <td width="16%" align="right"><span class="TableRow2">新闻标题</span>：</td>
            <td style="PADDING-LEFT: 10px"><span class="TableRow2">
            <input name="xwbt" type="text" id="xwbt" size="60">
</span></td>
            </tr><tr bgcolor="#E8F1FF">
            <td width="16%" align="right"><span class="TableRow2">新闻类别</span>：</td>
            <td style="PADDING-LEFT: 10px"><span class="TableRow2">
              <select name="xwlb_id" id="xwlb_id">
			  <%for i=1 to rsxwlb.recordcount%>
                <option value="<%=rsxwlb("id")%>"><%=rsxwlb("xwlb")%></option>
                <%rsxwlb.movenext
				next%>
              </select>
</span></td>
            </tr>
          <tr bgcolor="#E8F1FF">
            <td width="16%" align="right"><span class="TableRow2">新闻内容</span>：</td>
            <td style="PADDING-LEFT: 10px"><span class="TableRow2"><textarea name="xwnr" style="display:none" cols="1" rows="1"></textarea><IFRAME ID="eWebEditor1" SRC="edit/ewebeditor.htm?id=xwnr&style=coolblue" FRAMEBORDER="0" SCROLLING="no" WIDTH="550" HEIGHT="350"></IFRAME>
            </span></td>
          </tr> 
          <tr bgcolor="#E8F1FF">
            <td height="30"></td>
            <td height="30" style="PADDING-LEFT: 10px">
              <input type="submit" name="Submit" value="添加">
              <input type="reset" name="Submit2" value="重置">
            </td>
          </tr>
      </table></td>
    </tr>
  </table>
</form>
</body>
</html>
