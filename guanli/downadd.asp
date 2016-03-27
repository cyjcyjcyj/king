<!--#include file="lianjie.asp"----->
<%set rsxwlb=server.CreateObject("adodb.recordset")
rsxwlb.open"select * from class",cn,1,3%>
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
<form name="form" method="post" action="down_add.asp">
  <table class="tableBorder" width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
    <tr>
      <td  align="center" background="images/admin_bg_1.gif" height="25"><span class="style1">下载内容添加</span></td>
    </tr>
    <tr>
      <td valign="top">
        <table width="100%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF"><tr bgcolor="#E8F1FF">
            <td width="16%" align="right"><span class="TableRow2">下载文件标题</span>：</td>
            <td style="PADDING-LEFT: 10px"><span class="TableRow2">
            <input name="btmc" type="text" id="btmc" size="60">
</span></td>
            </tr><tr bgcolor="#E8F1FF">
            <td width="16%" align="right"><span class="TableRow2">英文下载文件标题</span>：</td>
            <td style="PADDING-LEFT: 10px"><span class="TableRow2">
            <input name="enbtmc" type="text" id="enbtmc" size="60">
</span></td>
            </tr><tr bgcolor="#E8F1FF">
            <td width="16%" align="right"><span class="TableRow2">下载类别</span>：</td>
            <td style="PADDING-LEFT: 10px"><span class="TableRow2">
              <select name="class_id" id="class_id">
			  <%for i=1 to rsxwlb.recordcount%>
                <option value="<%=rsxwlb("id")%>"><%=rsxwlb("xxlb")%></option>
                <%rsxwlb.movenext
				next%>
              </select>
</span></td>
            </tr>
          <tr bgcolor="#E8F1FF">
            <td width="16%" align="right"><span class="TableRow2">下载内容</span>：</td>
            <td style="PADDING-LEFT: 10px"><span class="TableRow2"><textarea name="nrjs" style="display:none" cols="1" rows="1"></textarea><IFRAME ID="eWebEditor1" SRC="edit/ewebeditor.htm?id=nrjs&style=coolblue" FRAMEBORDER="0" SCROLLING="no" WIDTH="550" HEIGHT="350"></IFRAME>
            </span></td>
          </tr>  <tr bgcolor="#E8F1FF">
            <td width="16%" align="right"><span class="TableRow2">英文下载内容</span>：</td>
            <td style="PADDING-LEFT: 10px"><span class="TableRow2"><textarea name="ennrjs" style="display:none" cols="1" rows="1"></textarea><IFRAME ID="eWebEditor1" SRC="edit/ewebeditor.htm?id=ennrjs&style=coolblue" FRAMEBORDER="0" SCROLLING="no" WIDTH="550" HEIGHT="350"></IFRAME>
            </span></td>
          </tr> <tr bgcolor="#E8F1FF">
            <td width="16%" align="right"><span class="TableRow2">下载地址一</span>：</td>
            <td style="PADDING-LEFT: 10px"><span class="TableRow2">
              <input name="down1" type="text" id="down12" size="60">            
              <input name="Submit1" type="button" onClick="window.open('upfile.asp?formname=form&editname=down1&uppath=images/guanggao&filelx=jpg','','status=no,scrollbars=no,top=20,left=110,width=420,height=165')" value="上 传">
[此处从本地上传]            </span></td>
            </tr><tr bgcolor="#E8F1FF">
            <td width="16%" align="right"><span class="TableRow2">下载地址二</span>：</td>
            <td style="PADDING-LEFT: 10px"><span class="TableRow2">
            <input name="down2" type="text" id="down2" size="60">
[此处从网上链接] </span></td>
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
