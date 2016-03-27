<!--#include file="lianjie.asp"-->
<%leixing=request("leixing")%>
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
</style>
</head>

<body>
<form  name="myform" method="post" action="gsry_add.asp?leixing=<%=leixing%>">
  <table class="tableBorder" width="100%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#FFFFFF">
    <tr>
      <td  align="center" background="images/admin_bg_1.gif" height="25"><span class="style1"><%=leixing%></span></td>
    </tr>
    <tr>
      <td valign="top">
        <table width="100%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#FFFFFF">
          <tr bgcolor="#E8F1FF">
            <td width="20%" align="right">图片说明：</td>
            <td width="80%" style="PADDING-LEFT: 10px"><input name="tpsm" type="text" id="tpsm" size="50">
            [没有可保持空]</td>
          </tr>
          <tr bgcolor="#E8F1FF">
            <td align="right" valign="top">相关图片：</td>
            <td style="PADDING-LEFT: 10px"><input name="tpdz" type="text" id="tpdz" value="images/upfile/lyx.jpg" size="50">
            <input name="Submit1" type="button" onClick="window.open('upfile.asp?formname=myform&editname=tpdz&uppath=images/guanggao&filelx=jpg','','status=no,scrollbars=no,top=20,left=110,width=420,height=165')" value="上 传">
            [建议尺寸　292＊411]</td>
          </tr><tr bgcolor="#E8F1FF">
            <td align="right" valign="middle">详细介绍：</td>
            <td style="PADDING-LEFT: 10px"><textarea name="cptpjqsm" style="display:none" cols="1" rows="1">项目内容</textarea><IFRAME ID="eWebEditor1" SRC="edit/ewebeditor.htm?id=cptpjqsm&style=coolblue" FRAMEBORDER="0" SCROLLING="no" WIDTH="550" HEIGHT="350"></IFRAME></td>
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

