<!--#include file="lianjie.asp"-->
<%leixing=request("leixing")%>
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
            <td width="20%" align="right">ͼƬ˵����</td>
            <td width="80%" style="PADDING-LEFT: 10px"><input name="tpsm" type="text" id="tpsm" size="50">
            [û�пɱ��ֿ�]</td>
          </tr>
          <tr bgcolor="#E8F1FF">
            <td align="right" valign="top">���ͼƬ��</td>
            <td style="PADDING-LEFT: 10px"><input name="tpdz" type="text" id="tpdz" value="images/upfile/lyx.jpg" size="50">
            <input name="Submit1" type="button" onClick="window.open('upfile.asp?formname=myform&editname=tpdz&uppath=images/guanggao&filelx=jpg','','status=no,scrollbars=no,top=20,left=110,width=420,height=165')" value="�� ��">
            [����ߴ硡292��411]</td>
          </tr><tr bgcolor="#E8F1FF">
            <td align="right" valign="middle">��ϸ���ܣ�</td>
            <td style="PADDING-LEFT: 10px"><textarea name="cptpjqsm" style="display:none" cols="1" rows="1">��Ŀ����</textarea><IFRAME ID="eWebEditor1" SRC="edit/ewebeditor.htm?id=cptpjqsm&style=coolblue" FRAMEBORDER="0" SCROLLING="no" WIDTH="550" HEIGHT="350"></IFRAME></td>
          </tr>
          <tr bgcolor="#E8F1FF">
            <td height="30"></td>
            <td height="30" style="PADDING-LEFT: 10px">
              <input type="submit" name="Submit" value="���">
              <input type="reset" name="Submit2" value="����">
            </td>
          </tr>
      </table></td>
    </tr>
  </table>
</form>
</body>
</html>

