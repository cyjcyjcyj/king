<!--#include file="conn.asp"-->
<%set rsdd=server.CreateObject("adodb.recordset")
rsdd.open"select * from dd ",cn,1,3%>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">

<style type="text/css">
<!--
body,td,th {
	font-size: 12px;
}
.style1 {
	color: #DBBF90;
	font-weight: bold;
}
body {
	background-color: #E8F1FF;
}
-->
</style><table width="100%" height="22"  border="0" align="center" cellpadding="0" cellspacing="0" background="images/admin_bg_1.gif">
  <TR align="center" bgcolor="#DEEFFF" background="images/admin_bg_1.gif">
      <TD height="13" colspan="2" bgcolor="#FFFFFF" background="images/admin_bg_1.gif"><span class="style1">�����б�</span></TD>
  </TR>
</table>
<%for i=1 to rsdd.recordcount%>

<table class="unnamed15" id="AutoNumber1" 
                  style="BORDER-COLLAPSE: collapse" cellspacing="1" cellpadding="0" 
                  align="center" border="0">
  <tbody>
   
    
    
    <tr bgcolor="#E8F1FF">
      <td width="107" height="30" align="left">������</td>
      <td width="746" height="30" align="left"><%=rsdd("s1")%></td>
    </tr>
    <tr bgcolor="#E8F1FF">
      <td width="107" height="27" align="left"><p>��λ��<br />
      </p></td>
      <td height="27" align="left"><%=rsdd("s2")%></td>
    </tr>
    <tr bgcolor="#E8F1FF">
      <td height="27" align="left">�绰��</td>
      <td height="27" align="left"><%=rsdd("s3")%></td>
    </tr>
    <tr bgcolor="#E8F1FF">
      <td height="27" align="left">���棺</td>
      <td height="27" align="left"><%=rsdd("s4")%></td>
    </tr>
    <tr bgcolor="#E8F1FF">
      <td height="27" align="left">���䣺</td>
      <td height="27" align="left"><%=rsdd("s5")%></td>
    </tr>
    <tr bgcolor="#E8F1FF">
      <td height="27" align="left">��ů�����</td>
      <td height="27" align="left"><%=rsdd("s6")%></td>
    </tr>
    <tr bgcolor="#E8F1FF">
      <td width="107" height="27" align="left"><p>��ƷӦ�ã�<br />
      </p></td>
      <td height="27" align="left"><%=rsdd("s7")%></td>
    </tr>
    <tr bgcolor="#E8F1FF">
      <td width="107" height="27" align="left">Ԥ�ڵ繦�ʣ�</td>
      <td height="27" align="left"><%=rsdd("s8")%></td>
    </tr>
    <tr>
      <td align="left" colspan="2" height="25"><strong>ѡ ��  �� �� Ʒ</strong></td>
    </tr>
    <tr bgcolor="#E8F1FF">
      <td height="34" align="left" nowrap="nowrap">��Ʒ��</td>
      <td align="left"><%=rsdd("s9")%></td>
    </tr>
    <tr>
      <td height="17" align="left"><p><br />
      </p></td>
      <td align="left">&nbsp;</td>
    </tr>
    <tr bgcolor="#E8F1FF">
      <td height="27" align="left">������</td>
      <td height="27" align="left"><%=rsdd("s10")%></td>
    </tr>
    <tr>
      <td align="left" colspan="2" height="25"><strong>�� ע</strong></td>
    </tr>
    <tr>
      <td width="107" height="33" align="middle" valign="center"><div align="center">��ע��</div></td>
      <td><%=rsdd("s11")%></td>
    </tr>
    <tr bgcolor="#E8F1FF">
      <td height="40" colspan="2" align="center"><a href="dd_del.asp?id=<%=rsdd("id")%>">ɾ��</a><br />
      <hr /></td>
    </tr>
  </tbody>
</table>
<%rsdd.movenext
next%>
