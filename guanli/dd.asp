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
      <TD height="13" colspan="2" bgcolor="#FFFFFF" background="images/admin_bg_1.gif"><span class="style1">订单列表</span></TD>
  </TR>
</table>
<%for i=1 to rsdd.recordcount%>

<table class="unnamed15" id="AutoNumber1" 
                  style="BORDER-COLLAPSE: collapse" cellspacing="1" cellpadding="0" 
                  align="center" border="0">
  <tbody>
   
    
    
    <tr bgcolor="#E8F1FF">
      <td width="107" height="30" align="left">姓名：</td>
      <td width="746" height="30" align="left"><%=rsdd("s1")%></td>
    </tr>
    <tr bgcolor="#E8F1FF">
      <td width="107" height="27" align="left"><p>单位：<br />
      </p></td>
      <td height="27" align="left"><%=rsdd("s2")%></td>
    </tr>
    <tr bgcolor="#E8F1FF">
      <td height="27" align="left">电话：</td>
      <td height="27" align="left"><%=rsdd("s3")%></td>
    </tr>
    <tr bgcolor="#E8F1FF">
      <td height="27" align="left">传真：</td>
      <td height="27" align="left"><%=rsdd("s4")%></td>
    </tr>
    <tr bgcolor="#E8F1FF">
      <td height="27" align="left">邮箱：</td>
      <td height="27" align="left"><%=rsdd("s5")%></td>
    </tr>
    <tr bgcolor="#E8F1FF">
      <td height="27" align="left">采暖面积：</td>
      <td height="27" align="left"><%=rsdd("s6")%></td>
    </tr>
    <tr bgcolor="#E8F1FF">
      <td width="107" height="27" align="left"><p>产品应用：<br />
      </p></td>
      <td height="27" align="left"><%=rsdd("s7")%></td>
    </tr>
    <tr bgcolor="#E8F1FF">
      <td width="107" height="27" align="left">预期电功率：</td>
      <td height="27" align="left"><%=rsdd("s8")%></td>
    </tr>
    <tr>
      <td align="left" colspan="2" height="25"><strong>选 择  的 产 品</strong></td>
    </tr>
    <tr bgcolor="#E8F1FF">
      <td height="34" align="left" nowrap="nowrap">产品：</td>
      <td align="left"><%=rsdd("s9")%></td>
    </tr>
    <tr>
      <td height="17" align="left"><p><br />
      </p></td>
      <td align="left">&nbsp;</td>
    </tr>
    <tr bgcolor="#E8F1FF">
      <td height="27" align="left">数量：</td>
      <td height="27" align="left"><%=rsdd("s10")%></td>
    </tr>
    <tr>
      <td align="left" colspan="2" height="25"><strong>备 注</strong></td>
    </tr>
    <tr>
      <td width="107" height="33" align="middle" valign="center"><div align="center">备注：</div></td>
      <td><%=rsdd("s11")%></td>
    </tr>
    <tr bgcolor="#E8F1FF">
      <td height="40" colspan="2" align="center"><a href="dd_del.asp?id=<%=rsdd("id")%>">删除</a><br />
      <hr /></td>
    </tr>
  </tbody>
</table>
<%rsdd.movenext
next%>
