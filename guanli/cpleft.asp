<meta http-equiv="Content-Type" content="text/html; charset=gb2312"><%set rsgly=server.CreateObject("adodb.recordset")
rsgly.open"select * from dlb",cn,1,3%><style type="text/css">
<!--
body {
	margin-left: 0px;
	margin-top: 0px;
	margin-right: 0px;
	margin-bottom: 0px;
}
.xwleft {
	border: 1px solid #999999;
}
-->
</style><table id="__01" width="165" height="437" border="0" cellpadding="0" cellspacing="0">
	<tr>
		<td width="165" height="25" valign="top">
			<img src="images/cplbzb.gif" width="165" height="25" alt=""></td>
	</tr>
	<tr>
		<td height="70" align="center" valign="top"><table width="100%" height="61"  border="0" cellpadding="0" cellspacing="0" bordercolor="#000000">
          <%for i=1 to rsgly.recordcount%>
          <tr>
            <td bgcolor="#D6D3CE"><img src="guanli/images/d.gif" width="15" height="15"><a href="cpzs.asp?dlb_id=<%=rsgly("id")%>"><%=rsgly("dlb")%></a></td>
          </tr>
          <tr>
            <td height="43"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                <%set rsxlb=server.CreateObject("adodb.recordset")
			rsxlb.open"select * from xlb where dlb_id="&rsgly("id"),cn,1,3
			for j=1 to rsxlb.recordcount%>
                <tr bgcolor="#EFEFEF">
                  <td width="8%">&nbsp;</td>
                  <td width="92%"><img src="guanli/images/x.gif" width="15" height="15"><a href="cp_view.asp?xlb_id=<%=rsxlb("id")%>&dlb_id=<%=rsgly("id")%>"><%=rsxlb("xlb")%></a></td>
                </tr>
                <%rsxlb.movenext
			  next%>
              </table>
                <br></td>
          </tr>
          <%rsgly.movenext
		  next%>
        </table></td>
	</tr>
	<tr><form  method="post" action="xw_search.asp">
		<td height="25">
			<img src="images/xw_search.gif" width="165" height="25" alt=""></td>
	</tr>
	<TR>
      <TD align=center bgColor=#efefef height=25>      关键字      </TD>
  </TR>
	<TR>
	  <TD align=middle bgColor=#efefef height=25><INPUT name=key 
                        class=xwleft id="key" size=19></TD>
	  </TR>
	<TR>
      <TD align=center bgColor=#efefef height=30>
	  <INPUT class=input type=submit value=搜索 name=Submit></TD></form>
  </TR>
	<tr>
		<td height="202" align="right" valign="top">&nbsp;</td>
	</tr>
	
</table>
