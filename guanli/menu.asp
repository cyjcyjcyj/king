<%

 lujing="shuju/woai#asp#shuju.mdb"
set cn=server.CreateObject("adodb.connection")
cn.open"provider=microsoft.jet.oledb.4.0;data source="&server.MapPath(""&lujing&"")
set rsgsxm1=server.CreateObject("adodb.recordset")
rsgsxm1.open"select * from gsxmnr where leixing='��ҵ���'",cn,1,3
set rsgsxm2=server.CreateObject("adodb.recordset")
rsgsxm2.open"select * from gsxmnr where leixing='�豸չʾ'",cn,1,3
set rsgsxm3=server.CreateObject("adodb.recordset")
rsgsxm3.open"select * from gsxmnr where leixing='��������'",cn,1,3
set rsgsxm4=server.CreateObject("adodb.recordset")
rsgsxm4.open"select * from gsxmnr where leixing='��������'",cn,1,3
set rsgsxm5=server.CreateObject("adodb.recordset")
rsgsxm5.open"select * from gsxmnr where leixing='��������'",cn,1,3

%>

<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>����ҳ��</title>
<style type=text/css>
body  { background:#799AE1; margin:0px; font:9pt ����; }
table  { border:0px; }
td  { font:normal 12px ����; }
img  { vertical-align:bottom; border:0px; }
a  { font:normal 12px ����; color:#000000; text-decoration:none; }
a:hover  { color:#428EFF;text-decoration:underline; }
.sec_menu  { border-left:1px solid white; border-right:1px solid white; border-bottom:1px solid white; overflow:hidden; background:#D6DFF7; }
.menu_title  { }
.menu_title span  { position:relative; top:2px; left:8px; color:#215DC6; font-weight:bold; }
.menu_title2  { }
.menu_title2 span  { position:relative; top:2px; left:8px; color:#428EFF; font-weight:bold; }
</style>
<SCRIPT language=javascript1.2>
function showsubmenu(sid)
{
whichEl = eval("submenu" + sid);
if (whichEl.style.display == "none")
{
eval("submenu" + sid + ".style.display=\"\";");
}
else
{
eval("submenu" + sid + ".style.display=\"none\";");
}
}
</SCRIPT>
			<table cellpadding=0 cellspacing=0 width=158 align=center>
				<tr> 
					<td height=38 class=menu_title onMouseOver=this.className='menu_title2'; onMouseOut=this.className='menu_title'; background=images/title.gif></td>
				</tr>
				<tr>
					<td height=25 background=images/title_bg_quit.gif>&nbsp;
						<a href="guanli.asp" target=_top><b>������ҳ</b></a> | <a href="tuichu.asp" target="_top"><b>�˳�</b></a>
					</td>
				</tr>
			</table>
			&nbsp; 
			<table cellpadding=0 cellspacing=0 width=158 align=center>
				<tr>
				  <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/admin_left_2.gif" id=menuTitle1 onclick="showsubmenu(1)"> 
				    <b><span>����Ա����</span></b>
				  </td>
				</tr>
				<tr>
					<td style="display:" id='submenu1'>
						<div class=sec_menu style="width:158">
						    <table cellpadding=0 cellspacing=0 align=center width=150>
						      <tr> 
						        <td height=20>�� <a href=glyadd.asp target=right>����Ա���</a> </td>
						      </tr>
						      <tr> 
						        <td height=20>�� <a href=gly_list.asp target=right>����Ա�б�</a> </td>
						      </tr>
						      <!--<tr> 
						        <td height=20>�� <a href=shihemanage.asp?action=no target=right>�ʺ��������</a></td>
						      </tr>-->
						    </table>
						</div>
					</td>
				</tr>
			</table>&nbsp;			&nbsp;
			<table cellpadding=0 cellspacing=0 width=158 align=center>
              <tr>
                <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/admin_left_7.gif" id=menuTitle1 onclick="showsubmenu(1331)"> <span>��������</span> </td>
              </tr>
              <tr>
                <td style="display:" id='submenu1331'>
                  <div class=sec_menu style="width:158">
                    <table cellpadding=0 cellspacing=0 align=center width=150>
                      <%for i=1 to rsgsxm1.recordcount%>
                      <tr>
                        <td height=20>�� <a href=gsxmnrxiugai.asp?id=<%=rsgsxm1("id")%> target=right><%=rsgsxm1("gsxm")%></a></td>
                      </tr>
                      <%rsgsxm1.movenext
								next
								rsgsxm1.close
								set rsgsxm1=nothing
								%>
                    </table>
                </div></td>
              </tr>
            </table>&nbsp;&nbsp;
            <table cellpadding=0 cellspacing=0 width=158 align=center>
              <tr>
                <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/admin_left_7.gif" id=menuTitle1 onclick="showsubmenu(12)"> <span>��Ʒ����</span> </td>
              </tr>
              <tr>
                <td style="display:" id='submenu12'>
                  <div class=sec_menu style="width:158">
                    <table cellpadding=0 cellspacing=0 align=center width=150>
                    <tr>
									<td height=20>�� <a href=cplb_list.asp target=right>��Ʒ����б�</a></td>
								</tr>
                      <tr>
                        <td height=20>�� <a href=dckt_list.asp target=right>��Ʒ�б�</a></td>
                      </tr>
                      <tr>
                        <td height=20>�� <a href=dcktadd.asp target=right>��Ӳ�Ʒ</a></td>
                      </tr>
                    </table>
                  </div></td>
              </tr>
            </table>&nbsp;&nbsp;            &nbsp;
			<table cellpadding=0 cellspacing=0 width=158 align=center>
				<tr>
				  <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/admin_left_7.gif" id=menuTitle1 onclick="showsubmenu(131)"> 
				    <span>��˾��̬</span>
				  </td>
				</tr>
				<tr>
					<td style="display:" id='submenu131'>
						<div class=sec_menu style="width:158">
							<table cellpadding=0 cellspacing=0 align=center width=150>
								<tr>
									<td height=20>�� <a href=xw_list.asp target=right>���������б�</a></td>
								</tr>
								<tr> 
									<td height=20>�� <a href=xwadd.asp target=right>�������</a></td>
								</tr>
							</table>
						</div>
					</td>
				</tr>
			</table>
			������			&nbsp;&nbsp;&nbsp;&nbsp;            &nbsp;
            <table cellpadding=0 cellspacing=0 width=158 align=center>
              <tr>
                <td height=25 class=menu_title onmouseover=this.className='menu_title2'; onmouseout=this.className='menu_title'; background="images/admin_left_4.gif" id=menuTitle1 onclick="showsubmenu(1555)"> <span>���Թ���</span> </td>
              </tr>
              <tr>
                <td style="display:" id='submenu1555'>
                  <div class=sec_menu style="width:158">
                    <table cellpadding=0 cellspacing=0 align=center width=150>
                      <tr>
                        <td height=20>��<a href=book_admin.asp target=right>�����б�</a></td>
                      </tr>
                    </table>
                  </div></td>
              </tr>
            </table>&nbsp;       &nbsp;&nbsp;&nbsp;<br>
			<p>&nbsp;
</p>
