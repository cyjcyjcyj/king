<HTML>
<HEAD>
<TITLE>eWebEditor���߱༭�� - ʹ������</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<style>
body,td,input,textarea {font-size:9pt}
</style>
</HEAD>
<BODY>




<p><b>eWebEditor ��׼����ʾ����</b></p>
	<Script Language=JavaScript>
	// ���ϴ�ͼƬ���ļ�ʱ����������������ͼƬ·�����ɸ���ʵ����Ҫ���Ĵ˺���
	function doChange(objText, objDrop){
		if (!objDrop) return;
		var str = objText.value;
		var arr = str.split("|");
		var nIndex = objDrop.selectedIndex;
		objDrop.length=1;
		for (var i=0; i<arr.length; i++){
			objDrop.options[objDrop.length] = new Option(arr[i], arr[i]);
		}
		objDrop.selectedIndex = nIndex;
	}
	function selectpic(a,b){
	var avalue=document.getElementById(a).value;
		if (avalue!="") {
		alert(avalue);	
		}else {alert("���ϴ�ͼƬ��")}

	}
	</Script>
<FORM method="POST" name="myform" action="submit.asp">
<TABLE border="0" cellpadding="2" cellspacing="1">
<TR>
  <TD>&nbsp;</TD>
  <TD>&nbsp;</TD>
</TR>
<TR>
  <TD>ȡԴ�ļ���</TD>
  <TD><input type="text" name="d_originalfilename"></TD>
</TR>
<TR>
  <TD>ȡ������ļ����������Ҫ��·������������򣬿���������ı������onchange�¼�</TD>
  <TD><input type="text" name="d_savefilename"></TD>
</TR>
<TR>
  <TD>ȡ������ļ�������·������ʹ�ô�·�������������</TD>
  <TD><input type="text" name="d_savepathfilename"></TD>
</TR>
<TR>
  <TD>1</TD>
  <TD><input name="textfield1" type="text" id="textfield1">
    <input type="button" name="Button" value="ѡ��ͼƬ" onClick='selectpic("d_savepathfilename","textfield1")'></TD>
</TR>
<TR>
  <TD>2</TD>
  <TD><input type="text" name="textfield2">
    <input type="button" name="Button2" value="ѡ��ͼƬ" onClick='selectpic("d_savepathfilename","textfield2")'></TD>
</TR>
<TR>
  <TD>&nbsp;</TD>
  <TD><select name="d_picture" size=1><option value=''>��</option></select></TD>
</TR>
<TR>
	<TD>�༭���ݣ�</TD>
	<TD>
		<INPUT type="hidden" name="content1" value="<P align=center><FONT color=#ff0000><FONT face='Arial Black' size=7><STRONG>eWeb<FONT color=#0000ff>Editor</FONT><FONT color=#000000><SUP>&trade;</SUP></FONT></STRONG></FONT></FONT></P><P align=right><FONT style='BACKGROUND-COLOR: #ffff00' color=#ff0000><STRONG>Version 3.6 ��������רҵ��</STRONG></FONT></P><P>����ʽΪϵͳĬ����ʽ��coolblue������ѵ��ÿ��550px���߶�350px��</FONT></P><P>����ΪһЩ�߼����ù��ܵ����ӣ�������ýű�����ĶԱ༭������һϵ�в�����</P><P><B><TABLE borderColor=#ff9900 cellSpacing=2 cellPadding=3 align=center bgColor=#ffffff border=1 heihgt=''><TBODY><TR><TD bgColor=#00ff00><STRONG>&nbsp;������Щ���ݣ���û�д�����ʾ��˵����װ�Ѿ���ȷ��ɣ�</STRONG></TD></TR></TBODY></TABLE></B></P>">
		<IFRAME ID="eWebEditor1" src="../ewebeditor.htm?id=content1&style=coolblue&originalfilename=d_originalfilename&savefilename=d_savefilename&savepathfilename=d_savepathfilename" frameborder="0" scrolling="no" width="550" height="350"></IFRAME>	</TD>
</TR>
<TR>
	<TD colspan=2 align=right>
	<INPUT type=submit name=b1 value="�ύ"> 
	<INPUT type=reset name=b2 value="����"> 
	<INPUT type=button name=b3 value="�鿴Դ�ļ�" onClick="location.replace('view-source:'+location)">	</TD>
</TR>
<TR>
	<TD>ȡ�����ݣ�</TD>
	<TD><TEXTAREA cols=50 rows=5 id=myTextArea style="width:550px">�����ȡֵ����ť����һ��Ч����</TEXTAREA></TD>
</TR>
<TR>
	<TD colspan=2 align=right>
	<INPUT type=button name=b4 value="ȡֵ" onClick="myTextArea.value=eWebEditor1.getHTML()"> 
	<INPUT type=button name=b5 value="��ֵ" onClick="eWebEditor1.setHTML('<b>Hello My World!</b>')">
	<INPUT type=button name=b6 value="����״̬" onClick="eWebEditor1.setMode('CODE')">
	<INPUT type=button name=b7 value="���״̬" onClick="eWebEditor1.setMode('EDIT')">
	<INPUT type=button name=b7 value="Image" onClick="eWebEditor1.showDialog('img.htm', true)">
	<INPUT type=button name=b8 value="�ı�״̬" onClick="eWebEditor1.setMode('TEXT')">
	<INPUT type=button name=b8 value="Ԥ��״̬" onClick="eWebEditor1.setMode('VIEW')">
	<INPUT type=button name=b9 value="��ǰλ�ò���" onClick="eWebEditor1.insertHTML('This is Insert Function!')">
	<INPUT type=button name=b10 value="β��׷��" onClick="eWebEditor1.appendHTML('This is Append Function!')">	</TD>
</TR>
</TABLE>
</FORM>





<hr>





<Script Language=JavaScript>
function eWebEditorPopUp(form, field, width, height) {
	window.open("../popup.htm?style=popup&form="+form+"&field="+field, "", "width="+width+",height="+height+",toolbar=no,location=no,status=no,menubar=no,scrollbars=no,resizable=no");
}
</Script>

<p><b>eWebEditor ��������ʾ����</b></p>

<TABLE border="0" cellpadding="2" cellspacing="1">
<TR>
	<TD>�༭���ݣ�</TD>
	<TD>
		<FORM ACTION="" METHOD="" NAME="myform2">
		<TEXTAREA NAME="myField" COLS="50" ROWS="10" style="width:550px">&lt;i&gt;eWebEditor ��������ʾ��&lt;/i&gt;</TEXTAREA><br>
		<INPUT TYPE="BUTTON" NAME="btn" VALUE="HTML�༭" ONCLICK="eWebEditorPopUp('myform2', 'myField', 580, 380)">
		</FORM>
	</TD>
</TR>
<TR>
	<TD align=right colspan=2>
	<INPUT type=button name=b3 value="�鿴Դ�ļ�" onClick="location.replace('view-source:'+location)"> 
	</TD>
</TR>
</TABLE>
<p><b>�����HTML�༭����ť���ڵ������ڱ༭һЩ���ݣ�Ȼ��㡰���淵�ء���ť����һ��Ч����</b></p>



</BODY>
</HTML>
