<!--#include file = "private.asp"-->

<%
'######################################
' eWebEditor v3.00 - Advanced online web based WYSIWYG HTML editor.
' Copyright (c) 2003-2004 eWebSoft.com
'
' For further information go to http://www.ewebsoft.com/
' This copyright notice MUST stay intact for use.
'######################################
%>

<%

sPosition = sPosition & "��̨������ҳ"

Call Header()
Call Content()
Call Footer()




Sub Content()
	%>

	<table border=0 cellspacing=1 align=center class=navi>
	<tr><th><%=sPosition%></th></tr>
	</table>

	<br>

	<table border=0 cellspacing=1 align=center class=list>
	<tr><th colspan=2>eWebEditor 3.6 ��Ȩ��������ϵ����������֧��</th></tr>
	<tr>
		<td width="15%">����汾��</td>
		<td width="85%">eWebEditor Version 3.6 for ASP ��������רҵ��</td>
	</tr>
	<tr>
		<td width="15%">��Ȩ���У�</td>
		<td width="85%"><a href="http://www.ewebsoft.com" target="_blank">eWebSoft.com</a>&nbsp;&nbsp;�ѻ�ù��Ұ�Ȩ�ְ䷢�ġ�������������Ȩ�Ǽ�֤�顷,�ǼǺ�:2004SR06549</td>
	</tr>
	<tr>
		<td width="15%">����������</td>
		<td width="85%">eWeb�����Ŷ�</td>
	</tr>
	<tr>
		<td width="15%">��ҳ��ַ��</td>
		<td width="85%"><a href="http://www.eWebSoft.com" target=_blank>http://www.eWebSoft.com</a>&nbsp;&nbsp;&nbsp;<a href="http://www.webasp.net" target=_blank>http://www.webasp.net</a></td>
	</tr>
	<tr>
		<td width="15%">��Ʒ���ܣ�</td>
		<td width="85%"><a href="http://www.eWebSoft.com/Product/eWebEditor/" target=_blank>http://www.eWebSoft.com/Product/eWebEditor/</a></td>
	</tr>
	<tr>
		<td width="15%">��̳��ַ��</td>
		<td width="85%"><a href="http://bbs.webasp.net" target=_blank>http://bbs.webasp.net</a></td>
	</tr>
	<tr>
		<td width="15%">��ϵ��ʽ��</td>
		<td width="85%">OICQ:589808&nbsp;&nbsp;&nbsp;&nbsp;Email:<a href="mailto:service@ewebsoft.com">service@ewebsoft.com</a></td>
	</tr>
	</table>

	<br>

	<table border=0 cellspacing=1 align=center class=list>
	<tr><th colspan=2>���������йز���</th><th colspan=2>���֧���йز���</th></tr>
	<tr>
		<td width="15%">����������</td>
		<td width="45%"><%=Request.ServerVariables("SERVER_NAME")%></td>
		<td width="20%">ADO ���ݶ���</td>
		<td width="20%"><%=Get_ObjInfo("adodb.connection", 1)%></td>
	</tr>
	<tr>
		<td width="15%">������IP��</td>
		<td width="45%"><%=Request.ServerVariables("LOCAL_ADDR")%></td>
		<td width="20%">FSO �ı��ļ���д��</td>
		<td width="20%"><%=Get_ObjInfo("scripting.filesystemobject", 0)%></td>
	</tr>
	<tr>
		<td width="15%">�������˿ڣ�</td>
		<td width="45%"><%=Request.ServerVariables("SERVER_PORT")%></td>
		<td width="20%">Stream �ļ�����</td>
		<td width="20%"><%=Get_ObjInfo("Adodb.Stream", 0)%></td>
	</tr>
	<tr>
		<td width="15%">������ʱ�䣺</td>
		<td width="45%"><%=Now()%></td>
		<td width="20%">Jmail �ʼ��շ���</td>
		<td width="20%"><%=Get_ObjInfo("JMail.SMTPMail", 1)%></td>
	</tr>
	<tr>
		<td width="15%">IIS�汾��</td>
		<td width="45%"><%=Request.ServerVariables("SERVER_SOFTWARE")%></td>
		<td width="20%">ASPmail ���ţ�</td>
		<td width="20%"><%=Get_ObjInfo("SMTPsvg.Mailer", 1)%></td>
	</tr>
	<tr>
		<td width="15%">����������ϵͳ��</td>
		<td width="45%"><%=Request.ServerVariables("OS")%></td>
		<td width="20%">CDONTS ����SMTP���ţ�</td>
		<td width="20%"><%=Get_ObjInfo("CDONTS.NewMail", 1)%></td>
	</tr>
	<tr>
		<td width="15%">�ű���ʱʱ�䣺</td>
		<td width="45%"><%=Server.ScriptTimeout%> ��</td>
		<td width="20%">LyfUpload �ϴ������</td>
		<td width="20%"><%=Get_ObjInfo("LyfUpload.UploadFile", 1)%></td>
	</tr>
	<tr>
		<td width="15%">վ������·����</td>
		<td width="45%"><%=request.ServerVariables("APPL_PHYSICAL_PATH")%></td>
		<td width="20%">AspUpload �ϴ������</td>
		<td width="20%"><%=Get_ObjInfo("Persits.Upload.1", 1)%></td>
	</tr>
	<tr>
		<td width="15%">������CPU������</td>
		<td width="45%"><%=Request.ServerVariables("NUMBER_OF_PROCESSORS")%> ��</td>
		<td width="20%">SA-FileUp �ϴ������</td>
		<td width="20%"><%=Get_ObjInfo("SoftArtisans.FileUp", 1)%></td>
	</tr>
	<tr>
		<td width="15%">�������������棺</td>
		<td width="45%"><%=ScriptEngine & "/" & ScriptEngineMajorVersion & "." & ScriptEngineMinorVersion & "." & ScriptEngineBuildVersion %></td>
		<td width="20%">AspJpeg ͼ���������</td>
		<td width="20%"><%=Get_ObjInfo("Persits.Jpeg",1)%></td>
	</tr>
	</table>

	
	<%
End Sub

Function Get_ObjInfo(obj, ver)
	On Error Resume Next
	Dim objTest, sTemp
	Set objTest = Server.CreateObject(obj)
	If Err.Number <> 0 Then
		Err.Clear
		Get_ObjInfo = "<font class=red><b>��</b></font>&nbsp;<font class=gray>��֧��</font>"
	Else
		sTemp = ""
		If ver = 1 Then
			sTemp = objTest.version
			If IsNull(sTemp) Then sTemp = objTest.about
			sTemp = Replace(sTemp, "Version", "")
			sTemp = "&nbsp;<font class=tims><font class=blue>" & sTemp & "</font></font>"
		End If
		Get_ObjInfo = "<b>��</b>&nbsp;<font class=gray>֧��</font>" & sTemp
	End If
	Set objTest = Nothing
	If Err.Number <> 0 Then Err.Clear
End Function

%>