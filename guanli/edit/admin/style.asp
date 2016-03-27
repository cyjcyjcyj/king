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

Dim sStyleID, sStyleName, sStyleDir, sStyleCSS, sStyleUploadDir, sStyleWidth, sStyleHeight, sStyleMemo, nStyleIsSys, sStyleStateFlag, sStyleDetectFromWord, sStyleInitMode, sStyleBaseUrl, sStyleUploadObject, sStyleAutoDir, sStyleBaseHref, sStyleContentPath, sStyleAutoRemote, sStyleShowBorder, sAutoDetectLanguage, sDefaultLanguage
Dim sSLTFlag, sSLTMinSize, sSLTOkSize, sSYFlag, sSYText, sSYFontColor, sSYFontSize, sSYFontName, sSYPicPath, sSLTSYObject, sSLTSYExt, sSYMinSize, sSYShadowColor, sSYShadowOffset
Dim sStyleFileExt, sStyleFlashExt, sStyleImageExt, sStyleMediaExt, sStyleRemoteExt, sStyleFileSize, sStyleFlashSize, sStyleImageSize, sStyleMediaSize, sStyleRemoteSize
Dim sToolBarID, sToolBarName, sToolBarOrder, sToolBarButton

Dim nStyleID

sPosition = sPosition & "��ʽ����"

If sAction = "STYLEPREVIEW" Then
	Call InitStyle()
	Call ShowStylePreview()
	Response.End
End If


Call Header()
Call ShowPosition()
Call Content()
Call Footer()


Sub Content()
	Select Case sAction
	Case "UPDATECONFIG"
		Call DoUpdateConfig()
	Case "COPY"
		Call InitStyle()
		Call DoCopy()
		Call ShowStyleList()
	Case "STYLEADD"
		Call ShowStyleForm("ADD")
	Case "STYLESET"
		Call InitStyle()
		Call ShowStyleForm("SET")
	Case "STYLEADDSAVE"
		Call CheckStyleForm()
		Call DoStyleAddSave()
	Case "STYLESETSAVE"
		Call CheckStyleForm()
		Call DoStyleSetSave()
	Case "STYLEDEL"
		Call InitStyle()
		Call DoStyleDel()
		Call ShowStyleList()
	Case "CODE"
		Call InitStyle()
		Call ShowStyleCode()
	Case "TOOLBAR"
		Call InitStyle()
		Call ShowToolBarList()
	Case "TOOLBARADD"
		Call InitStyle()
		Call DoToolBarAdd()
		Call ShowToolBarList()
	Case "TOOLBARMODI"
		Call InitStyle()
		Call DoToolBarModi()
		Call ShowToolBarList()
	Case "TOOLBARDEL"
		Call InitStyle()
		Call DoToolBarDel()
		Call ShowToolBarList()
	Case "BUTTONSET"
		Call InitStyle()
		Call InitToolBar()
		Call ShowButtonList()
	Case "BUTTONSAVE"
		Call InitStyle()
		Call InitToolBar()
		Call DoButtonSave()
	Case Else
		Call ShowStyleList()
	End Select
End Sub


Sub ShowPosition()
	Response.Write "<table border=0 cellspacing=1 align=center class=navi>" & _
		"<tr><th>" & sPosition & "</th></tr>" & _
		"<tr><td align=center>[<a href='?'>������ʽ�б�</a>]&nbsp;&nbsp;&nbsp;&nbsp;[<a href='?action=styleadd'>�½�һ��ʽ</a>]&nbsp;&nbsp;&nbsp;&nbsp;[<a href='?action=updateconfig'>����������ʽ��ǰ̨�����ļ�</a>]&nbsp;&nbsp;&nbsp;&nbsp;[<a href='#' onclick='history.back()'>����ǰһҳ</a>]</td></tr>" & _
		"</table><br>"
End Sub

Sub ShowMessage(str)
	Response.Write "<table border=0 cellspacing=1 align=center class=list><tr><td>" & str & "</td></tr></table><br>"
End Sub

Sub ShowStyleList()
	Call ShowMessage("<b class=blue>����Ϊ��ǰ������ʽ�б�</b>")

	Response.Write "<table border=0 cellpadding=0 cellspacing=1 class=list align=center>" & _
		"<form action='?action=del' method=post name=myform>" & _
		"<tr align=center>" & _
			"<th width='10%'>��ʽ��</th>" & _
			"<th width='10%'>��ѿ��</th>" & _
			"<th width='10%'>��Ѹ߶�</th>" & _
			"<th width='45%'>˵��</th>" & _
			"<th width='25%'>����</th>" & _
		"</tr>"

	Dim sManage, i, aCurrStyle
	For i = 1 To Ubound(aStyle)
		aCurrStyle = Split(aStyle(i), "|||")
		sManage = "<a href='?action=stylepreview&id=" & i & "' target='_blank'>Ԥ��</a>|<a href='?action=code&id=" & i & "'>����</a>|<a href='?action=styleset&id=" & i & "'>����</a>|<a href='?action=toolbar&id=" & i & "'>������</a>|<a href='?action=copy&id=" & i & "'>����</a>|<a href='?action=styledel&id=" & i & "' onclick=""return confirm('��ʾ����ȷ��Ҫɾ������ʽ��')"">ɾ��</a>"
		Response.Write "<tr align=center>" & _
			"<td>" & outHTML(aCurrStyle(0)) & "</td>" & _
			"<td>" & aCurrStyle(4) & "</td>" & _
			"<td>" & aCurrStyle(5) & "</td>" & _
			"<td align=left>" & outHTML(aCurrStyle(26)) & "</td>" & _
			"<td>" & sManage & "</td>" & _
			"</tr>"
	Next
	
	Response.Write "</table><br>"

	Call ShowMessage("<b class=blue>��ʾ��</b>�����ͨ����������һ��ʽ�Դﵽ�����½���ʽ��Ŀ�ġ�")

End Sub

Sub DoCopy()
	Dim i, b, sNewID, sNewName
	b = False
	i = 0
	Do While b = False
		i = i + 1
		sNewName = sStyleName & i
		If StyleName2ID(sNewName) = -1 Then
			b = True
		End If
	Loop

	Dim nNewStyleID
	nNewStyleID = Ubound(aStyle) + 1
	Redim Preserve aStyle(nNewStyleID)
	aStyle(nNewStyleID) = sNewName & Mid(aStyle(nStyleID), Len(sStyleName)+1)

	Dim nToolbarNum, nNewToolbarID, aCurrToolbar
	nToolbarNum = Ubound(aToolbar)
	For i = 1 To nToolbarNum
		aCurrToolbar = Split(aToolbar(i), "|||")
		If aCurrToolbar(0) = sStyleID Then
			nNewToolbarID = Ubound(aToolbar) + 1
			Redim Preserve aToolbar(nNewToolbarID)
			aToolbar(nNewToolbarID) = nNewStyleID & "|||" & aCurrToolbar(1) & "|||" & aCurrToolbar(2) & "|||" & aCurrToolbar(3)
		End If
	Next

	Call WriteConfig()
	Call WriteStyle(nNewStyleID)
	Call GoUrl("?")

End Sub

Function StyleName2ID(str)
	Dim i
	StyleName2ID = -1
	For i = 1 To UBound(aStyle)
		If Lcase(Split(aStyle(i), "|||")(0)) = Lcase(str) Then
			StyleName2ID = i
			Exit Function
		End If
	Next
End Function

Sub ShowStyleForm(sFlag)
	Dim s_Title, s_Button, s_Action
	Dim s_FormStateFlag, s_FormDetectFromWord, s_FormInitMode, s_FormBaseUrl, s_FormUploadObject, s_FormAutoDir, s_FormAutoRemote, s_FormShowBorder, s_FormAutoDetectLanguage, s_FormDefaultLanguage, s_FormSLTFlag, s_FormSYFlag, s_FormSLTSYObject
	
	If sFlag = "ADD" Then
		sStyleID = ""
		sStyleName = ""
		sStyleDir = "standard"
		sStyleCSS = "office"
		sStyleUploadDir = "UploadFile/"
		sStyleBaseHref = "http://Localhost/eWebEditor/"
		sStyleContentPath = "UploadFile/"
		sStyleWidth = "600"
		sStyleHeight = "400"
		sStyleMemo = ""
		nStyleIsSys = 0
		s_Title = "������ʽ"
		s_Action = "StyleAddSave"
		sStyleFileExt = "rar|zip|exe|doc|xls|chm|hlp"
		sStyleFlashExt = "swf"
		sStyleImageExt = "gif|jpg|jpeg|bmp"
		sStyleMediaExt = "rm|mp3|wav|mid|midi|ra|avi|mpg|mpeg|asf|asx|wma|mov"
		sStyleRemoteExt = "gif|jpg|bmp"
		sStyleFileSize = "500"
		sStyleFlashSize = "100"
		sStyleImageSize = "100"
		sStyleMediaSize = "100"
		sStyleRemoteSize = "100"
		sStyleStateFlag = "1"
		sStyleAutoRemote = "1"
		sStyleShowBorder = "0"
		sAutoDetectLanguage = "1"
		sDefaultLanguage = "zh-cn"
		sStyleUploadObject = "0"
		sStyleAutoDir = "0"
		sStyleDetectFromWord = "1"
		sStyleInitMode = "EDIT"
		sStyleBaseUrl = "0"
		sSLTFlag = "0"
		sSLTMinSize = "300"
		sSLTOkSize = "120"
		sSYFlag = "0"
		sSYText = "��Ȩ����..."
		sSYFontColor = "000000"
		sSYFontSize = "12"
		sSYFontName = "����"
		sSYPicPath = ""
		sSLTSYObject = "0"
		sSLTSYExt = "bmp|jpg|jpeg|gif"
		sSYMinSize = "100"
		sSYShadowColor = "FFFFFF"
		sSYShadowOffset = "1"
	Else
		sStyleName = inHTML(sStyleName)
		sStyleDir = inHTML(sStyleDir)
		sStyleCSS = inHTML(sStyleCSS)
		sStyleUploadDir = inHTML(sStyleUploadDir)
		sStyleBaseHref = inHTML(sStyleBaseHref)
		sStyleContentPath = inHTML(sStyleContentPath)
		sStyleMemo = inHTML(sStyleMemo)
		sSYText = inHTML(sSYText)
		sSYFontColor = inHTML(sSYFontColor)
		sSYFontSize = inHTML(sSYFontSize)
		sSYFontName = inHTML(sSYFontName)
		sSYPicPath = inHTML(sSYPicPath)
		s_Title = "������ʽ"
		s_Action = "StyleSetSave"
	End If

	s_FormStateFlag = InitSelect("d_stateflag", Split("��ʾ|����ʾ", "|"), Split("1|0", "|"), sStyleStateFlag, "", "")
	s_FormAutoRemote = InitSelect("d_autoremote", Split("�Զ��ϴ�|���Զ��ϴ�", "|"), Split("1|0", "|"), sStyleAutoRemote, "", "")
	s_FormShowBorder = InitSelect("d_showborder", Split("Ĭ����ʾ|Ĭ�ϲ���ʾ", "|"), Split("1|0", "|"), sStyleShowBorder, "", "")
	s_FormAutoDetectLanguage = InitSelect("d_autodetectlanguage", Split("�Զ����|���Զ����", "|"), Split("1|0", "|"), sAutoDetectLanguage, "", "")
	s_FormDefaultLanguage = InitSelect("d_defaultlanguage", Split("��������|��������|Ӣ��", "|"), Split("zh-cn|zh-tw|en", "|"), sDefaultLanguage, "", "")
	s_FormUploadObject = InitSelect("d_uploadobject", Split("������ϴ���|ASPUpload�ϴ����|SA-FileUp�ϴ����|LyfUpload�ϴ����", "|"), Split("0|1|2|3", "|"), sStyleUploadObject, "", "")
	s_FormAutoDir = InitSelect("d_autodir", Split("��ʹ��|��Ŀ¼|����Ŀ¼|������Ŀ¼", "|"), Split("0|1|2|3", "|"), sStyleAutoDir, "", "")
	s_FormDetectFromWord = InitSelect("d_detectfromword", Split("�Զ��������ʾ|���Զ����", "|"), Split("1|0", "|"), sStyleDetectFromWord, "", "")
	s_FormInitMode = InitSelect("d_initmode", Split("����ģʽ|�༭ģʽ|�ı�ģʽ|Ԥ��ģʽ", "|"), Split("CODE|EDIT|TEXT|VIEW", "|"), sStyleInitMode, "", "")
	s_FormBaseUrl = InitSelect("d_baseurl", Split("���·��|���Ը�·��|����ȫ·��", "|"), Split("0|1|2", "|"), sStyleBaseUrl, "", "")

	s_FormSLTFlag = InitSelect("d_sltflag", Split("ʹ��|��ʹ��", "|"), Split("1|0", "|"), sSLTFlag, "", "")
	s_FormSYFlag = InitSelect("d_syflag", Split("��ʹ��|����ˮӡ|ͼƬˮӡ", "|"), Split("0|1|2", "|"), sSYFlag, "", "")
	s_FormSLTSYObject = InitSelect("d_sltsyobject", Split("AspJpegͼ�����", "|"), Split("0", "|"), sSLTSYObject, "", "")

	s_Button = "<tr><td align=center colspan=4><input type=submit value='  �ύ  ' align=absmiddle>&nbsp;<input type=reset name=btnReset value='  ����  '></td></tr>"

	Response.Write "<table border=0 cellpadding=0 cellspacing=1 align=center class=form>" & _
			"<form action='?action=" & s_Action & "&id=" & sStyleID & "' method=post name=myform>" & _
			"<tr><th colspan=4>&nbsp;&nbsp;" & s_Title & "������Ƶ������ɿ�˵������*��Ϊ�����</th></tr>" & _
			"<tr><td width='15%'>��ʽ���ƣ�</td><td width='35%'><input type=text class=input size=20 name=d_name title='���ô���ʽ�����֣���Ҫ��������ţ����50���ַ�����' value=""" & sStyleName & """> <span class=red>*</span></td><td width='15%'>��ʼģʽ��</td><td width='35%'>" & s_FormInitMode & " <span class=red>*</span></td></tr>" & _
			"<tr><td width='15%'>�ϴ������</td><td width='35%'>" & s_FormUploadObject & " <span class=red>*</span></td><td width='15%'>�Զ�Ŀ¼��</td><td width='35%'>" & s_FormAutoDir & " <span class=red>*</span></td></tr>" & _
			"<tr><td width='15%'>ͼƬĿ¼��</td><td width='35%'><input type=text class=input size=20 name=d_dir title='��Ŵ���ʽͼƬ�ļ���Ŀ¼����������ButtonImage�£����50���ַ�����' value=""" & sStyleDir & """> <span class=red>*</span></td><td width='15%'>��ʽĿ¼��</td><td width='35%'><input type=text class=input size=20 name=d_css title='��Ŵ���ʽcss�ļ���Ŀ¼����������CSS�£����50���ַ�����' value=""" & sStyleCSS & """> <span class=red>*</span></td></tr>" & _
			"<tr><td width='15%'>��ѿ�ȣ�</td><td width='35%'><input type=text class=input name=d_width size=20 title='�������Ч���Ŀ�ȣ�������' value='" & sStyleWidth & "'> <span class=red>*</span></td><td width='15%'>��Ѹ߶ȣ�</td><td width='35%'><input type=text class=input name=d_height size=20 title='�������Ч���ĸ߶ȣ�������' value='" & sStyleHeight & "'> <span class=red>*</span></td></tr>" & _
			"<tr><td width='15%'>״ ̬ ����</td><td width='35%'>" & s_FormStateFlag & " <span class=red>*</span></td><td width='15%'>Wordճ����</td><td width='35%'>" & s_FormDetectFromWord & " <span class=red>*</span></td></tr>" & _
			"<tr><td width='15%'>Զ���ļ���</td><td width='35%'>" & s_FormAutoRemote & " <span class=red>*</span></td><td width='15%'>ָ�����룺</td><td width='35%'>" & s_FormShowBorder & " <span class=red>*</span></td></tr>" & _
			"<tr><td width='15%'>�Զ����Լ�⣺</td><td width='35%'>" & s_FormAutoDetectLanguage & " <span class=red>*</span></td><td width='15%'>Ĭ�����ԣ�</td><td width='35%'>" & s_FormDefaultLanguage & " <span class=red>*</span></td></tr>" & _
			"<tr><td>��ע˵����</td><td colspan=3><input type=text name=d_memo size=90 title='����ʽ��˵�����������ڵ���' value=""" & sStyleMemo & """></td></tr>" & _
			"<tr><td colspan=4><span class=red>&nbsp;&nbsp;&nbsp;�ϴ��ļ���ϵͳ�ļ�·��������ã�ֻ����ʹ�����·��ģʽʱ����Ҫ������ʾ·��������·������</span></td></tr>" & _
			"<tr><td width='15%'>·��ģʽ��</td><td width='35%'>" & s_FormBaseUrl & " <span class=red>*</span> <a href='#baseurl'>˵��</a></td><td width='15%'>�ϴ�·����</td><td width='35%'><input type=text class=input size=20 name=d_uploaddir title='�ϴ��ļ������·�������eWebEditor��Ŀ¼�ļ���·�������50���ַ�����' value=""" & sStyleUploadDir & """> <span class=red>*</span></td></tr>" & _
			"<tr><td width='15%'>��ʾ·����</td><td width='35%'><input type=text class=input size=20 name=d_basehref title='��ʾ����ҳ�����·����������&quot;/&quot;��ͷ�����50���ַ�����' value=""" & sStyleBaseHref & """></td><td width='15%'>����·����</td><td width='35%'><input type=text class=input size=20 name=d_contentpath title='ʵ�ʱ����������е�·���������ʾ·����·����������&quot;/&quot;��ͷ�����50���ַ�����' value=""" & sStyleContentPath & """></td></tr>" & _
			"<tr><td colspan=4><span class=red>&nbsp;&nbsp;&nbsp;�����ϴ��ļ����ͼ��ļ���С���ã��ļ���С��λΪKB��0��ʾû�����ƣ���</span></td></tr>" & _
			"<tr><td width='15%'>ͼƬ���ͣ�</td><td width='35%'><input type=text class=input name=d_imageext size=20 title='����ͼƬ��ص��ϴ������250���ַ�����' value='" & sStyleImageExt & "'></td><td width='15%'>ͼƬ���ƣ�</td><td width='35%'><input type=text class=input name=d_imagesize size=20 title='�����ͣ���λKB' value='" & sStyleImageSize & "'>KB</td></tr>" & _
			"<tr><td width='15%'>Flash���ͣ�</td><td width='35%'><input type=text class=input name=d_flashext size=20 title='���ڲ���Flash���������250���ַ�����' value='" & sStyleFlashExt & "'></td><td width='15%'>Flash���ƣ�</td><td width='35%'><input type=text class=input name=d_flashsize size=20 title='�����ͣ���λKB' value='" & sStyleFlashSize & "'>KB</td></tr>" & _
			"<tr><td width='15%'>ý�����ͣ�</td><td width='35%'><input type=text class=input name=d_mediaext size=20 title='���ڲ���ý���ļ������250���ַ�����' value='" & sStyleMediaExt & "'></td><td width='15%'>ý�����ƣ�</td><td width='35%'><input type=text class=input name=d_mediasize size=20 title='�����ͣ���λKB' value='" & sStyleMediaSize & "'>KB</td></tr>" & _
			"<tr><td width='15%'>�������ͣ�</td><td width='35%'><input type=text class=input name=d_fileext size=20 title='���ڲ��������ļ������250���ַ�����' value='" & sStyleFileExt & "'></td><td width='15%'>�������ƣ�</td><td width='35%'><input type=text class=input name=d_filesize size=20 title='�����ͣ���λKB' value='" & sStyleFileSize & "'>KB</td></tr>" & _
			"<tr><td width='15%'>Զ�����ͣ�</td><td width='35%'><input type=text class=input name=d_remoteext size=20 title='�����Զ��ϴ�Զ���ļ������250���ַ�����' value='" & sStyleRemoteExt & "'></td><td width='15%'>Զ�����ƣ�</td><td width='35%'><input type=text class=input name=d_remotesize size=20 title='�����ͣ���λKB' value='" & sStyleRemoteSize & "'>KB</td></tr>" & _
			"<tr><td colspan=4><span class=red>&nbsp;&nbsp;&nbsp;����ͼ��ˮӡ������ã�</span></td></tr>" & _
			"<tr><td width='15%'>ͼ�δ��������</td><td width='35%'>" & s_FormSLTSYObject & "</td><td width='15%'>����ͼ����չ����</td><td width='35%'><input type=text name=d_sltsyext size=20 class=input value=""" & sSLTSYExt & """></td></tr>" & _
			"<tr><td width='15%'>����ͼʹ��״̬��</td><td width='35%'>" & s_FormSLTFlag & "</td><td width='15%'>����ͼ��������</td><td width='35%'><input type=text name=d_sltminsize size=20 class=input title='ͼ�εĳ���ֻ�дﵽ����С����Ҫ��ʱ�Ż���������ͼ��������' value='" & sSLTMinSize & "'>px</td></tr>" & _
			"<tr><td width='15%'>����ͼ���ɳ��ȣ�</td><td width='35%'><input type=text name=d_sltoksize size=20 class=input title='���ɵ�����ͼ����ֵ��������' value='" & sSLTOkSize & "'>px</td><td width='15%'>&nbsp;</td><td width='35%'>&nbsp;</td></tr>" & _
			"<tr><td width='15%'>ˮӡʹ��״̬��</td><td width='35%'>" & s_FormSYFlag & "</td><td width='15%'>ˮӡ���������</td><td width='35%'><input type=text name=d_syminsize size=20 class=input title='ͼ�εĿ��ֻ�дﵽ����С���Ҫ��ʱ�Ż�����ˮӡ��������' value='" & sSYMinSize & "'>px</td></tr>" & _
			"<tr><td width='15%'>����ˮӡ���ݣ�</td><td width='35%'><input type=text name=d_sytext size=20 class=input title='��ʹ������ˮӡʱ����������' value=""" & sSYText & """></td><td width='15%'>����ˮӡ������ɫ��</td><td width='35%'><input type=text name=d_syfontcolor size=20 class=input title='��ʹ������ˮӡʱ���ֵ���ɫ' value=""" & sSYFontColor & """></td></tr>" & _
			"<tr><td width='15%'>����ˮӡ��Ӱ��ɫ��</td><td width='35%'><input type=text name=d_syshadowcolor size=20 class=input title='��ʹ������ˮӡʱ��������Ӱ��ɫ' value=""" & sSYShadowColor & """></td><td width='15%'>����ˮӡ��Ӱ��С��</td><td width='35%'><input type=text name=d_syshadowoffset size=20 class=input title='��ʹ������ˮӡʱ���ֵ���Ӱ��С' value=""" & sSYShadowOffset & """>px</td></tr>" & _
			"<tr><td width='15%'>����ˮӡ�����С��</td><td width='35%'><input type=text name=d_syfontsize size=20 class=input title='��ʹ������ˮӡʱ���ֵ������С' value=""" & sSYFontSize & """>px</td><td width='15%'>����ˮӡ�������ƣ�</td><td width='35%'><input type=text name=d_syfontname size=20 class=input title='��ʹ������ˮӡʱ���ֵ�������' value=""" & sSYFontName & """></td></tr>" & _
			"<tr><td width='15%'>ͼƬˮӡͼƬ·����</td><td width='35%'><input type=text name=d_sypicpath size=20 class=input title='��ʹ��ͼƬˮӡʱͼƬ��·��' value=""" & sSYPicPath & """></td><td width='15%'></td><td width='35%'></td></tr>" & _
			"" & s_Button & _
			"</form>" & _
			"</table><br>"

	Dim sMsg
	sMsg = "<a name=baseurl></a><p><span class=blue><b>·��ģʽ����˵����</b></span><br>" & _
		"<b>���·����</b>ָ���е�����ϴ����Զ������ļ�·�����༭����""UploadFile/...""��""../UploadFile/...""��ʽ���֣���ʹ�ô�ģʽʱ����ʾ·��������·�������ʾ·��������""/""��ͷ�ͽ�β������·�������в�����""/""��ͷ��<br>" & _
		"<b>���Ը�·����</b>ָ���е�����ϴ����Զ������ļ�·�����༭����""/eWebEditor/UploadFile/...""������ʽ���֣���ʹ�ô�ģʽʱ����ʾ·��������·�������<br>" & _
		"<b>����ȫ·����</b>ָ���е�����ϴ����Զ������ļ�·�����༭����""http://xxx.xxx.xxx/eWebEditor/UploadFile/...""������ʽ���֣���ʹ�ô�ģʽʱ����ʾ·��������·�������</p>"

	Call ShowMessage(sMsg)

End Sub

Sub InitStyle()
	Dim b, aCurrStyle
	b = False
	sStyleID = Trim(Request("id"))
	If IsNumeric(sStyleID) = True Then
		nStyleID = Clng(sStyleID)
		If nStyleID <= Ubound(aStyle) Then
			aCurrStyle = Split(aStyle(nStyleID), "|||")
			sStyleName = aCurrStyle(0)
			sStyleDir = aCurrStyle(1)
			sStyleCSS = aCurrStyle(2)
			sStyleUploadDir = aCurrStyle(3)
			sStyleBaseHref = aCurrStyle(22)
			sStyleContentPath = aCurrStyle(23)
			sStyleWidth = aCurrStyle(4)
			sStyleHeight = aCurrStyle(5)
			sStyleMemo = aCurrStyle(26)
			sStyleFileExt = aCurrStyle(6)
			sStyleFlashExt = aCurrStyle(7)
			sStyleImageExt = aCurrStyle(8)
			sStyleMediaExt = aCurrStyle(9)
			sStyleRemoteExt = aCurrStyle(10)
			sStyleFileSize = aCurrStyle(11)
			sStyleFlashSize = aCurrStyle(12)
			sStyleImageSize = aCurrStyle(13)
			sStyleMediaSize = aCurrStyle(14)
			sStyleRemoteSize = aCurrStyle(15)
			sStyleStateFlag = aCurrStyle(16)
			sStyleAutoRemote = aCurrStyle(24)
			sStyleShowBorder = aCurrStyle(25)
			sAutoDetectLanguage = aCurrStyle(27)
			sDefaultLanguage = aCurrStyle(28)
			sStyleUploadObject = aCurrStyle(20)
			sStyleAutoDir = aCurrStyle(21)
			sStyleDetectFromWord = aCurrStyle(17)
			sStyleInitMode = aCurrStyle(18)
			sStyleBaseUrl = aCurrStyle(19)
			sSLTFlag = aCurrStyle(29)
			sSLTMinSize = aCurrStyle(30)
			sSLTOkSize = aCurrStyle(31)
			sSYFlag = aCurrStyle(32)
			sSYText = aCurrStyle(33)
			sSYFontColor = aCurrStyle(34)
			sSYFontSize = aCurrStyle(35)
			sSYFontName = aCurrStyle(36)
			sSYPicPath = aCurrStyle(37)
			sSLTSYObject = aCurrStyle(38)
			sSLTSYExt = aCurrStyle(39)
			sSYMinSize = aCurrStyle(40)
			sSYShadowColor = aCurrStyle(41)
			sSYShadowOffset = aCurrStyle(42)
			b = True
		End If
	End If
	If b = False Then
		GoError "��Ч����ʽID�ţ���ͨ��ҳ���ϵ����ӽ��в�����"
	End If
End Sub

Sub CheckStyleForm()
	sStyleName = Trim(Request("d_name"))
	sStyleDir = Trim(Request("d_dir"))
	sStyleCSS = Trim(Request("d_css"))
	sStyleUploadDir = Trim(Request("d_uploaddir"))
	sStyleBaseHref = Trim(Request("d_basehref"))
	sStyleContentPath = Trim(Request("d_contentpath"))
	sStyleWidth = Trim(Request("d_width"))
	sStyleHeight = Trim(Request("d_height"))
	sStyleMemo = Trim(Request("d_memo"))
	sStyleImageExt = Trim(Request("d_imageext"))
	sStyleFlashExt = Trim(Request("d_flashext"))
	sStyleMediaExt = Trim(Request("d_mediaext"))
	sStyleRemoteExt = Trim(Request("d_remoteext"))
	sStyleFileExt = Trim(Request("d_fileext"))
	sStyleImageSize = Trim(Request("d_imagesize"))
	sStyleFlashSize = Trim(Request("d_flashsize"))
	sStyleMediaSize = Trim(Request("d_mediasize"))
	sStyleRemoteSize = Trim(Request("d_remotesize"))
	sStyleFileSize = Trim(Request("d_filesize"))
	sStyleStateFlag = Trim(Request("d_stateflag"))
	sStyleAutoRemote = Trim(Request("d_autoremote"))
	sStyleShowBorder = Trim(Request("d_showborder"))
	sAutoDetectLanguage = Trim(Request("d_autodetectlanguage"))
	sDefaultLanguage = Trim(Request("d_defaultlanguage"))
	sStyleUploadObject = Trim(Request("d_uploadobject"))
	sStyleAutoDir = Trim(Request("d_autodir"))
	sStyleDetectFromWord = Trim(Request("d_detectfromword"))
	sStyleInitMode = Trim(Request("d_initmode"))
	sStyleBaseUrl = Trim(Request("d_baseurl"))
	sSLTFlag = Trim(Request("d_sltflag"))
	sSLTMinSize = Trim(Request("d_sltminsize"))
	sSLTOkSize = Trim(Request("d_sltoksize"))
	sSYFlag = Trim(Request("d_syflag"))
	sSYText = Trim(Request("d_sytext"))
	sSYFontColor = Trim(Request("d_syfontcolor"))
	sSYFontSize = Trim(Request("d_syfontsize"))
	sSYFontName = Trim(Request("d_syfontname"))
	sSYPicPath = Trim(Request("d_sypicpath"))
	sSLTSYObject = Trim(Request("d_sltsyobject"))
	sSLTSYExt = Trim(Request("d_sltsyext"))
	sSYMinSize = Trim(Request("d_syminsize"))
	sSYShadowColor = Trim(Request("d_syshadowcolor"))
	sSYShadowOffset = Trim(Request("d_syshadowoffset"))

	sStyleUploadDir = Replace(sStyleUploadDir, "\", "/")
	sStyleBaseHref = Replace(sStyleBaseHref, "\", "/")
	sStyleContentPath = Replace(sStyleContentPath, "\", "/")
	If Right(sStyleUploadDir, 1) <> "/" Then sStyleUploadDir = sStyleUploadDir & "/"
	If Right(sStyleBaseHref, 1) <> "/" Then sStyleBaseHref = sStyleBaseHref & "/"
	If Right(sStyleContentPath, 1) <> "/" Then sStyleContentPath = sStyleContentPath & "/"

	If sStyleName = "" Then
		GoError "��ʽ������Ϊ�գ�"
	End If
	If IsSafeStr(sStyleName) = False Then
		GoError "��ʽ��������������ַ���"
	End If
	If sStyleDir = "" Then
		GoError "��ťͼƬĿ¼������Ϊ�գ�"
	End If
	If IsSafeStr(sStyleDir) = False Then
		GoError "��ťͼƬĿ¼��������������ַ���"
	End If
	If sStyleCSS = "" Then
		GoError "��ʽCSSĿ¼������Ϊ�գ�"
	End If
	If IsSafeStr(sStyleCSS) = False Then
		GoError "��ʽCSSĿ¼��������������ַ���"
	End If

	If sStyleUploadDir = "" Then
		GoError "�ϴ�·������Ϊ�գ�"
	End If
	If IsSafeStr(sStyleUploadDir) = False Then
		GoError "�ϴ�·��������������ַ���"
	End If
	Select Case sStyleBaseUrl
	Case "0"
		If sStyleBaseHref = "" Then
			GoError "��ʹ�����·��ģʽʱ����ʾ·������Ϊ�գ�"
		End If
		If IsSafeStr(sStyleBaseHref) = False Then
			GoError "��ʹ�����·��ģʽʱ����ʾ·��������������ַ���"
		End If
		If Left(sStyleBaseHref, 1) <> "/" Then
			GoError "��ʹ�����·��ģʽʱ����ʾ·��������&quot;/&quot;��ͷ��"
		End If

		If sStyleContentPath = "" Then
			GoError "��ʹ�����·��ģʽʱ������·������Ϊ�գ�"
		End If
		If IsSafeStr(sStyleContentPath) = False Then
			GoError "��ʹ�����·��ģʽʱ������·��������������ַ���"
		End If
		If Left(sStyleContentPath, 1) = "/" Then
			GoError "��ʹ�����·��ģʽʱ������·��������&quot;/&quot;��ͷ��"
		End If
	Case "1", "2"
		sStyleBaseHref = ""
		sStyleContentPath = ""
	End Select
	
	If IsNumeric(sStyleWidth) = False Then
		GoError "����д��Ч��������ÿ�ȣ�"
	End If
	If IsNumeric(sStyleHeight) = False Then
		GoError "����д��Ч��������ø߶ȣ�"
	End If

	If IsNumeric(sStyleImageSize) = False Then
		GoError "����д��Ч��ͼƬ���ƴ�С��"
	End If
	If IsNumeric(sStyleFlashSize) = False Then
		GoError "����д��Ч��Flash���ƴ�С��"
	End If
	If IsNumeric(sStyleMediaSize) = False Then
		GoError "����д��Ч��ý���ļ����ƴ�С��"
	End If
	If IsNumeric(sStyleFileSize) = False Then
		GoError "����д��Ч�������ļ����ƴ�С��"
	End If
	If IsNumeric(sStyleRemoteSize) = False Then
		GoError "����д��Ч��Զ���ļ����ƴ�С��"
	End If

	If IsNumeric(sSLTMinSize) = False Then
		GoError "����д��Ч������ͼʹ����С��������������Ϊ�գ���Ϊ�����ͣ�"
	End If
	If IsNumeric(sSLTOkSize) = False Then
		GoError "����д��Ч������ͼ���ɳ��ȣ�����Ϊ�գ���Ϊ�����ͣ�"
	End If

	If IsNumeric(sSYMinSize) = False Then
		GoError "����д��Ч��ˮӡ����С�������������Ϊ�գ���Ϊ�����ͣ�"
	End If
	If sSYText = "" Then
		GoError "����д��Чˮӡ�������ݣ�����Ϊ�գ�"
	End If
	If isValidColor(sSYFontColor) = False Then
		GoError "����д��Ч��ˮӡ������ɫ��6λ���ȣ����ɫ��000000��"
	End If
	If isValidColor(sSYShadowColor) = False Then
		GoError "����д��Ч��ˮӡ������Ӱ��ɫ��6λ���ȣ����ɫ��FFFFFF��"
	End If
	If IsNumeric(sSYShadowOffset) = False Then
		GoError "����д��Ч��ˮӡ������Ӱ��С������Ϊ�գ���Ϊ�����ͣ�"
	End If
	If IsNumeric(sSYFontSize) = False Then
		GoError "����д��Ч��ˮӡ���ִ�С������Ϊ�գ���Ϊ�����ͣ�"
	End If
	If sSYFontName = "" Then
		GoError "����дˮӡ�����������ƣ�����Ϊ�գ�"
	End If

End Sub

Sub DoStyleAddSave()

	If StyleName2ID(sStyleName) <> -1 Then
		GoError "����ʽ���Ѿ����ڣ�������һ����ʽ����"
	End If

	Dim nNewStyleID
	nNewStyleID = Ubound(aStyle) + 1
	Redim Preserve aStyle(nNewStyleID)

	aStyle(nNewStyleID) = sStyleName & "|||" & sStyleDir & "|||" & sStyleCSS & "|||" & sStyleUploadDir & "|||" & sStyleWidth & "|||" & sStyleHeight & "|||" & sStyleFileExt & "|||" & sStyleFlashExt & "|||" & sStyleImageExt & "|||" & sStyleMediaExt & "|||" & sStyleRemoteExt & "|||" & sStyleFileSize & "|||" & sStyleFlashSize & "|||" & sStyleImageSize & "|||" & sStyleMediaSize & "|||" & sStyleRemoteSize & "|||" & sStyleStateFlag & "|||" & sStyleDetectFromWord & "|||" & sStyleInitMode & "|||" & sStyleBaseUrl & "|||" & sStyleUploadObject & "|||" & sStyleAutoDir & "|||" & sStyleBaseHref & "|||" & sStyleContentPath & "|||" & sStyleAutoRemote & "|||" & sStyleShowBorder & "|||" & sStyleMemo & "|||" & sAutoDetectLanguage & "|||" & sDefaultLanguage & "|||" & sSLTFlag & "|||" & sSLTMinSize & "|||" & sSLTOkSize & "|||" & sSYFlag & "|||" & sSYText & "|||" & sSYFontColor & "|||" & sSYFontSize & "|||" & sSYFontName & "|||" & sSYPicPath & "|||" & sSLTSYObject & "|||" & sSLTSYExt & "|||" & sSYMinSize & "|||" & sSYShadowColor & "|||" & sSYShadowOffset

	Call WriteConfig()
	Call WriteStyle(nNewStyleID)

	Call ShowMessage("<b><span class=red>��ʽ���ӳɹ���</span></b><li><a href='?action=toolbar&id=" & nNewStyleID & "'>���ô���ʽ�µĹ�����</a>")

End Sub

Sub DoUpdateConfig()
	Dim i
	Call WriteConfig()
	For i = 1 To UBound(aStyle)
		Call WriteStyle(i)
	Next
	Call ShowMessage("<b><span class=red>������ʽ��ǰ̨�����ļ����²����ɹ���</span></b><li><a href='?'>����������ʽ�б�</a>")
End Sub

Sub DoStyleSetSave()
	Dim n, s_OldStyleName
	sStyleID = Trim(Request("id"))
	If IsNumeric(sStyleID) = True Then
		n = StyleName2ID(sStyleName)
		If CStr(n) <> sStyleID And n <> -1 Then
			GoError "����ʽ���Ѿ����ڣ�������һ����ʽ����"
		End If
		
		If Clng(sStyleID) < 1 And Clng(sStyleID)>UBound(aStyle) Then
			GoError "��Ч����ʽID�ţ���ͨ��ҳ���ϵ����ӽ��в�����"
		End If

		s_OldStyleName = Split(aStyle(Clng(sStyleID)), "|||")(0)

		aStyle(Clng(sStyleID)) = sStyleName & "|||" & sStyleDir & "|||" & sStyleCSS & "|||" & sStyleUploadDir & "|||" & sStyleWidth & "|||" & sStyleHeight & "|||" & sStyleFileExt & "|||" & sStyleFlashExt & "|||" & sStyleImageExt & "|||" & sStyleMediaExt & "|||" & sStyleRemoteExt & "|||" & sStyleFileSize & "|||" & sStyleFlashSize & "|||" & sStyleImageSize & "|||" & sStyleMediaSize & "|||" & sStyleRemoteSize & "|||" & sStyleStateFlag & "|||" & sStyleDetectFromWord & "|||" & sStyleInitMode & "|||" & sStyleBaseUrl & "|||" & sStyleUploadObject & "|||" & sStyleAutoDir & "|||" & sStyleBaseHref & "|||" & sStyleContentPath & "|||" & sStyleAutoRemote & "|||" & sStyleShowBorder & "|||" & sStyleMemo & "|||" & sAutoDetectLanguage & "|||" & sDefaultLanguage & "|||" & sSLTFlag & "|||" & sSLTMinSize & "|||" & sSLTOkSize & "|||" & sSYFlag & "|||" & sSYText & "|||" & sSYFontColor & "|||" & sSYFontSize & "|||" & sSYFontName & "|||" & sSYPicPath & "|||" & sSLTSYObject & "|||" & sSLTSYExt & "|||" & sSYMinSize & "|||" & sSYShadowColor & "|||" & sSYShadowOffset

	Else
		GoError "��Ч����ʽID�ţ���ͨ��ҳ���ϵ����ӽ��в�����"
	End If

	Call WriteConfig()
	If LCase(s_OldStyleName) <> LCase(sStyleName) Then
		Call DeleteFile(s_OldStyleName)
	End If
	Call WriteStyle(Clng(sStyleID))

	Call ShowMessage("<b><span class=red>��ʽ�޸ĳɹ���</span></b><li><a href='?action=stylepreview&id=" & sStyleID & "' target='_blank'>Ԥ������ʽ</a><li><a href='?action=toolbar&id=" & sStyleID & "'>���ô���ʽ�µĹ�����</a>")

End Sub

Sub DoStyleDel()
	aStyle(Clng(sStyleID)) = ""
	Call WriteConfig()
	Call DeleteFile(sStyleName)
	Call GoUrl("?")
End Sub

Sub ShowStylePreview()
	Response.Write "<html><head>" & _
		"<title>��ʽԤ��</title>" & _
		"<meta http-equiv='Content-Type' content='text/html; charset=gb2312'>" & _
		"</head><body>" & _
		"<input type=hidden name=content1  value=''>" & _
		"<iframe ID='eWebEditor1' src='../ewebeditor.htm?id=content1&style=" & sStyleName & "' frameborder=0 scrolling=no width='" & sStyleWidth & "' HEIGHT='" & sStyleHeight & "'></iframe>" & _
		"</body></html>"
End Sub

Sub ShowStyleCode()
	Response.Write "<table border=0 cellspacing=1 align=center class=list>" & _
		"<tr><th>��ʽ��" & outHTML(sStyleName) & "������ѵ��ô������£�����XXX��ʵ�ʹ����ı�������޸ģ���</th></tr>" & _
		"<tr><td><textarea rows=5 cols=65 style='width:100%'><IFRAME ID=""eWebEditor1"" SRC=""ewebeditor.htm?id=XXX&style=" & sStyleName & """ FRAMEBORDER=""0"" SCROLLING=""no"" WIDTH=""" & sStyleWidth & """ HEIGHT=""" & sStyleHeight & """></IFRAME></textarea></td></tr>" & _
		"</table>"
End Sub

Sub ShowToolBarList()

	Call ShowMessage("<b class=blue>��ʽ��" & outHTML(sStyleName) & "���µĹ���������</b>")

	Dim s_AddForm, s_ModiForm, i, aCurrToolbar

	If nStyleIsSys = 1 Then
		s_AddForm = ""
	Else
		Dim nMaxOrder
		nMaxOrder = 0
		For i = 1 To UBound(aToolbar)
			aCurrToolbar = Split(aToolbar(i), "|||")
			If aCurrToolbar(0) = sStyleID Then
				If Clng(aCurrToolbar(3)) > nMaxOrder Then
					nMaxOrder = Clng(aCurrToolbar(3))
				End If
			End If
		Next
		nMaxOrder = nMaxOrder + 1

		s_AddForm = "<hr width='80%' align=center size=1><table border=0 cellpadding=4 cellspacing=0 align=center>" & _
		"<form action='?id=" & sStyleID & "&action=toolbaradd' name='addform' method=post>" & _
		"<tr><td>����������<input type=text name=d_name size=20 class=input value='������" & nMaxOrder & "'> ����ţ�<input type=text name=d_order size=5 value='" & nMaxOrder & "' class=input> <input type=submit name=b1 value='����������'></td></tr>" & _
		"</form></table><hr width='80%' align=center size=1>"
	End If

	Dim s_Manage, s_SubmitButton
	s_ModiForm = "<form action='?id=" & sStyleID & "&action=toolbarmodi' name=modiform method=post>" & _
		"<table border=0 cellpadding=0 cellspacing=1 align=center class=form>" & _
		"<tr align=center><th>ID</th><th>��������</th><th>�����</th><th>����</th></tr>"

	For i = 1 To UBound(aToolbar)
		aCurrToolbar = Split(aToolbar(i), "|||")
		If aCurrToolbar(0) = sStyleID Then
			s_Manage = "<a href='?id=" & sStyleID & "&action=buttonset&toolbarid=" & i & "'>��ť����</a>"
			s_Manage = s_Manage & "|<a href='?id=" & sStyleID & "&action=toolbardel&delid=" & i & "'>ɾ��</a>"
			s_ModiForm = s_ModiForm & "<tr align=center>" & _
				"<td>" & i & "</td>" & _
				"<td><input type=text name='d_name" & i & "' value=""" & inHTML(aCurrToolbar(2)) & """ size=30 class=input></td>" & _
				"<td><input type=text name='d_order" & i & "' value='" & aCurrToolbar(3) & "' size=5 class=input></td>" & _
				"<td>" & s_Manage & "</td>" & _
				"</tr>"
		End If
	Next

	s_SubmitButton = "<tr><td colspan=4 align=center><input type=submit name=b1 value='  �޸�  '></td></tr>"
	s_ModiForm = s_ModiForm & s_SubmitButton & "</table></form>"

	Response.Write s_AddForm & s_ModiForm

End Sub

Sub DoToolBarAdd()
	Dim s_Name, s_Order
	s_Name = Trim(Request("d_name"))
	s_Order = Trim(Request("d_order"))
	If s_Name = "" Or GetTrueLen(s_Name) > 50 Then
		GoError "������������Ϊ�գ��ҳ��Ȳ��ܴ���50���ַ����ȣ�"
	End If
	If IsNumeric(s_Order) = False Then
		GoError "��Ч�Ĺ���������ţ�����ű���Ϊ���֣�"
	End If

	Dim nToolbarNum
	nToolbarNum = Ubound(aToolbar) + 1
	Redim Preserve aToolbar(nToolbarNum)
	aToolbar(nToolbarNum) = sStyleID & "||||||" & s_Name & "|||" & s_Order

	Call WriteConfig()
	Call WriteStyle(Clng(sStyleID))

	Response.Write "<script language=javascript>alert(""��������" & outHTML(s_Name) & "�����Ӳ����ɹ���"");</script>"
	Call GoUrl("?action=toolbar&id=" & sStyleID)
End Sub

Sub DoToolBarModi()
	Dim s_Name, s_Order, i, aCurrToolbar

	For i = 1 To UBound(aToolbar)
		aCurrToolbar = Split(aToolbar(i), "|||")
		If aCurrToolbar(0) = sStyleID Then
			s_Name = Trim(Request("d_name" & i))
			s_Order = Trim(Request("d_order" & i))
			If s_Name = "" Or IsNumeric(s_Order) = False Then 
				aCurrToolbar(0) = ""
				s_Name = ""
			End If
			aToolbar(i) = aCurrToolbar(0) & "|||" & aCurrToolbar(1) & "|||" & s_Name & "|||" & s_Order
		End If
	Next

	Call WriteConfig()
	Call WriteStyle(Clng(sStyleID))

	Response.Write "<script language=javascript>alert('�������޸Ĳ����ɹ���');</script>"
	Call GoUrl("?action=toolbar&id=" & sStyleID)

End Sub

Sub DoToolBarDel()
	Dim s_DelID
	s_DelID = Trim(Request("delid"))
	If IsNumeric(s_DelID) = True Then
		aToolbar(Clng(s_DelID)) = ""
		Call WriteConfig()
		Call WriteStyle(Clng(sStyleID))
		Response.Write "<script language=javascript>alert('��������ID��" & s_DelID & "��ɾ�������ɹ���');</script>"
		Call GoUrl("?action=toolbar&id=" & sStyleID)
	End If
End Sub

Sub InitToolBar()
	Dim b, aCurrToolbar, nToolbarID
	b = False
	sToolBarID = Trim(Request("toolbarid"))
	If IsNumeric(sToolBarID) = True Then
		If Clng(sToolBarID) <= UBound(aToolbar) And Clng(sToolBarID) > 0 Then
			aCurrToolbar = Split(aToolbar(Clng(sToolbarID)), "|||")
			sToolBarName = aCurrToolbar(2)
			sToolBarOrder = aCurrToolbar(3)
			sToolBarButton = aCurrToolbar(1)
			b = True
		End If
	End If
	If b = False Then
		GoError "��Ч�Ĺ�����ID�ţ���ͨ��ҳ���ϵ����ӽ��в�����"
	End If
End Sub

Sub ShowButtonList()

	Call ShowMessage("<b class=blue>��ǰ��ʽ��<span class=red>" & outHTML(sStyleName) & "</span>&nbsp;&nbsp;��ǰ��������<span class=red>" & outHTML(sToolBarName) & "</span></b>")
	
	Dim i, n

	Dim s_Option1
	s_Option1 = ""
	For i = 1 To UBound(aButton, 1)
		If aButton(i, 8) = 1 Then
			s_Option1 = s_Option1 & "<option value='" & aButton(i, 1) & "'>" & aButton(i, 2) & "</option>"
		End If
	Next

	Dim aSelButton, s_Option2, s_Temp
	aSelButton = Split(sToolBarButton, "|")
	s_Option2 = ""
	For i = 0 To UBound(aSelButton)
		s_Temp = Code2Title(aSelButton(i))
		If s_Temp <> "" Then
			s_Option2 = s_Option2 & "<option value='" & aSelButton(i) & "'>" & s_Temp & "</option>"
		End If
	Next


%>

<script language=javascript>
function Add() {
	var sel1=document.myform.d_b1;
	var sel2=document.myform.d_b2;
	if (sel1.selectedIndex<0) {
		alert("��ѡ��һ����ѡ��ť��");
		return;
	}
	sel2.options[sel2.length]=new Option(sel1.options[sel1.selectedIndex].innerHTML,sel1.options[sel1.selectedIndex].value);
}

function Del() {
	var sel=document.myform.d_b2;
	var nIndex = sel.selectedIndex;
	var nLen = sel.length;
	if (nLen<1) return;
	if (nIndex<0) {
		alert("��ѡ��һ����ѡ��ť��");
		return;
	}
	for (var i=nIndex;i<nLen-1;i++) {
		sel.options[i].value=sel.options[i+1].value;
		sel.options[i].innerHTML=sel.options[i+1].innerHTML;
	}
	sel.length=nLen-1;
}

function Up() {
	var sel=document.myform.d_b2;
	var nIndex = sel.selectedIndex;
	var nLen = sel.length;
	if ((nLen<1)||(nIndex==0)) return;
	if (nIndex<0) {
		alert("��ѡ��һ��Ҫ�ƶ�����ѡ��ť��");
		return;
	}
	var sValue=sel.options[nIndex].value;
	var sHTML=sel.options[nIndex].innerHTML;
	sel.options[nIndex].value=sel.options[nIndex-1].value;
	sel.options[nIndex].innerHTML=sel.options[nIndex-1].innerHTML;
	sel.options[nIndex-1].value=sValue;
	sel.options[nIndex-1].innerHTML=sHTML;
	sel.selectedIndex=nIndex-1;
}

function Down() {
	var sel=document.myform.d_b2;
	var nIndex = sel.selectedIndex;
	var nLen = sel.length;
	if ((nLen<1)||(nIndex==nLen-1)) return;
	if (nIndex<0) {
		alert("��ѡ��һ��Ҫ�ƶ�����ѡ��ť��");
		return;
	}
	var sValue=sel.options[nIndex].value;
	var sHTML=sel.options[nIndex].innerHTML;
	sel.options[nIndex].value=sel.options[nIndex+1].value;
	sel.options[nIndex].innerHTML=sel.options[nIndex+1].innerHTML;
	sel.options[nIndex+1].value=sValue;
	sel.options[nIndex+1].innerHTML=sHTML;
	sel.selectedIndex=nIndex+1;
}

function checkform() {
	var sel=document.myform.d_b2;
	var nLen = sel.length;
	var str="";
	for (var i=0;i<nLen;i++) {
		if (i>0) str+="|";
		str+=sel.options[i].value;
	}
	document.myform.d_button.value=str;
	return true;
}

</script>

<%


	Dim s_SubmitButton
	If nStyleIsSys = 1 Then
		s_SubmitButton = ""
	Else
		s_SubmitButton = "<input type=submit name=b value=' �������� '>"
	End If
	Response.Write "<table border=0 cellpadding=5 cellspacing=0 align=center>" & _
		"<form action='?action=buttonsave&id=" & sStyleID & "&toolbarid=" & sToolBarID & "' method=post name=myform onsubmit='return checkform()'>" & _
		"<tr align=center><td>��ѡ��ť</td><td></td><td>��ѡ��ť</td><td></td></tr>" & _
		"<tr align=center>" & _
			"<td><select name='d_b1' size=20 style='width:250px' ondblclick='Add()'>" & s_Option1 & "</select></td>" & _
			"<td><input type=button name=b1 value=' �� ' onclick='Add()'><br><br><input type=button name=b1 value=' �� ' onclick='Del()'></td>" & _
			"<td><select name='d_b2' size=20 style='width:250px' ondblclick='Del()'>" & s_Option2 & "</select></td>" & _
			"<td><input type=button name=b3 value='��' onclick='Up()'><br><br><br><input type=button name=b4 value='��' onclick='Down()'></td>" & _
		"</tr>" & _
		"<input type=hidden name='d_button' value=''>" & _
		"<tr><td colspan=4 align=right>" & s_SubmitButton & "</td></tr>" & _
		"</form></table>"


	Response.Write "<table border=0 cellspacing=1 align=center class=list>" & _
		"<tr><th colspan=4>�����ǰ�ťͼƬ���ձ���������������ⰴť����ûͼ����</th></tr>"
	n = 0
	Dim m
	m = 0
	For i = 1 To UBound(aButton)
		If aButton(i, 8) = 1 Then
			m = m + 1
			n = m Mod 4
			If n = 1 Then
				Response.Write "<tr>"
			End If
			Response.Write "<td>"
			If aButton(i, 3) <> "" Then
				Response.Write "<img border=0 align=absmiddle src='../buttonimage/" & sStyleDir & "/" & aButton(i, 3) & "'>"
			End If
			Response.Write aButton(i, 2)
			Response.Write "</td>"
			If n = 0 Then
				Response.Write "</tr>"
			End If
		End If
	Next
	If n > 0 Then
		For i = 1 To 4 - n
			Response.Write "<td>&nbsp;</td>"
		Next
		Response.Write "</tr>"
	End if
	Response.Write "</table>"
End Sub

Function Code2Title(s_Code)
	Dim i
	Code2Title = ""
	For i = 1 To UBound(aButton)
		If UCase(aButton(i, 1)) = UCase(s_Code) Then
			Code2Title = aButton(i, 2)
			Exit Function
		End If
	Next
End Function

Sub DoButtonSave()
	Dim s_Button, nToolBarID, aCurrToolbar
	s_Button = Trim(Request("d_button"))

	nToolBarID = Clng(sToolBarID)
	aCurrToolbar = Split(aToolbar(nToolBarID), "|||")
	aToolbar(nToolBarID) = aCurrToolbar(0) & "|||" & s_Button & "|||" & aCurrToolbar(2) & "|||" & aCurrToolbar(3)

	Call WriteConfig()
	Call WriteStyle(Clng(sStyleID))

	Call ShowMessage("<b><span class=red>��������ť���ñ���ɹ���</span></b><li><a href='?action=stylepreview&id=" & sStyleID & "' target='_blank'>Ԥ������ʽ</a><li><a href='?action=toolbar&id=" & sStyleID & "'>���ع���������</a><li><a href='?action=buttonset&id=" & sStyleID & "&toolbarid=" & sToolBarID & "'>�������ô˹������µİ�ť</a>")

End Sub

Function InitSelect(s_FieldName, a_Name, a_Value, v_InitValue, s_Sql, s_AllName)
	Dim i
	InitSelect = "<select name='" & s_FieldName & "' size=1>"
	If s_AllName <> "" Then
		InitSelect = InitSelect & "<option value=''>" & s_AllName & "</option>"
	End If
	If s_Sql <> "" Then
		oRs.Open s_Sql, oConn, 0, 1
		Do While Not oRs.Eof
			InitSelect = InitSelect & "<option value=""" & inHTML(oRs(1)) & """"
			If oRs(1) = v_InitValue Then
				InitSelect = InitSelect & " selected"
			End If
			InitSelect = InitSelect & ">" & outHTML(oRs(0)) & "</option>"
			oRs.MoveNext
		Loop
		oRs.Close
	Else
		For i = 0 To UBound(a_Name)
			InitSelect = InitSelect & "<option value=""" & inHTML(a_Value(i)) & """"
			If a_Value(i) = v_InitValue Then
				InitSelect = InitSelect & " selected"
			End If
			InitSelect = InitSelect & ">" & outHTML(a_Name(i)) & "</option>"
		Next
	End If
	InitSelect = InitSelect & "</select>"
End Function



Sub WriteStyle(n_StyleID)
	Dim sConfig
	Dim aTmpStyle

	sConfig = ""
	aTmpStyle = Split(aStyle(n_StyleID), "|||")
	sConfig = sConfig & "config.ButtonDir = """ & aTmpStyle(1) & """;" & Vbcrlf
	sConfig = sConfig & "config.StyleUploadDir = """ & aTmpStyle(3) & """;" & Vbcrlf
	sConfig = sConfig & "config.InitMode = """ & aTmpStyle(18) & """;" & Vbcrlf
	sConfig = sConfig & "config.AutoDetectPasteFromWord = """ & aTmpStyle(17) & """;" & Vbcrlf
	sConfig = sConfig & "config.BaseUrl = """ & aTmpStyle(19) & """;" & Vbcrlf
	sConfig = sConfig & "config.BaseHref = """ & aTmpStyle(22) & """;" & Vbcrlf
	sConfig = sConfig & "config.AutoRemote = """ & aTmpStyle(24) & """;" & Vbcrlf
	sConfig = sConfig & "config.ShowBorder = """ & aTmpStyle(25) & """;" & Vbcrlf
	sConfig = sConfig & "config.StateFlag = """ & aTmpStyle(16) & """;" & Vbcrlf
	sConfig = sConfig & "config.CssDir = """ & aTmpStyle(2) & """;" & Vbcrlf
	sConfig = sConfig & "config.AutoDetectLanguage = """ & aTmpStyle(27) & """;" & Vbcrlf
	sConfig = sConfig & "config.DefaultLanguage = """ & aTmpStyle(28) & """;" & Vbcrlf
	sConfig = sConfig & Vbcrlf
	sConfig = sConfig & "function showToolbar(){" & Vbcrlf
	sConfig = sConfig & Vbcrlf

	sConfig = sConfig & chr(9) & "document.write ("""
	sConfig = sConfig & "<table border=0 cellpadding=0 cellspacing=0 width='100%' class='Toolbar' id='eWebEditor_Toolbar'>"

	Dim s_Order, s_ID, n, aTmpToolbar
	s_Order = ""
	s_ID = ""
	For n = 1 To UBound(aToolbar)
		If aToolbar(n) <> "" Then
			aTmpToolbar = Split(aToolbar(n), "|||")
			If aTmpToolbar(0) = CStr(n_StyleID) Then
				If s_ID <> "" Then
					s_ID = s_ID & "|"
					s_Order = s_Order & "|"
				End If
				s_ID = s_ID & CStr(n)
				s_Order = s_Order & aTmpToolbar(3)
			End If
		End If
	Next

	Dim a_ID, a_Order, aTmpButton, i
	If s_ID <> "" Then
		a_ID = Split(s_ID, "|")
		a_Order = Split(s_Order, "|")
		For n = 0 To UBound(a_Order)
			a_Order(n) = Clng(a_Order(n))
			a_ID(n) = Clng(a_ID(n))
		Next
		a_ID = Sort(a_ID, a_Order)
		For n = 0 To UBound(a_ID)
			aTmpToolbar = Split(aToolbar(a_ID(n)), "|||")
			aTmpButton = Split(aTmpToolbar(1), "|")

			sConfig = sConfig & "<tr><td><div class=yToolbar>"
			For i = 0 To UBound(aTmpButton)
				If UCase(aTmpButton(i)) = "MAXIMIZE" Then
					sConfig = sConfig & """);" & Vbcrlf
					sConfig = sConfig & Vbcrlf

					sConfig = sConfig & chr(9) & "if (sFullScreen==""1""){" & Vbcrlf
					sConfig = sConfig & chr(9) & chr(9) & "document.write (""" & Code2HTML("Minimize", aTmpStyle(1)) & """);" & Vbcrlf
					sConfig = sConfig & chr(9) & "}else{" & Vbcrlf
					sConfig = sConfig & chr(9) & chr(9) & "document.write (""" & Code2HTML(aTmpButton(i), aTmpStyle(1)) & """);" & Vbcrlf
					sConfig = sConfig & chr(9) & "}" & Vbcrlf
					sConfig = sConfig & Vbcrlf

					sConfig = sConfig & chr(9) & "document.write ("""
				Else
					sConfig = sConfig & Code2HTML(aTmpButton(i), aTmpStyle(1))
				End If
			Next
			sConfig = sConfig & "</div></td></tr>"
		Next
	Else
		sConfig = sConfig & "<tr><td></td></tr>"
	End If

	sConfig = sConfig & "</table>"");" & Vbcrlf
	sConfig = sConfig & Vbcrlf
	sConfig = sConfig & "}" & Vbcrlf	

	Call WriteFile("../style/" & LCase(aTmpStyle(0)) & ".js", sConfig)

End Sub

Function Code2HTML(s_Code, s_ButtonDir)
	Dim i
	Code2HTML = ""
	For i = 1 To UBound(aButton, 1)
		If UCase(aButton(i, 1)) = UCase(s_Code) Then
			Select Case aButton(i, 5)
			Case 0
				Code2HTML = "<DIV CLASS=" & aButton(i, 7) & " TITLE='""+lang[""" & aButton(i, 1) & """]+""' onclick=\""" & aButton(i, 6) & "\""><IMG CLASS=Ico SRC='buttonimage/" & s_ButtonDir & "/" & aButton(i, 3) & "'></DIV>"
			Case 1
				If aButton(i, 4) <> "" Then
					Code2HTML = "<SELECT CLASS=" & aButton(i, 7) & " onchange=\""" & aButton(i, 6) & "\"">" & aButton(i, 4) & "</SELECT>"
				Else
					Code2HTML = "<SELECT CLASS=" & aButton(i, 7) & " onchange=\""" & aButton(i, 6) & "\"">""+lang[""" & aButton(i, 1) & """]+""</SELECT>"
				End If
			Case 2
				Code2HTML = "<DIV CLASS=" & aButton(i, 7) & ">" & aButton(i, 4) & "</DIV>"
			End Select
			Exit Function
		End If
	Next
End Function

Sub DeleteFile(s_StyleName)
	On Error Resume Next
	Dim oFSO, sMapFileName
	Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
	sMapFileName = Server.MapPath("../style/" & LCase(s_StyleName) & ".js")
	If oFSO.FileExists(sMapFileName) Then
		oFSO.DeleteFile(sMapFileName)
	End If
	Set oFSO = Nothing
End Sub

Sub GoUrl(url)
	Response.Write "<script language=javascript>location.href=""" & url & """;</script>"
	Response.End
End Sub

Function isValidColor(str)
	Dim re
	Set re = new RegExp
	re.Pattern = "[A-Fa-f0-9]{6}"
	isValidColor = re.Test(str)
	Set re = Nothing
End Function
%>