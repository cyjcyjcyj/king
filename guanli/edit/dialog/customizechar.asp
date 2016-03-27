<% Option Explicit %>
<%
Session("eWebEditor_Original_CodePage") = Session.CodePage
Session.CodePage = 65001
Call ShowMain()
Session.CodePage = Session("eWebEditor_Original_CodePage")




Sub ShowMain()
%>
<HTML>
<HEAD>
<META http-equiv=Content-Type content="text/html; charset=utf-8">
<script language="javascript" src="dialog.js"></script>

<script language=javascript>
document.write ("<link href='../language/" + AvailableLangs["Active"] + ".css' type='text/css' rel='stylesheet'>");
document.write ("<link href='dialog.css' type='text/css' rel='stylesheet'>");

document.write ("<TITLE>" + lang["DlgEmot"] + "</TITLE>");

function emotclick(){
	if ("IMG"==event.srcElement.tagName.toUpperCase()) {
		var url = event.srcElement.getAttribute("src", 2).substr(17);
		dialogArguments.insertHTML("<IMG id=eWebEditor_TempElement_Img SRC=" + url + ">");

		var oTempElement = dialogArguments.eWebEditor.document.getElementById("eWebEditor_TempElement_Img");
		oTempElement.src = url;
		oTempElement.removeAttribute("id");

		window.returnValue = null;
		window.close();
	}
}
document.onclick=emotclick;

function InitDocument(){
	adjustDialog();
}
</script>
<title>插入其他特殊符号</title></HEAD>

<BODY onload="InitDocument()">
<table border=0 cellpadding=0 cellspacing=5 id=tabDialogSize><tr><td>

<%
Call ShowList()
%>

</td></tr></table>
</BODY></HTML>

<%

End Sub

Sub ShowList()
	On Error Resume Next
	If IsObjInstalled("Scripting.FileSystemObject") = False Then Exit Sub

	Dim oFSO, oUploadFolder, oUploadFiles, oUploadFile
	Set oFSO = Server.CreateObject("Scripting.FileSystemObject")
	Set oUploadFolder = oFSO.GetFolder(Server.MapPath("../../Periodical/Character/"))
	Set oUploadFiles = oUploadFolder.Files

	Dim s_List, n, m
	s_List = ""
	n = 0
	m = 0
	For Each oUploadFile In oUploadFiles
		n = n + 1
		m = n mod 10
		If m = 1 Then
			s_List = s_List & "<tr>"
		End If
		s_List = s_List & "<td width=20 height=20><img border=0 src='../../periodical/character/" & oUploadFile.Name & "'></td>"		
		If m = 0 Then
			s_List = s_List & "</tr>"
		End If
	Next
	If m > 0 Then
		For n = 1 To 10-m
			s_List = sList & "<td width=20 height=20>&nbsp;</td>"
		Next
		s_List = sList & "</tr>"
	End If
	If s_List <> "" Then
		s_List = "<TABLE cellSpacing=5 border=0 align=center bgcolor=#FFFFFF><TBODY>" & s_List & "</TBODY></TABLE>"
	End If
	Response.Write s_List
	Set oUploadFolder = Nothing
	Set oUploadFiles = Nothing
End Sub


Function IsObjInstalled(strClassString)
	On Error Resume Next
	IsObjInstalled = False
	Err = 0
	Dim xTestObj
	Set xTestObj = Server.CreateObject(strClassString)
	If 0 = Err Then IsObjInstalled = True
	Set xTestObj = Nothing
	Err = 0
End Function

%>