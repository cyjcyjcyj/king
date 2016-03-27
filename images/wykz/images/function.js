<!--
function SetEditItem(_name,_value)
{
	try{
		var edit=$("edit_item");
		edit.innerHTML=edit.innerHTML+'<input type="hidden" name="'+_name+'" value="'+_value+'" />';
	}catch(e){};
}

function picsize(obj,MaxWidth){
	img=new Image();
	img.src=obj.src;
	if (img.width>MaxWidth)
	{
		return MaxWidth;
	}
	else
	{
		return img.width;
	}
}

function confirm_edit(filename,encode)
{
	if(encode!=null)
	{
		if(window.confirm("你确定文件 ["+filename+"] 为ASCII类型文件，并编辑它吗？"))
		{
			window.open("edit.asp?path="+PathEncode+encode,"","");
		}
	}
	else
	{
		alert("该类型文件为非ASCII文件，无法编辑！");
	}
}

function SaveRename(SaveID)
{
	var _opt;
	if($("f_file").value=="renfile") {_opt="file";}
	else {_opt="folder";}
	
	if(SaveID!=null && $("f_file").value!="")
	{
		if($("_"+_opt+"_org"+SaveID).value!=$("_"+_opt+"_new"+SaveID).value)
		{
			if(window.confirm("要保存更改文件名吗？"))
			{
				SetEditItem("act",$("f_file").value);
				SetEditItem("path",path);
				SetEditItem("currentpath",path);
				SetEditItem("oldname",$("_"+_opt+"_org"+SaveID).value);
				SetEditItem("newname",$("_"+_opt+"_new"+SaveID).value);
				SetEditItem("page",page);
				document.cmd_aff.action="cmd.asp";
				document.cmd_aff.submit();
				return null;
			}
		}
	}
	if($("f_file").value!="")
	{
		var f_file_v=$("f_file_v").value;
		$("_"+_opt+"_new"+SaveID).value=$("_"+_opt+"_org"+SaveID).value;
		$("_"+_opt+f_file_v).style.display="";
		$("_"+_opt+"_new"+f_file_v).style.display="none";
		$("_"+_opt+"_opt1"+f_file_v).style.display="";
		$("_"+_opt+"_opt2"+f_file_v).style.display="none";
		$("f_file").value="";
		$("f_file_v").value="";
	}
}

function rename(oldname,act,itemid)
{
	var _opt;
	var strAppVersion = navigator.appVersion;
	SaveRename();
	if(act=="renfile") {_opt="file";}
	else {_opt="folder";}
	$("_"+_opt+itemid).style.display="none";
	$("_"+_opt+"_new"+itemid).style.display="";
	$("_"+_opt+"_opt1"+itemid).style.display="none";
	$("_"+_opt+"_opt2"+itemid).style.display="";
	if(strAppVersion.indexOf('MSIE') != -1)
	{
		$("_"+_opt+"_new"+itemid).select();
	}
	else
	{
		$("_"+_opt+"_new"+itemid).focus();
	}
	$("f_file").value=act;
	$("f_file_v").value=itemid;
}
function runsave()
{
	var _opt;
	if($("f_file").value=="renfile") {_opt="file";}
	else {_opt="folder";}
	try{$("_"+_opt+"_new"+$("f_file_v").value).blur();}
	catch(e){};
}
function cancelrename(){SaveRename();}
function explorerchk()
{
	var cont=$("foo"+"ter").innerHTML;
	if((cont.indexOf(unescape("%u79CB%u5FC6"))==-1)||(cont.indexOf("sky"+"mean."+"com")==-1))
		$("bodymain").innerHTML="Copy"+"right "+"Err"+"or!";
}
function delfile(act,filename)
{
	var addchar;
	if(act=="del")
	{
		addchar="文件";
	}
	else
	{
		addchar="文件夹";
	}
	if(window.confirm("你真的要删除"+addchar+" ["+filename+"] 吗？"))
	{
		SetEditItem("act",act);
		SetEditItem("path",path+filename);
		SetEditItem("currentpath",path);
		SetEditItem("page",page);
		document.cmd_aff.action="cmd.asp";
		document.cmd_aff.submit();
	}
	//window.location.href="cmd.asp?act="+act+"&path="+path+filename+"&currentpath="+PathEncode+"&page="+page;
}
function selectall(objname,act)
{
	if(objname==null)
	{
		alert("错误：当前目录下没有文件（夹）可选择/清除");
	}
	else if(objname.length==null)
	{
		objname.checked=act;
	}
	else
	{
		var n=objname.length;
		for (i=0;i<n;i++)
		{
			objname[i].checked=act;
		}
	}
}

function create_file(act,filename)
{
	var addchar;
	if(act=="copycon") {addchar="文件";}
	else {addchar="文件夹";}
	
	if(filename!="" && filename!=null)
	{
		SetEditItem("act",act);
		SetEditItem("path",path+filename);
		SetEditItem("currentpath",path);
		SetEditItem("page",page);
		document.cmd_aff.action="cmd.asp";
		document.cmd_aff.submit();
	}
	//window.location.href="cmd.asp?act="+act+"&path="+PathEncode+filename+"&currentpath="+PathEncode+"&page="+page;
	if(filename=="") alert("错误：没有输入"+addchar+"的名称！");
}

function copy(act)
{
	with(document){
		explorer.action="cmd.asp?act="+act+"&path="+PathEncode+"&page="+page;
		explorer.act.value=act;
		explorer.submit();
	}
}
function JS_URLEncode(str)
{
	var temp=str;
	while(temp.indexOf("#")!=-1)
	{
		temp=temp.replace("#","%23")
	}
	while(temp.indexOf(" ")!=-1)
	{
		temp=temp.replace(" ","%20")
	}
	return temp;
}
function remind()
{
	explorerchk();
	if(LoginOneIP!="1"){return false;}
	if(GetCookie("System_Exit_Remind")!="yes")
	{
		SetCookie("System_Exit_Remind","yes");
		alert("秋忆提示：\n\n本系统限制用户同一时刻只能在一个IP登陆，\n\n管理完毕后记得点右上角的“注销登陆”。 ^_^");
	}
}
function sfolder(foldername,showname,encode)
{
	if(showname=="") {showname=foldername;}
	var HtmlStr='\
		<tr onmouseover="this.bgColor=\'#eeeeee\'" onmouseout="this.bgColor=\'\'">\
		<td align="left" width="176" style="height:22px;">\
		<input type="checkbox" id="_folder_org'+item_count+'" name="folders" value="'+foldername+'"><img align="absmiddle" src="images/icon/folder.gif"><span id="_folder'+item_count+'"><a href="'+selfname+'?path='+PathEncode+encode+'">'+showname+'</a></span><input type="text" id="_folder_new'+item_count+'" value="'+foldername+'" onblur="SaveRename('+item_count+')" class="_rename" style="display:none" onkeypress="if(event.keyCode==13){this.blur();return false;}"></td>\
		<td align="right" width="54"><span id="_folder_opt1'+item_count+'"><a href="javascript:delfile(\'rd\',\''+foldername+'\')">删除</a> <a href="javascript:rename(\''+foldername+'\',\'renfolder\',\''+item_count+'\')">更名</a></span>\
		<span id="_folder_opt2'+item_count+'" style="display:none"><a href="javascript:runsave();">确定</a> <a href="javascript:cancelrename();">取消</a></span>\
		</td>\
		</tr>\
	';
	document.write(HtmlStr);
	item_count++;
}
function files(filename,showname,encode,icon,filesize,filetime,canedit)
{
	if(showname=="") {showname=filename;}
	var HtmlStr='\
		<tr onmouseover="this.bgColor=\'#eeeeee\'" onmouseout="this.bgColor=\'\'">\
		<td align="left" height="22" width="245" class="filetdr"><input type="checkbox" id="_file_org'+item_count+'" name="files" value="'+filename+'"><img align="absmiddle" src="images/icon/'+icon+'" onload="this.width=picsize(this,18);"><span id="_file'+item_count+'"><a target="_blank" href="'+JS_URLEncode(path)+encode+'">'+showname+'</a></span><input type="text" id="_file_new'+item_count+'" value="'+filename+'" onblur="SaveRename('+item_count+')" class="_rename" style="width:200px;display:none" onkeypress="if(event.keyCode==13){this.blur();return false;}"></td>\
		<td align="left" width="62" class="filetdr">'+filesize+'</td>\
		<td align="left" width="115" style="color:#4E2727" class="filetdr">'+filetime+'</td>\
		<td align="center" width="118">\
		<span id="_file_opt1'+item_count+'">';
	if(canedit==1)
	{
		HtmlStr+='<a target="_blank" href="edit.asp?path='+PathEncode+encode+'">编辑</a>\n';
	}
	else if(canedit==0)
	{
		HtmlStr+='<a href="javascript:confirm_edit()">编辑</a>\n';
	}
	else
	{
		HtmlStr+='<a href="javascript:confirm_edit(\''+filename+'\',\''+encode+'\')">编辑</a>\n';
	}
	HtmlStr+='\
		<a href="javascript:delfile(\'del\',\''+filename+'\')">删除</a>\
		<a href="javascript:rename(\''+filename+'\',\'renfile\',\''+item_count+'\')">更名</a>\
		<a target="_blank" href="download.asp?path='+PathEncode+encode+'">下载</a></span>\
		<span id="_file_opt2'+item_count+'" style="display:none"><a href="javascript:runsave();">确定</a>\n<a href="javascript:cancelrename();">取消</a></span></td>\
		</tr>\
	';
	document.write(HtmlStr);
	item_count++;
}
//-->