<!--
function NoGroupwareProgress()
{
	var Form = document.form1;
	var ProgressID = (new Date()).getTime() % 1000000000;
	var Ver = navigator.appVersion;
	Form.action = Form.action + "&ProgressID=" + ProgressID;
	
	if (Ver.indexOf("MSIE") > -1 && Ver.substr(Ver.indexOf("MSIE") + 5, 1) > 4) {
		window.showModelessDialog("Progress.asp?Count=0&ProgressID=" + ProgressID, null, "dialogWidth=360px; dialogHeight:180px; help:no; status:no");
	} 
	else 
	{
		window.open("Progress.asp?Count=0&ProgressID=" + ProgressID, "_blank", "left=240,top=240,width=360,height=160,menubar=no,toolbar=no,location=no,scrollbars=no,status=no,dialog=yes");
	}
	document.getElementById("submitbutton").disabled=true;
	Form.submit();
	return true;
}

function AspUpProgress()
{
	var Form = document.form1;
	var strAppVersion = navigator.appVersion;
	Form.action = Form.action + "&action=uploadfile&PID=" + uploadpid;
	
	if(strAppVersion.indexOf('MSIE') != -1 && strAppVersion.substr(strAppVersion.indexOf('MSIE')+5,1) > 4)
	{
		winstyle = "dialogWidth=375px; dialogHeight:175px; center:yes;status:no";
		window.showModelessDialog("framebar.asp?to=10&b=IE&PID="+uploadpid,window,winstyle);
	}
	else
	{
		window.open("framebar.asp?to=10&b=NN&PID="+uploadpid,"","width=370,height=165,menubar=no,toolbar=no,location=no,scrollbars=no,status=no,dialog=yes", true);
	}
	document.getElementById("submitbutton").disabled=true;
	Form.submit();
	return true;
}

function SetInput()
{
	var str='<br>';
	var upcount=document.getElementById("upcount").value;
	if(!(upcount>'0'&&upcount<='9'))
	{
		upcount=1;
		document.getElementById("upcount").value=upcount;
	}
	document.getElementById("upcount2").value=upcount;
	for(var i=1;i<=upcount;i++)
	{
		str+='<span id="upspan'+i+'">文件';
		if(i<10) {str+='0'+i;}
		else {str+=i;}
		str+=':<input type="file" id="file'+i+'" name="file'+i+'"';
		if(IsMSIE) {str+=' class="tx1"';}
		else {str+=' size="60"';}
		str+='>&nbsp;<a href="javascript:delinput('+i+')">删除</a><br><br></span>';
	}
	document.getElementById("upid").innerHTML=str+'<br>';
}

function UpSubmit()
{
	var upcount2=document.getElementById("upcount2").value;
	document.getElementById("realpath").value=document.getElementById("visualpath").value;
	if(document.getElementById("realpath").value=="")
	{
		alert("上传位置不能为空！");
		return false;
	}

	if(uploadobj==0)
	{
		NoGroupwareProgress();
	}
	else if(uploadobj==1)
	{
		AspUpProgress();
	}
	else
	{
		document.getElementById("submitbutton").disabled=true;
		document.form1.submit();
	}
}
function resetform(uppath)
{
	document.getElementById("visualpath").value=uppath;
	document.form1.reset();
}
function delinput(inputid)
{
	if(uploadobj==3)
		alert("使用[LyfUpload组件]不能删除上传个数！");
	else
	{
		var upcount2=document.getElementById("upcount2").value;
		if(upcount2==1)
		{
			alert("至少要留一个吧！别那么狠心全删除~~~");
		}
		else
		{
			document.getElementById("file"+inputid).disabled=true;
			document.getElementById("upspan"+inputid).innerHTML="";
			document.getElementById("upspan"+inputid).style.display="none";
			document.getElementById("upcount").value=upcount2-1;
			document.getElementById("upcount2").value=upcount2-1;
		}
	}
}
//-->