
<HTML><HEAD><TITLE>网站后台管理</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">

<LINK href="images/Admin_Style.css" type=text/css rel=stylesheet>


<STYLE>BODY {
	FONT-SIZE: 12px; FONT-FAMILY: "宋体"; TEXT-DECORATION: none
}
TD {
	FONT-SIZE: 12px; COLOR: #666666; LINE-HEIGHT: 18px; FONT-FAMILY: "宋体"; TEXT-DECORATION: none
}
.S1 {
	FONT-WEIGHT: bold; FONT-SIZE: 16px; color: #DBBF90; FONT-FAMILY: "宋体"; TEXT-DECORATION: none
}
</STYLE>

<META content="MSHTML 6.00.2800.1555" name=GENERATOR></HEAD>
<BODY bgColor=#ffffff leftMargin=0 topMargin=0 marginheight="0" 
marginwidth="0"><BR><BR><BR>
<TABLE id=__01 height=425 cellSpacing=0 cellPadding=0 width=646 align=center 
border=0>
  <TBODY>
  <TR>
    <TD colSpan=3><IMG height=114 alt="" 
      src="images/login_01.gif" width=646></TD></TR>
  <TR>
    <TD><IMG height=311 alt="" src="images/login_02.gif" 
      width=88></TD>
    <TD vAlign=top width=476 background=images/login_03.gif 
    height=311><BR>
      <FORM name=LoginForm onsubmit=return(CheckForm(this)) 
      action=gly_yz.asp method=post>
      <TABLE cellSpacing=0 cellPadding=0 width="80%" align=center border=0>
        <TBODY>
        <TR>
          <TD align=right width=87 height=30>登录名：</TD>
          <TD vAlign=center><INPUT 
            name=yhm class=textbox id="yhm" tabIndex=1 size=21> </TD></TR>
        <TR>
          <TD align=right height=30>密　码：</TD>
          <TD><span class="table-xia">
            <input name="mm" type="Password" class="textbox" size="15" maxlength="15">
          </span> 
          </TD>
        </TR>
        <TR>
          <TD align=right height=50>验证码：</TD>
          <TD height=50><INPUT 
            name=yzm class=textbox id="yzm" tabIndex=3 size=6 maxLength=6> 
          <FONT color=red>&lt;-请在验证码框输入</FONT> <img src="code.asp" width="38" height="10">&nbsp;</TD></TR>
        <TR>
          <TD align=middle colSpan=2 height=50><INPUT 
            onmouseover=nereidFade(this,100,10,5) 
            style="FILTER: alpha(opacity=50)" tabIndex=5 
            onmouseout=nereidFade(this,50,10,5) type=image 
            src="images/dl.gif" border=0 name=enter> 
            &nbsp;&nbsp;<IMG onmouseover=nereidFade(this,100,10,5) 
            style="FILTER: alpha(opacity=50); CURSOR: hand" 
            onclick="javascript:location.href='http://192.168.1.101:83/';" 
            tabIndex=5 onmouseout=nereidFade(this,50,10,5) 
            src="images/fh.gif" border=0 name=enter2 
            type="button"></TD></TR>
        <TR>
          <TD colSpan=2></TD></TR></TBODY></TABLE>
      </FORM>
      <HR align=left width="80%" color=#efefef SIZE=1>
      　</TD>
    <TD><IMG height=311 alt="" src="images/login_04.gif" 
      width=82></TD></TR></TBODY></TABLE>
<SCRIPT language=javascript>
<!--
function document.onreadystatechange()
{  var app=navigator.appName;
  var verstr=navigator.appVersion;
  if(app.indexOf('Netscape') != -1) {
    alert('友情提示：\n    您使用的是Netscape浏览器，可能会导致无法使用后台的部分功能。建议您使用 IE6.0 或以上版本。');
  } else if(app.indexOf('Microsoft') != -1) {
    if (verstr.indexOf('MSIE 3.0')!=-1 || verstr.indexOf('MSIE 4.0') != -1 || verstr.indexOf('MSIE 5.0') != -1 || verstr.indexOf('MSIE 5.1') != -1)
      alert('友情提示：\n    您的浏览器版本太低，可能会导致无法使用后台的部分功能。建议您使用 IE6.0 或以上版本。');
  }
  document.LoginForm.yhm.focus();
}
function CheckForm(ObjForm) {
  if(ObjForm.yhm.value == '') {
    alert('请输入管理账号！');
    ObjForm.yhm.focus();
    return false;
  }
  if(ObjForm.mm.value == '') {
    alert('请输入密码！');
    ObjForm.mm.focus();
    return false;
  }
  if (ObjForm.yzm.value.length=="")
  {
    alert('验证码不能为空！');
    ObjForm.yzm.focus();
    return false;
  }
 
  
  
  
}





//-->
</SCRIPT>
 
 
 
 
 <SCRIPT language=javascript>
<!--
function checkspace(checkstr) {
  var str = '';
  for(i = 0; i < checkstr.length; i++) {
    str = str + ' ';
  }
  return (str == checkstr);
}
function check()
{
  if(checkspace(document.admininfo.admin.value)) {
	document.admininfo.admin.focus();
    alert("管理员用户名不能为空！");
	return false;
  }
  if(checkspace(document.admininfo.UserPassword.value)) {
	document.admininfo.UserPassword.focus();
    alert("密码不能为空！");
	return false;
  }
    if(checkspace(document.admininfo.passcode.value)) {
	document.admininfo.passcode.focus();
    alert("验证码不能为空！");
	return false;
  }
	document.admininfo.submit();
  }
  
  //定义当前是否大写的状态
window.onload=
	function()
	{
		password1=null;		
		initCalc();
	}

var CapsLockValue=0;

var check;
function setVariables() {
tablewidth=630;  // logo width, in pixels
tableheight=20;  // logo height, in pixels
if (navigator.appName == "Netscape") {
horz=".left";
vert=".top";
docStyle="document.";
styleDoc="";
innerW="window.innerWidth";
innerH="window.innerHeight";
offsetX="window.pageXOffset";
offsetY="window.pageYOffset";
}
else {
horz=".pixelLeft";
vert=".pixelTop";
docStyle="";
styleDoc=".style";
innerW="document.body.clientWidth";
innerH="document.body.clientHeight";
offsetX="document.body.scrollLeft";
offsetY="document.body.scrollTop";
   }
}
function checkLocation() {
if (check) {
objectXY="softkeyboard";
var availableX=eval(innerW);
var availableY=eval(innerH);
var currentX=eval(offsetX);
var currentY=eval(offsetY);
x=availableX-tablewidth+currentX;
//y=availableY-tableheight+currentY;
y=currentY;

evalMove();
}
setTimeout("checkLocation()",0);
}
function evalMove() {
//eval(docStyle + objectXY + styleDoc + horz + "=" + x);
eval(docStyle + objectXY + styleDoc + vert + "=" + y);
}
	self.onError=null;
	currentX = currentY = 0;  
	whichIt = null;           
	lastScrollX = 0; lastScrollY = 0;
	NS = (document.layers) ? 1 : 0;
	IE = (document.all) ? 1: 0;
	function heartBeat() {
		if(IE) { diffY = document.body.scrollTop; diffX = document.body.scrollLeft; }
	    if(NS) { diffY = self.pageYOffset; diffX = self.pageXOffset; }
		if(diffY != lastScrollY) {
	                percent = .1 * (diffY - lastScrollY);
	                if(percent > 0) percent = Math.ceil(percent);
	                else percent = Math.floor(percent);
					if(IE) document.all.softkeyboard.style.pixelTop += percent;
					if(NS) document.softkeyboard.top += percent; 
	                lastScrollY = lastScrollY + percent;}
		if(diffX != lastScrollX) {
			percent = .1 * (diffX - lastScrollX);
			if(percent > 0) percent = Math.ceil(percent);
			else percent = Math.floor(percent);
			if(IE) document.all.softkeyboard.style.pixelLeft += percent;
			if(NS) document.softkeyboard.left += percent;
			lastScrollX = lastScrollX + percent;	}		}
	function checkFocus(x,y) { 
	        stalkerx = document.softkeyboard.pageX;
	        stalkery = document.softkeyboard.pageY;
	        stalkerwidth = document.softkeyboard.clip.width;
	        stalkerheight = document.softkeyboard.clip.height;
	        if( (x > stalkerx && x < (stalkerx+stalkerwidth)) && (y > stalkery && y < (stalkery+stalkerheight))) return true;
	        else return false;}
	function grabIt(e) {
	    check = false;
		if(IE) {
			whichIt = event.srcElement;
			while (whichIt.id.indexOf("softkeyboard") == -1) {
				whichIt = whichIt.parentElement;
				if (whichIt == null) { return true; } }
			whichIt.style.pixelLeft = whichIt.offsetLeft;
		    whichIt.style.pixelTop = whichIt.offsetTop;
			currentX = (event.clientX + document.body.scrollLeft);
	   		currentY = (event.clientY + document.body.scrollTop); 	
		} else { 
	        window.captureEvents(Event.MOUSEMOVE);
	        if(checkFocus (e.pageX,e.pageY)) { 
	                whichIt = document.softkeyboard;
	                StalkerTouchedX = e.pageX-document.softkeyboard.pageX;
	                StalkerTouchedY = e.pageY-document.softkeyboard.pageY;} }
	    return true;	}
	function moveIt(e) {
		if (whichIt == null) { return false; }
		if(IE) {
		    newX = (event.clientX + document.body.scrollLeft);
		    newY = (event.clientY + document.body.scrollTop);
		    distanceX = (newX - currentX);    distanceY = (newY - currentY);
		    currentX = newX;    currentY = newY;
		    whichIt.style.pixelLeft += distanceX;
		    whichIt.style.pixelTop += distanceY;
			if(whichIt.style.pixelTop < document.body.scrollTop) whichIt.style.pixelTop = document.body.scrollTop;
			if(whichIt.style.pixelLeft < document.body.scrollLeft) whichIt.style.pixelLeft = document.body.scrollLeft;
			if(whichIt.style.pixelLeft > document.body.offsetWidth - document.body.scrollLeft - whichIt.style.pixelWidth - 20) whichIt.style.pixelLeft = document.body.offsetWidth - whichIt.style.pixelWidth - 20;
			if(whichIt.style.pixelTop > document.body.offsetHeight + document.body.scrollTop - whichIt.style.pixelHeight - 5) whichIt.style.pixelTop = document.body.offsetHeight + document.body.scrollTop - whichIt.style.pixelHeight - 5;
			event.returnValue = false;
		} else { 
			whichIt.moveTo(e.pageX-StalkerTouchedX,e.pageY-StalkerTouchedY);
	        if(whichIt.left < 0+self.pageXOffset) whichIt.left = 0+self.pageXOffset;
	        if(whichIt.top < 0+self.pageYOffset) whichIt.top = 0+self.pageYOffset;
        if( (whichIt.left + whichIt.clip.width) >= (window.innerWidth+self.pageXOffset-17)) whichIt.left = ((window.innerWidth+self.pageXOffset)-whichIt.clip.width)-17;
	        if( (whichIt.top + whichIt.clip.height) >= (window.innerHeight+self.pageYOffset-17)) whichIt.top = ((window.innerHeight+self.pageYOffset)-whichIt.clip.height)-17;
	        return false;}
	    return false;	}
	function dropIt() {whichIt = null;
	    if(NS) window.releaseEvents (Event.MOUSEMOVE);
	    return true;	}
	if(NS) {window.captureEvents(Event.MOUSEUP|Event.MOUSEDOWN);
		window.onmousedown = grabIt;
	 	window.onmousemove = moveIt;
		window.onmouseup = dropIt;	}
	if(IE) {
		document.onmousedown = grabIt;
	 	document.onmousemove = moveIt;
		document.onmouseup = dropIt;	}
//	if(NS || IE) action = window.setInterval("heartBeat()",1);



	document.write("<DIV align=center id=\"softkeyboard\" name=\"softkeyboard\" style=\"position:absolute; left:0px; top:0px; width:500px; z-index:180;display:none\">  <table id=\"CalcTable\" width=\"\" border=\"0\" align=\"center\" cellpadding=\"0\" cellspacing=\"0\" bgcolor=\"\">           <FORM id=Calc name=Calc action=\"\" method=post autocomplete=\"off\">       <tr> <td title=\"为了保证后台登陆安全,建议使用密码输入器输入密码!\" align=\"right\" valign=\"middle\" bgcolor=\"\" style=\"cursor: default;height:30\"> <INPUT type=hidden value=\"\" name=password>  <INPUT type=hidden value=ok name=action2>&nbsp<font style=\"font-size:13px;\">电子商务购物网站管理系统</font>&nbsp;&nbsp;密码输入器&nbsp&nbsp&nbsp&nbsp&nbsp;&nbsp;&nbsp&nbsp;&nbsp;&nbsp&nbsp;&nbsp;&nbsp&nbsp;&nbsp;&nbsp&nbsp;&nbsp;&nbsp&nbsp;&nbsp&nbsp;<INPUT style=\"width:100px;height:20px;background-color:#54BAF1;\" type=button value=\"使用键盘输入\" bgtype=\"1\" onclick=\"password1.readOnly=0;password1.focus();softkeyboard.style.display='none';password1.value='';\"><span style=\"width:2px;\"></span></td>      </tr>      <tr align=\"center\">         <td align=\"center\" bgcolor=\"#FFFFFF\"> <table align=\"center\" width=\"%\" border=\"0\" cellspacing=\"1\" cellpadding=\"0\">\n          <tr align=\"left\" valign=\"middle\"> \n            <td> <input type=button value=\" ~ \"></td>\n            <td> <input type=button value=\" ! \"></td>\n            <td> <input type=button  value=\" @ \"></td>\n            <td> <input type=button value=\" # \"></td>\n            <td> <input type=button value=\" $ \"></td>\n            <td> <input type=button value=\" % \"></td>\n            <td> <input type=button value=\" ^ \"></td>\n            <td> <input type=button value=\" & \"></td>\n            <td> <input type=button value=\" * \"></td>\n            <td> <input type=button value=\" ( \"></td>\n            <td> <input type=button value=\" ) \"></td>\n            <td> <input type=button value=\" _ \"></td>\n            <td> <input type=button value=\" + \"></td>\n            <td> <input type=button value=\" | \"></td>\n            <td colspan=\"1\" rowspan=\"2\"> <input name=\"button10\" type=button value=\" 退格\" onclick=\"setpassvalue();\"  onDblClick=\"setpassvalue();\" style=\"width:100px;height:42px\"> \n            </td>\n          </tr>\n          <tr align=\"left\" valign=\"middle\"> \n            <td> <input type=button value=\" ` \"></td>\n            <td> <input type=button value=\" 1 \"></td>\n            <td> <input type=button value=\" 2 \"></td>\n            <td> <input type=button value=\" 3 \"></td>\n            <td> <input type=button value=\" 4 \"></td>\n            <td> <input type=button value=\" 5 \"></td>\n            <td> <input type=button value=\" 6 \"></td>\n            <td> <input type=button value=\" 7 \"></td>\n            <td> <input type=button value=\" 8 \"></td>\n            <td> <input type=button value=\" 9 \"></td>\n            <td> <input name=\"button6\" type=button value=\" 0 \"></td>\n            <td> <input type=button value=\" - \"></td>\n            <td> <input type=button value=\" = \"></td>\n            <td> <input type=button value=\" \\ \"></td>\n            <td> </td>\n          </tr>\n          <tr align=\"left\" valign=\"middle\"> \n            <td> <input type=button value=\" q \"></td>\n            <td> <input type=button value=\" w \"></td>\n            <td> <input type=button value=\" e \"></td>\n            <td> <input type=button value=\" r \"></td>\n            <td> <input type=button value=\" t \"></td>\n            <td> <input type=button value=\" y \"></td>\n            <td> <input type=button value=\" u \"></td>\n            <td> <input type=button value=\" i \"></td>\n            <td> <input type=button value=\" o \"></td>\n            <td> <input name=\"button8\" type=button value=\" p \"></td>\n            <td> <input name=\"button9\" type=button value=\" { \"></td>\n            <td> <input type=button value=\" } \"></td>\n            <td> <input type=button value=\" [ \"></td>\n            <td> <input type=button value=\" ] \"></td>\n            <td><input name=\"button9\" type=button onClick=\"capsLockText();setCapsLock();\"  onDblClick=\"capsLockText();setCapsLock();\" value=\"切换大/小写\" style=\"width:100px;\"></td>\n          </tr>\n          <tr align=\"left\" valign=\"middle\"> \n            <td> <input type=button value=\" a \"></td>\n            <td> <input type=button value=\" s \"></td>\n            <td> <input type=button value=\" d \"></td>\n            <td> <input type=button value=\" f \"></td>\n            <td> <input type=button value=\" g \"></td>\n            <td> <input type=button value=\" h \"></td>\n            <td> <input type=button value=\" j \"></td>\n            <td> <input name=\"button3\" type=button value=\" k \"></td>\n            <td> <input name=\"button4\" type=button value=\" l \"></td>\n            <td> <input name=\"button5\" type=button value=\" : \"></td>\n            <td> <input name=\"button7\" type=button value=\" &quot; \"></td>\n            <td> <input type=button value=\" ; \"></td>\n            <td> <input type=button value=\" ' \"></td>\n            <td rowspan=\"2\" colspan=\"2\"> <input name=\"button12\" type=button onclick=\"OverInput();\" value=\"   确定  \" style=\"width:130px;height:42\"></td>\n          </tr>\n          <tr align=\"left\" valign=\"middle\"> \n            <td> <input name=\"button2\" type=button value=\" z \"></td>\n            <td> <input type=button value=\" x \"></td>\n            <td> <input type=button value=\" c \"></td>\n            <td> <input type=button value=\" v \"></td>\n            <td> <input type=button value=\" b \"></td>\n            <td> <input type=button value=\" n \"></td>\n            <td> <input type=button value=\" m \"></td>\n            <td> <input type=button value=\" &lt; \"></td>\n            <td> <input type=button value=\" &gt; \"></td>\n            <td> <input type=button value=\" ? \"></td>\n            <td> <input type=button value=\" , \"></td>\n            <td> <input type=button value=\" . \"></td>\n            <td> <input type=button value=\" / \"></td>\n          </tr>\n        </table></td>    </FORM>      </tr>  </table></DIV>")
//给输入的密码框添加新值
	function addValue(newValue)
	{
		if (CapsLockValue==0)
		{
			var str=Calc.password.value;
			if(str.length<password1.maxLength)
			{
				Calc.password.value += newValue;
			}			
			if(str.length<=password1.maxLength)
			{
				password1.value=Calc.password.value;
			}
		}
		else
		{
			var str=Calc.password.value;
			if(str.length<password1.maxLength)
			{
				Calc.password.value += newValue.toUpperCase();
			}
			if(str.length<=password1.maxLength)
			{
				password1.value=Calc.password.value;
			}
		}
	}
//实现BackSpace键的功能
	function setpassvalue()
	{
		var longnum=Calc.password.value.length;
		var num
		num=Calc.password.value.substr(0,longnum-1);
		Calc.password.value=num;
		var str=Calc.password.value;
			password1.value=Calc.password.value;
	}
//输入完毕
	function OverInput()
	{
		//m_pass.mempass.value=Calc.password.value;
		var str=Calc.password.value;
			password1.value=Calc.password.value;
			//alert(theForm.value);
		//theForm.value=m_pass.mempass.value;
		softkeyboard.style.display="none";
		Calc.password.value="";
		password1.readOnly=1;
		//password1.value=Calc.password.value;
	}
//关闭软键盘
	function closekeyboard(theForm)
	{
		//eval("var theForm="+theForm+";");
		//theForm.value="";
		softkeyboard.style.display="none";
		//Calc.password.value="";

	}
//显示软键盘
	function showkeyboard()
	{
		if(event.y+140)
		softkeyboard.style.top=event.y+document.body.scrollTop+15;
		if((event.x-250)>0)
		{
			softkeyboard.style.left=event.x-250;
		}
		else
		{
			softkeyboard.style.left=0;
		}
		
		softkeyboard.style.display="block";
		password1.readOnly=1;
		password1.blur();
		//password1.value="";
	}

//设置是否大写的值
function setCapsLock()
{
	if (CapsLockValue==0)
	{
		CapsLockValue=1
//		Calc.showCapsLockValue.value="当前是大写 ";
	}
	else 
	{
		CapsLockValue=0
//		Calc.showCapsLockValue.value="当前是小写 ";
	}
}


function setCalcborder()
{
	CalcTable.style.border="1px solid #0090FD"
}

function setHead()
{
	CalcTable.cells[0].style.backgroundColor="#7EDEFF"	
}

function setCalcButtonBg()
{
	for(var i=0;i<Calc.elements.length;i++)
	{
		if(Calc.elements[i].type=="button"&&Calc.elements[i].bgtype!="1")
		{
	//		if(i==10)
//	alert(123);
			Calc.elements[i].style.borderTopWidth= 0
			Calc.elements[i].style.borderRightWidth= 2
			Calc.elements[i].style.borderBottomWidth= 2
			Calc.elements[i].style.borderLeftWidth= 0
			Calc.elements[i].style.borderTopStyle= "none";
			Calc.elements[i].style.borderRightStyle= "solid";
			Calc.elements[i].style.borderBottomStyle= "solid";
			Calc.elements[i].style.borderLeftStyle= "none";
			//#46AC17
			Calc.elements[i].style.borderTopColor= "#118ACC";
			Calc.elements[i].style.borderRightColor= "#118ACC";
			Calc.elements[i].style.borderBottomColor= "#118ACC";
			Calc.elements[i].style.borderLeftColor= "#118ACC";
			//#CBF3B2
			Calc.elements[i].style.backgroundColor="#ADDEF8";

			
			
			var str1=Calc.elements[i].value;
			str1=str1.trim();
			/*
			if(str1=="`") 
			{
				Calc.elements[i].style.fontSize=14;
			}
			*/

			if(str1.length==1)
			{
				//Calc.elements[i].style.fontSize=16;
				//Calc.elements[i].style.fontWeight='bold';
			}
			
			var thisButtonValue=Calc.elements[i].value;
			thisButtonValue=thisButtonValue.trim();
			if(thisButtonValue.length==1)
			{
				Calc.elements[i].onclick=
					function ()
					{
						var thisButtonValue=this.value;
						thisButtonValue=thisButtonValue.trim();
						addValue(thisButtonValue);
						//alert(234)
					}
				Calc.elements[i].ondblclick=
					function ()
					{
						var thisButtonValue=this.value;
						thisButtonValue=thisButtonValue.trim();
						addValue(thisButtonValue);
						//alert(234)
					}
			}
			
		}
	}
}

function initCalc()
{
	setCalcborder();
	setHead();
	setCalcButtonBg();
}

String.prototype.trim = function()
{
    // 用正则表达式将前后空格
    // 用空字符串替代。
    return this.replace(/(^\s*)|(\s*$)/g, "");
}

var capsLockFlag;
capsLockFlag=true;

function capsLockText()
{
if(capsLockFlag)//改成大写
{
	for(var i=0;i<Calc.elements.length;i++)
	{
			var char=Calc.elements[i].value;
			var char=char.trim()
		if(Calc.elements[i].type=="button"&&char>="a"&&char<="z"&&char.length==1)
		{
		
			Calc.elements[i].value=" "+String.fromCharCode(char.charCodeAt(0)-32)+" "
		}
	}
}
else
{
	for(var i=0;i<Calc.elements.length;i++)
	{
			var char=Calc.elements[i].value;
			var char=char.trim()
		if(Calc.elements[i].type=="button"&&char>="A"&&char<="Z"&&char.length==1)
		{
		
			Calc.elements[i].value=" "+String.fromCharCode(char.charCodeAt(0)+32)+" "
		}
	}
}
capsLockFlag=!capsLockFlag;
}

window.onload=
	function ()
	{
		password1=null;		
		initCalc();
		

	}
//-->
</SCRIPT>


</BODY></HTML>
