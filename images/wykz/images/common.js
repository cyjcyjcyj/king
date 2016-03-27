<!--
function SetCookie(sName, sValue,iExpireDays)
{
	if (iExpireDays)
	{
		var dExpire = new Date();
		dExpire.setTime(dExpire.getTime()+parseInt(iExpireDays*24*60*60*1000));
		document.cookie = sName + "=" + escape(sValue) + "; expires=" + dExpire.toGMTString();
	}
	else
	{
		document.cookie = sName + "=" + escape(sValue);
	}
}

function GetCookie(sName)
{
	var arr = document.cookie.match(new RegExp("(^| )"+sName+"=([^;]*)(;|$)"));
	if(arr !=null){return unescape(arr[2])};
	return null;

}
function $(_sId){return document.getElementById(_sId);}
//-->