function winload()
{
huashuolayer2.style.top=20;
huashuolayer2.style.left=5;
huashuolayer3.style.top=20;
huashuolayer3.style.right=5;
}

if(document.body.offsetWidth>800){	
	{
	document.write("<div id=huashuolayer2 style='position: absolute;visibility:visible;z-index:1'><EMBED src='banner/ck.swf' quality=high  WIDTH=100 HEIGHT=240 TYPE='application/x-shockwave-flash' id=sinadl wmode=opaque></EMBED></div>"
	              +"<div id=huashuolayer3 style='position: absolute;visibility:visible;z-index:1'><EMBED src='banner/ck1.swf' quality=high  WIDTH=100 HEIGHT=240 TYPE='application/x-shockwave-flash' id=sinadl wmode=opaque></EMBED></div>");
	}
  winload()
}