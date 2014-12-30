var curClass;
var curCol;
function evTest()
{
	var ev = event.srcElement
	if (ev.nodeName=="TD")
	{
		curCol=ev.Color;
		ev=event.srcElement.parentNode.parentNode.parentNode;
		if (ev.id.substring(0,3)=="HL_")
		{
			hilite(event.srcElement.parentNode);
			if(document.getElementById("SH_"+ev.id.substring(3)))
			{
				hilite(eval("SH_"+ev.id.substring(3)))
			}
		}
	}
}
function hilite(obj)
{
	if(obj.className.substr(0,9)=="highlight")
	{
		obj.className=curClass
	}
	else
	{
		if(obj.nodeName=="TR")
		{
			curClass=obj.className;
		};
//		curClass=obj.className;
		if (curClass=="oor")
		{
			obj.className="highlightRed"
		}
		else
		{
			obj.className="highlightBlack"
		}
	}
}
document.onmouseover = evTest;
document.onmouseout = evTest;