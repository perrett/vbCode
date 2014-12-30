function handleEvent()
{
	var src = event.srcElement;
	var eType = event.type;
	
	if (src.attributes["action"])
	{
		src = src.parentNode;

		switch(eType)
		{
			case "mouseover":
				handleOver(src);
				break;
				
			case "mouseout":
				handleOut(src);
				break;
				
			case "click":
				event.cancelBubble=true;
				event.returnValue=false;
				handleClick(src);
				break;			
		}
	}
}

function handleClick(src)
{
	src.fireEvent("ondataavailable");
}

function handleOver(src)
{
	src.style.backgroundColor="#FFFF99";
	src.style.color="blue";
	src.style.borderTop="1px solid #E0E0E0";
	src.style.borderLeft="2px solid #E0E0E0";
	src.style.borderBottom="3px solid #4D4D4D";
	src.style.borderRight="3px solid #4D4D4D";
}

function handleOut(src)
{
	src.style.backgroundColor="white";
	src.style.color="black";
	src.style.borderTop="1px solid white";
	src.style.borderLeft="2px solid white";
	src.style.borderBottom="3px solid white";
	src.style.borderRight="3px solid white";
}

function handleInput()
{
	var title = document.title;
	var showBack="";
	
	if(document.body.attributes["param"])
	{
		showBack=document.body.attributes["param"].nodeValue;
	};
	
	document.write(showDir(document.titleshowBack));
	document.title=title;
	document.close;
}

document.ondatasetchanged=handleInput;
document.onclick=handleEvent;
document.onmouseover=handleEvent;
document.onmouseout=handleEvent;
