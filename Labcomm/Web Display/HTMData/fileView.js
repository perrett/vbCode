function showDir(dirName,showBack)
{
	objFS.FolderLocation=dirName;
	var sf = new Enumerator(objFS.Files);
	
	var imgSrc = "";
	var html = '<table id="bp">\n';
	var row1='';
	var colId = 0;

	html += '<tr><td class="foldertitle" colspan="6">' + document.title + '</td></tr>';
	if(showBack!="")
	{
		imgSrc = '<img action="over" height="32" width="32" alt="Click to open" border="0" src="C:\\ICE\\LABCOMM\\Icons\\back.ico">';
		row1 = '	<td onclick="handleClick" class="folder">' + imgSrc + '<br><span action="over" class="fixwidth">' + showBack + '</span></td>';
		colId++;
	};

	imgSrc = '<img action="over" height="32" width="32" alt="Click to open" border="0" src="C:\\ICE\\LABCOMM\\Icons\\';
	var fileImg = "";
	
	while (!sf.atEnd())
	{
		switch(sf.item(0).Type.substr(0,3))
		{
			case 'XMS':
				fileImg = imgSrc + 'xms.ico">';
				break;
				
			case "XEN":
				fileImg = imgSrc + 'xen.ico">';
				break;
				
			case "WNG":
				fileImg = imgSrc + 'warning.ico">';
				break;
				
			case "ERR":
				fileImg = imgSrc + 'error.ico">';
				break;
				
			default:
				fileImg = imgSrc + 'unknown.ico">';
				break;
		}
			
		row1 += '	<td onclick="handleClick" class="folder">' + fileImg + '<br><span action="over" class="fixwidth" title="'+ sf.item(0).Name + '">' + sf.item(0).Name.replace(/-/g,"_") + '</span></td>';
		colId++;

		if(colId==5)
		{
			html += '<tr>' + row1 + '</tr>\n';
			row1='';
			colId=0;
		};
		sf.moveNext();
	};
	
	if(colId>0)
	{
		html += '<tr>' + row1 + '</tr>\n';
	};
	
	html += '</table>\n';
	
	return html;	
}
