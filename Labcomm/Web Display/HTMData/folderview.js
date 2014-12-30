function showDir(dirName,showBack)
{
	objFS.FolderLocation=dirName;
	var sf = new Enumerator(objFS.SubFolders);
	
	var imgSrc = "";
	var html = '<table id="bp">\n';
	var row1='';
	var colId = 0;
	html += '<tr><td class="foldertitle" colspan="6">' + document.title + '</td></tr>';

	if(showBack=="YES")
	{
		imgSrc = '<img action="over" height="32" width="32" alt="Click to open" border="0" src="C:\\ICE\\LABCOMM\\Icons\\back.ico">';
		row1 = '	<td onclick="handleClick" class="folder">' + imgSrc + '<br><span action="over" class="fixwidth">Up</span></td>';
		colId++;
	};
	
	imgSrc = '<img action="over" height="32" width="32" alt="Click to open" border="0" src="C:\\ICE\\LABCOMM\\Icons\\folder.ico">';
	
	while (!sf.atEnd())
	{
		row1 += '	<td onclick="handleClick" class="folder">' + imgSrc + '<br><span action="over" class="fixwidth" title="'+ sf.item(0).Name + '">' + sf.item(0).Name + '</span></td>';
		colId++;

		if(colId==6)
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
