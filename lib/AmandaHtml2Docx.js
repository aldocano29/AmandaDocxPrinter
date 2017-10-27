var html2docx = {
	convertContent: function(input)
	{ 
		var output,
			inputDoc,
			id,
			doc,
			inNode,
			outNode,
			styleAttrNode,
			pCount = 0,
			tempStr,
			tempNode,
			val;

		var node = document.createElement('span');
		node.innerHTML = input;
		
		input = node;

		function newXMLnode(nodeName, text)
		{
			var name = nodeName;
			if (name.indexOf(":") < 0)
			{
				name = 'w:' + name;
			}
			var el = doc.createElement(name);
			if (text)
			{
				el.appendChild(doc.createTextNode(text));
			}
			return el;
		}

		function newHTMLnode(name, html)
		{
			var el = document.createElement(name);
			el.innerHTML = html || '';
			return el;
		}

		function color(str)
		{ 
			if (str.charAt(0) === '#')
			{
				return str.substr(1);
			}
			if (str.indexOf('rgb') < 0)
			{
				return str;
			}
			var values = (/rgb\((\d+), (\d+), (\d+)\)/).exec(str),
				red = +values[1],
				green = +values[2],
				blue = +values[3];
			return (blue | (green << 8) | (red << 16)).toString(16);
		}

		function paragraph(inNode)
		{ 
			var outNode = newHTMLnode('P');
			tempStr = '';
			for (var k = 0; k < inNode.childNodes.length; k++)
			{
				var inNodeChild = inNode.childNodes[k];
				if (inNodeChild.nodeName === 'pPr')
				{
					if (styleAttrNode = inNodeChild.getElementsByTagName('jc')[0])
					{
						outNode.style.textAlign = styleAttrNode.getAttribute('w:val');
					}
					if (styleAttrNode = inNodeChild.getElementsByTagName('pBdr')[0])
					{
						setBorders(outNode, styleAttrNode);
					}
				}
				if (inNodeChild.nodeName === 'r')
				{
					val = inNodeChild.textContent;
					if (inNodeChild.getElementsByTagName('b').length)
					{
						val = '<b>' + val + '</b>';
					}
					if (inNodeChild.getElementsByTagName('i').length)
					{
						val = '<i>' + val + '</i>';
					}
					if (inNodeChild.getElementsByTagName('u').length)
					{
						val = '<u>' + val + '</u>';
					}
					if (inNodeChild.getElementsByTagName('strike').length)
					{
						val = '<s>' + val + '</s>';
					}
					if (styleAttrNode = inNodeChild.getElementsByTagName('vertAlign')[0])
					{
						if (styleAttrNode.getAttribute('w:val') === 'subscript')
						{
							val = '<sub>' + val + '</sub>';
						}
						if (styleAttrNode.getAttribute('w:val') === 'superscript')
						{
							val = '<sup>' + val + '</sup>';
						}
					}
					if (styleAttrNode = inNodeChild.getElementsByTagName('sz')[0])
					{
						val = '<span style="font-size:' + (styleAttrNode.getAttribute('w:val') / 2) + 'pt">' + val + '</span>';
					}
					if (styleAttrNode = inNodeChild.getElementsByTagName('highlight')[0])
					{
						val = '<span style="background-color:' + styleAttrNode.getAttribute('w:val') + '">' + val + '</span>';
					}
					if (styleAttrNode = inNodeChild.getElementsByTagName('color')[0])
					{
						val = '<span style="color:#' + styleAttrNode.getAttribute('w:val') + '">' + val + '</span>';
					}
					
					tempStr += val;
				}
				outNode.innerHTML = tempStr;
				if (outNode.innerHTML === "")
				{
					outNode.innerHTML = "&nbsp;";
				}
			}
			return outNode;
		}

		function setBorders(htmlNode, bNode)
		{
			for (var bsp = 0; bsp < bNode.childNodes.length; bsp++)
			{
				if (bNode.childNodes[bsp].nodeName === 'top')
				{
					htmlNode.style.borderTopWidth = bNode.childNodes[bsp].getAttribute('w:sz') / 4 + "pt";
					htmlNode.style.borderTopStyle = "solid";
					htmlNode.style.borderTopColor = "black";
				}
				if (bNode.childNodes[bsp].nodeName === 'bottom')
				{
					htmlNode.style.borderBottomWidth = bNode.childNodes[bsp].getAttribute('w:sz') / 4 + "pt";
					htmlNode.style.borderBottomStyle = "solid";
					htmlNode.style.borderBottomColor = "black";
				}
				if (bNode.childNodes[bsp].nodeName === 'left')
				{
					htmlNode.style.borderLeftWidth = bNode.childNodes[bsp].getAttribute('w:sz') / 4 + "pt";
					htmlNode.style.borderLeftStyle = "solid";
					htmlNode.style.borderLeftColor = "black";
				}
				if (bNode.childNodes[bsp].nodeName === 'right')
				{
					htmlNode.style.borderRightWidth = bNode.childNodes[bsp].getAttribute('w:sz') / 4 + "pt";
					htmlNode.style.borderRightStyle = "solid";
					htmlNode.style.borderRightColor = "black";
				}
				if (bNode.childNodes[bsp].nodeName === 'insideH' || bNode.childNodes[bsp].nodeName === 'insideW')
				{
					htmlNode.style.borderCollapse = "collapse";
				}
			}
		}

		function table(inNode)
		{ 
			var outNode = newHTMLnode('TABLE');
			tempStr = '';
			for (var j = 0; j < inNode.childNodes.length; j++)
			{
				var tableChild = inNode.childNodes[j];
				if (tableChild.nodeName === 'tblPr')
				{
					var tableProperties = tableChild;
					for (var i = 0; i < tableProperties.childNodes.length; i++)
					{
						var prop = tableProperties.childNodes[i];
						if (prop.nodeName === 'tblBorders')
						{
							setBorders(outNode, prop);
						}
						if (prop.nodeName === 'tblW')
						{
							if (prop.getAttribute('w:type') === 'dxa')
							{
								outNode.style.width = prop.getAttribute('w:w') / 12 + "px";
							}
						}
					}
					if (styleAttrNode = tableChild.getElementsByTagName('jc')[0])
					{
						outNode.style.textAlign = styleAttrNode.getAttribute('w:val');
					}
				}
				if (tableChild.nodeName === 'tr')
				{
					var trNode = newHTMLnode('TR');
					for (var c = 0; c < tableChild.childNodes.length; c++)
					{
						var cell = tableChild.childNodes[c];
						var tdNode = newHTMLnode('TD');
						for (var cc = 0; cc < cell.childNodes.length; cc++)
						{
							var cellChild = cell.childNodes[cc];
							if (cellChild.nodeName === 'tcPr')
							{
								var cellProperties = cellChild;
								for (var k = 0; k < cellProperties.childNodes.length; k++)
								{
									var trProp = cellProperties.childNodes[k];
									if (trProp.nodeName === 'tcBorders')
									{
										setBorders(tdNode, trProp);
									}
									if (trProp.nodeName === 'gridSpan')
									{
										tdNode.colSpan = trProp.getAttribute("w:val");
									}
								}
							}
							if (cellChild.nodeName === 'p')
							{
								var p = paragraph(cellChild);
								tdNode.appendChild(p);
							}
						}
						trNode.appendChild(tdNode);
					}
					outNode.appendChild(trNode);
				}
			}
			return outNode;
		}

		function toXML(str)
		{
			return new DOMParser().parseFromString(str.replace(/<[a-zA-Z]*?:/g, '<').replace(/<\/[a-zA-Z]*?:/g, '</'), 'text/xml').firstChild;
		}
		if (input.nodeName)
		{ 
			doc = new DOMParser().parseFromString('<root></root>', 'text/xml');
			doc.getElementsByTagName('root')[0].appendChild(newXMLnode('body'));
			output = doc.getElementsByTagName('w:body')[0];
			var numberOfLists = 0;
			var linkData = [];
			var riditer = 6;
			var handleInnerNode = function(inNodeChild, outNode)
			{
				if (inNodeChild.nodeName === 'P' || inNodeChild.nodeName === 'BLOCKQUOTE')
				{
					for (var h = 0; h < inNodeChild.childNodes.length; h++)
					{
						handleInnerNode(inNodeChild.childNodes[h], outNode);
					}
				}
				else
				{
					var outNodeChild;
					if (inNodeChild.nodeName !== '#text' || inNodeChild.parentNode.nodeName === 'CODE')
					{
						tempStr = inNodeChild.outerHTML;
						if (inNodeChild.nodeName === 'A')
						{
							riditer++;
							var linkrid = "rId" + riditer;
							var hyNode = outNode.appendChild(newXMLnode('hyperlink'));
							var tempHref = inNodeChild.getAttribute('href');
							hyNode.setAttribute("r:id", linkrid);
							hyNode.setAttribute("w:history", "1");
							var tempTitle = inNodeChild.getAttribute('title');
							if (tempTitle)
							{
								hyNode.setAttribute("w:tooltip", tempTitle);
							}
							else
							{
								hyNode.setAttribute("w:tooltip", tempHref);
							}
							linkData.push(
							{
								href: tempHref,
								rid: linkrid
							});
							outNodeChild = hyNode.appendChild(newXMLnode('r'));
							
						}
						else if (inNodeChild.nodeName === 'TH' || inNodeChild.nodeName === 'TD')
						{
							var tcNode = outNode.appendChild(newXMLnode('tc'));
							var tcPrNode = tcNode.appendChild(newXMLnode('tcPr'));
							tcPrNode.appendChild(newXMLnode('tcW')).setAttribute("w:type", "dxa");
							var tcPNode = tcNode.appendChild(newXMLnode('p'));
							outNodeChild = tcPNode.appendChild(newXMLnode('r'));
							pCount++;
						}
						else
						{
							outNodeChild = outNode.appendChild(newXMLnode('r'));
						}
						
						styleAttrNode = outNodeChild.appendChild(newXMLnode('rPr'));
						if (inNodeChild.parentNode.nodeName === 'CODE' || (tempStr && tempStr.indexOf('<code>') > -1))
						{
							var fontNode = styleAttrNode.appendChild(newXMLnode('rFonts'));
							fontNode.setAttribute('w:ascii', "Courier");
							fontNode.setAttribute('w:hAnsi', "Courier");
							var shadeNode = styleAttrNode.appendChild(newXMLnode('shd'));
							shadeNode.setAttribute('w:color', "auto");
							shadeNode.setAttribute('w:fill', "EEEEEE");
							shadeNode.setAttribute('val', "clear");
						}
						if (tempStr)
						{
							if (tempStr.indexOf('<b>') > -1)
							{
								styleAttrNode.appendChild(newXMLnode('b'));
							}
							if (tempStr.indexOf('<a ') > -1)
							{
								styleAttrNode.appendChild(newXMLnode('rStyle')).setAttribute('val', 'Hyperlink');
							}
							if (tempStr.indexOf('<strong>') > -1)
							{
								styleAttrNode.appendChild(newXMLnode('b'));
							}
							if (tempStr.indexOf('<em>') > -1)
							{
								styleAttrNode.appendChild(newXMLnode('i'));
							}
							if (tempStr.indexOf('<i>') > -1)
							{
								styleAttrNode.appendChild(newXMLnode('i'));
							}
							if (tempStr.indexOf('<u>') > -1)
							{
								styleAttrNode.appendChild(newXMLnode('u')).setAttribute('val', 'single');
							}
							if (tempStr.indexOf('<s>') > -1)
							{
								styleAttrNode.appendChild(newXMLnode('strike'));
							}
							if (tempStr.indexOf('<del>') > -1)
							{
								styleAttrNode.appendChild(newXMLnode('strike'));
							}
							if (tempStr.indexOf('<sub>') > -1)
							{
								styleAttrNode.appendChild(newXMLnode('vertAlign')).setAttribute('val', 'subscript');
							}
							if (tempStr.indexOf('<sup>') > -1)
							{
								styleAttrNode.appendChild(newXMLnode('vertAlign')).setAttribute('val', 'superscript');
							}
							if (tempNode = inNodeChild.nodeName === 'SPAN' ? inNodeChild : inNodeChild.getElementsByTagName('SPAN')[0])
							{
								if (tempNode.style.fontSize)
								{
									styleAttrNode.appendChild(newXMLnode('sz')).setAttribute('val', parseInt(tempNode.style.fontSize, 10) * 2);
								}
								else if (tempNode.style.backgroundColor)
								{
									styleAttrNode.appendChild(newXMLnode('highlight')).setAttribute('val', color(tempNode.style.backgroundColor));
								}
								else if (tempNode.style.color)
								{
									styleAttrNode.appendChild(newXMLnode('color')).setAttribute('val', color(tempNode.style.color));
								}
							}
						}
					}
					else
					{
						outNodeChild = outNode.appendChild(newXMLnode('r'));
					}
					if (inNodeChild.nodeName === 'BR')
					{
						outNodeChild.appendChild(newXMLnode('br', inNodeChild.textContent));
					}
					else
					{
						outNodeChild.appendChild(newXMLnode('t', inNodeChild.textContent));
					}
				}
			};
			var convertNode = function(inNode, output, listNumber, listDepthNumber)
			{
				
				var lists = listNumber || 1;
				var listDepth = listDepthNumber || 0;
				var newPNode = function()
				{
					outNode = output.appendChild(newXMLnode('p'));
					pCount++;
					return outNode;
				};
				var nName = inNode.nodeName;
				
				if (nName === '#text')
				{}
				else
				{
					var isList = false;
					
					if (nName === 'OL' || nName === 'UL')
					{
						isList = nName;
						if (!listNumber)
						{
							numberOfLists++;
						}
						var tempNum = listNumber || numberOfLists;
						for (var t = 0; t < inNode.children.length; t++)
						{
							var inNodeChild = inNode.children[t];
							if (inNodeChild)
							{
								convertNode(inNodeChild, output, tempNum, listDepth);
							}
						}
						
					}
					else if (nName === 'TABLE')
					{
						var tblNode = output.appendChild(newXMLnode('tbl'));
						var tblPrNode = tblNode.appendChild(newXMLnode('tblPr'));
						var tblStyleNode = tblPrNode.appendChild(newXMLnode('tblStyle'));
						tblStyleNode.setAttribute('val', "TableGrid");
						var tblWNode = tblPrNode.appendChild(newXMLnode('tblW'));
						tblWNode.setAttribute('w:type', 'auto');
						tblWNode.setAttribute('w:w', '0');
						var tblLookNode = tblPrNode.appendChild(newXMLnode('tblLook'));
						tblLookNode.setAttribute('w:firstColumn', '1');
						tblLookNode.setAttribute('w:firstRow', '1');
						tblLookNode.setAttribute('w:lastColumn', '0');
						tblLookNode.setAttribute('w:lastRow', '0');
						tblLookNode.setAttribute('w:noHBand', '0');
						tblLookNode.setAttribute('w:noVBand', '1');
						tblLookNode.setAttribute('val', '04A0');
						for (var d = 0; d < inNode.children.length; d++)
						{
							var inNodeTableChild = inNode.children[d];
							if (inNodeTableChild)
							{
								convertNode(inNodeTableChild, tblNode, lists, listDepth);
							}
						}
						
					}
					else if (nName === 'THEAD')
					{
						var tblGridNode = output.appendChild(newXMLnode('tblGrid'));
						for (var v = 0; v < inNode.firstElementChild.children.length; v++)
						{
							tblGridNode.appendChild(newXMLnode('gridCol'));
						}
						if (inNode.firstElementChild)
						{
							convertNode(inNode.firstElementChild, output, lists, listDepth);
						}
						
					}
					else if (nName === 'TBODY')
					{
						for (var w = 0; w < inNode.children.length; w++)
						{
							if (inNode.children[w])
							{
								convertNode(inNode.children[w], output, lists, listDepth);
							}
						}
						
					}
					else
					{
						if (nName === 'TR')
						{
							outNode = output.appendChild(newXMLnode('tr'));
						}
						else
						{
							outNode = newPNode();
						}
						
						if (nName === 'PRE')
						{
							var node = inNode.firstElementChild.childNodes[0];
							var words = node.textContent.split(/\r\n|\r|\n/g);
							var parent = node.parentNode;
							for (var J = 0; J < words.length; ++J)
							{
								var newWord = document.createTextNode(words[J]);
								parent.insertBefore(newWord, node);
								if (J < (words.length - 1))
								{
									var newBreak = document.createElement("br");
									parent.insertBefore(newBreak, node);
								}
							}
							parent.removeChild(node);

							inNode = inNode.firstElementChild;
							var codeNode = outNode.appendChild(newXMLnode('pPr'));
							var shadeNode = codeNode.appendChild(newXMLnode('shd'));
							shadeNode.setAttribute('w:color', "auto");
							shadeNode.setAttribute('w:fill', "EEEEEE");
							shadeNode.setAttribute('val', "clear");
						}

						if (inNode.style && inNode.style.textAlign)
						{
							outNode.appendChild(newXMLnode('pPr')).appendChild(newXMLnode('jc')).setAttribute('val', inNode.style.textAlign);
						}

						if (nName.length == 2 && nName[0] == 'H' && !isNaN(nName[1]))
						{
							outNode.appendChild(newXMLnode('pPr')).appendChild(newXMLnode('pStyle')).setAttribute('val', 'Heading' + nName[1]);
						}
						else if (nName === 'BLOCKQUOTE')
						{
							outNode.appendChild(newXMLnode('pPr')).appendChild(newXMLnode('pStyle')).setAttribute('val', 'Quote');
						}

						if (nName === "LI")
						{
							var tempOutNodeChild = outNode.appendChild(newXMLnode('pPr'));
							tempOutNodeChild.appendChild(newXMLnode('pStyle')).setAttribute('val', "ListParagraph");
							var childChild = tempOutNodeChild.appendChild(newXMLnode('numPr'));
							childChild.appendChild(newXMLnode('ilvl')).setAttribute('val', listDepth);
							childChild.appendChild(newXMLnode('numId')).setAttribute('val', lists.toString());
						}
						for (var j = 0; j < inNode.childNodes.length; j++)
						{
							var inNodeChild = inNode.childNodes[j];

							if (inNodeChild.nodeName === 'OL' || inNodeChild.nodeName === 'UL')
							{
								var newlistDepth = listDepth + 1;
								convertNode(inNodeChild, output, lists, newlistDepth);

							}
							else if (inNodeChild.nodeName === '#text' && inNodeChild.nodeValue.length == 1 && inNodeChild.nodeValue.charCodeAt(0) == 10)
							{

							}
							else
							{
								handleInnerNode(inNodeChild, outNode);
							}
						}
					}
				}
			};
			for (var m = 0; m < input.children.length; m++)
			{
				convertNode(input.children[m], output);
			}
			output = {
				string: new XMLSerializer().serializeToString(output).replace(/<w:t>/g, '<w:t xml:space="preserve">').replace(/val=/g, 'w:val=').replace('<w:body>','').replace('</w:body>',''),
				charSpaceCount: input.textContent.length,
				charCount: input.textContent.replace(/\s/g, '').length,
				pCount: pCount,
				linkData: linkData
			};
		}
		return output;
	}
}