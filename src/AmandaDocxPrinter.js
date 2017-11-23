//Global Variables that will be used along the whole process
var docxUrl = null;
var docxOutName = null;
var procSwitch = null;
var dataSources = null;
var texts = null;
var docxtemplater = null;
var docxOut = null;
var querycursor = null;
var trigger = null;
var templateValidatorRes = null;
var dataSourceValidatorRes = null;
var dataSourceItem = null;

var amanda_docx_printer = {

  //Switch Method that will check the type of Data the user wants to pass to the Template
  docxSelector : function(dynamicAction){

    //The global variables are populated using the Dynamic Action attributes
    console.log("AmandaDocxPrinter: docxSelector Started");
    docxUrl = dynamicAction.attribute01;
    procSwitch = dynamicAction.attribute02;
    docxOutName = dynamicAction.attribute03;
    querycursor = dynamicAction.attribute04;
    templateValidatorRes = dynamicAction.attribute05;
    dataSourceValidatorRes = dynamicAction.attribute06;
    dataSourceItem = dynamicAction.attribute07;



    console.log("AmandaDocxPrinter: Dynamic Action: "+dynamicAction);
    console.log("AmandaDocxPrinter: docxUrl value: "+ docxUrl);
    console.log("AmandaDocxPrinter: docxOutName value: "+docxOutName);
    console.log("AmandaDocxPrinter: procSwitch value: "+procSwitch);
    console.log("AmandaDocxPrinter: dataSources value: "+dataSources);
    console.log("AmandaDocxPrinter: Template Validator Result value: "+templateValidatorRes);
    console.log("AmandaDocxPrinter: DataSource Validator Result value: "+dataSourceValidatorRes);
    console.log("AmandaDocxPrinter: DataSource Validator ITEM(s): "+dataSourceItem);
    console.log("QUERIES: ["+querycursor+"]");



    switch(procSwitch){
      case "REPLACEVARIABLES":
        console.log("Enters REPLACEVARIABLES Switch");
        amanda_docx_printer.docxLoader(amanda_docx_printer.docxAjaxCaller);
        break;

      case "VALIDATETEMPLATE":
        console.log("Enters VALIDATETEMPLATE Switch");
        amanda_docx_printer.docxLoader(amanda_docx_printer.docxTemplateValidator);
        break;



      case "VALIDATEDATASOURCE":
        console.log("Enters VALIDATEDATASOURCE Switch");
        amanda_docx_printer.docxDataSourceValidator();
        break;


      default:
        null;
    }

  },

  docxLoader : function(loaded){

    //This Method, loads the Template for the docx we will work on...
    JSZipUtils.getBinaryContent(docxUrl, function(error,content){
        if(error){
            console.log("An error occurred during the docxLoader Function: ["+error+"]");
            return;
        }

        //The content of the Template is loaded within the ZipContent Variable
        //The content of the xml data of the template will be loaded within the documentxml Variable
        //The content of the documentxml variable will be parsed as plain text within the strDocumentxml
        var zipContent = new JSZip(content);
        var documentxml = zipContent.file("word/document.xml");
		zipContent.file("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml" /><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/><Override PartName="/word/webSettings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"/><Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/><Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>');
		zipContent.file("word/numbering.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://purl.oclc.org/ooxml/officeDocument/relationships" xmlns:m="http://purl.oclc.org/ooxml/officeDocument/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:w16se="http://schemas.microsoft.com/office/word/2015/wordml/symex" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" mc:Ignorable="w14 w15 w16se wne wp14"><w:abstractNum w:abstractNumId="0" w15:restartNumberingAfterBreak="0"><w:nsid w:val="3613289E"/><w:multiLevelType w:val="multilevel"/><w:tmpl w:val="D7B8295A"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val=""/><w:lvlJc w:val="start"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="36pt"/></w:tabs><w:ind w:start="36pt" w:hanging="36pt"/></w:pPr><w:rPr><w:rFonts w:ascii="Symbol" w:hAnsi="Symbol" w:hint="default"/></w:rPr></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%2."/><w:lvlJc w:val="start"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="72pt"/></w:tabs><w:ind w:start="72pt" w:hanging="36pt"/></w:pPr></w:lvl><w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%3."/><w:lvlJc w:val="start"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="108pt"/></w:tabs><w:ind w:start="108pt" w:hanging="36pt"/></w:pPr></w:lvl><w:lvl w:ilvl="3"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%4."/><w:lvlJc w:val="start"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="144pt"/></w:tabs><w:ind w:start="144pt" w:hanging="36pt"/></w:pPr></w:lvl><w:lvl w:ilvl="4"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%5."/><w:lvlJc w:val="start"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="180pt"/></w:tabs><w:ind w:start="180pt" w:hanging="36pt"/></w:pPr></w:lvl><w:lvl w:ilvl="5"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%6."/><w:lvlJc w:val="start"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="216pt"/></w:tabs><w:ind w:start="216pt" w:hanging="36pt"/></w:pPr></w:lvl><w:lvl w:ilvl="6"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%7."/><w:lvlJc w:val="start"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="252pt"/></w:tabs><w:ind w:start="252pt" w:hanging="36pt"/></w:pPr></w:lvl><w:lvl w:ilvl="7"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%8."/><w:lvlJc w:val="start"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="288pt"/></w:tabs><w:ind w:start="288pt" w:hanging="36pt"/></w:pPr></w:lvl><w:lvl w:ilvl="8"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%9."/><w:lvlJc w:val="start"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="324pt"/></w:tabs><w:ind w:start="324pt" w:hanging="36pt"/></w:pPr></w:lvl></w:abstractNum><w:abstractNum w:abstractNumId="1" w15:restartNumberingAfterBreak="0"><w:nsid w:val="593170C5"/><w:multiLevelType w:val="multilevel"/><w:tmpl w:val="2130AC26"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="start"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="36pt"/></w:tabs><w:ind w:start="36pt" w:hanging="36pt"/></w:pPr></w:lvl><w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%2."/><w:lvlJc w:val="start"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="72pt"/></w:tabs><w:ind w:start="72pt" w:hanging="36pt"/></w:pPr></w:lvl><w:lvl w:ilvl="2"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%3."/><w:lvlJc w:val="start"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="108pt"/></w:tabs><w:ind w:start="108pt" w:hanging="36pt"/></w:pPr></w:lvl><w:lvl w:ilvl="3"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%4."/><w:lvlJc w:val="start"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="144pt"/></w:tabs><w:ind w:start="144pt" w:hanging="36pt"/></w:pPr></w:lvl><w:lvl w:ilvl="4"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%5."/><w:lvlJc w:val="start"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="180pt"/></w:tabs><w:ind w:start="180pt" w:hanging="36pt"/></w:pPr></w:lvl><w:lvl w:ilvl="5"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%6."/><w:lvlJc w:val="start"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="216pt"/></w:tabs><w:ind w:start="216pt" w:hanging="36pt"/></w:pPr></w:lvl><w:lvl w:ilvl="6"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%7."/><w:lvlJc w:val="start"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="252pt"/></w:tabs><w:ind w:start="252pt" w:hanging="36pt"/></w:pPr></w:lvl><w:lvl w:ilvl="7"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%8."/><w:lvlJc w:val="start"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="288pt"/></w:tabs><w:ind w:start="288pt" w:hanging="36pt"/></w:pPr></w:lvl><w:lvl w:ilvl="8"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%9."/><w:lvlJc w:val="start"/><w:pPr><w:tabs><w:tab w:val="num" w:pos="324pt"/></w:tabs><w:ind w:start="324pt" w:hanging="36pt"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="0"/></w:num></w:numbering>');
		zipContent.file("word/_rels/document.xml.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/webSettings" Target="webSettings.xml"/><Relationship Id="rId2" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/settings" Target="settings.xml"/><Relationship Id="rId1" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/styles" Target="styles.xml"/><Relationship Id="rId5" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/theme" Target="theme/theme1.xml"/><Relationship Id="rId4" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/fontTable" Target="fontTable.xml"/><Relationship Id="rId6" Type="http://purl.oclc.org/ooxml/officeDocument/relationships/numbering" Target="numbering.xml"/></Relationships>');


        var strDocumentxml = documentxml.asText();


        //Nonused Variables... For now
        var contentAsString = strDocumentxml;
        var contentAsXml = $(strDocumentxml);


        //The docxtemplater variable becomes an object of the Docxtemplater class
        //The content of the template zip is passed into the docxtemplater variable
        docxtemplater = new Docxtemplater();
        docxtemplater.loadZip(zipContent);

        //The "loaded()" method is called after the template has been loaded
        //We passed the loaded object when we called the docxLoader function.
        //The functions that might be called are: docxReplaceVariables, docxReplaceTexts and docxReplaceBoth
        loaded();


    });
  },


  docxTemplateValidator : function(){

    var eText = "";

    try{
      //The docxtemplater tries to render the template, if it has errors, then it will stop and throw the error...
      docxtemplater.render();
    }
    catch (e) {
      //console.log(JSON.stringify(e.properties));
      var errors = e.properties.errors.length;
      eText = "=======================================================================\n";
      if (errors > 0)
      {
        e.properties.errors.forEach(function(element, index) {
          //console.log(element.properties);
          eText = eText +
          "Error No.: "+(index+1)+ "\n"+
          "Error Type: "+element.properties.id+ "\n"+
          "Error Description: ["+element.properties.explanation+"]\n"+
          "=======================================================================\n";
        });

        $("#"+templateValidatorRes).val(eText);
      }
      apex.message.clearErrors();
      amanda_docx_printer.docxErrorHandler('e', e.name+": ["+e.properties.explanation+"]");
    }

      eText = "=======================================================================\n";
      eText = eText +
      "This template is valid and No errors were found within it...\n"+
      "=======================================================================\n";
      $("#"+templateValidatorRes).val(eText);
      apex.message.clearErrors();
      amanda_docx_printer.docxErrorHandler('s', "This Template is Valid: [No Errors Found]");
  },



  docxDataSourceValidator : async function(){
    try{
      var dsArray = dataSourceItem.split(',');
      querycursor = "";

      dsArray.forEach(function(element){
        querycursor = querycursor+$("#"+element).val()+"~";
      });

      dataSources = await amanda_docx_printer.docxDataSourceRender(querycursor);
      $("#"+dataSourceValidatorRes).val(JSON.stringify(dataSources, null, 4));
      console.log(JSON.stringify(dataSources));
      apex.message.clearErrors();
      amanda_docx_printer.docxErrorHandler('s', "No errors were found");
    }
    catch(e){
      apex.message.clearErrors();
      amanda_docx_printer.docxErrorHandler('e', e.responseText);
      console.log(JSON.stringify(e));
      $("#"+dataSourceValidatorRes).val("One of the Datasources is not valid, Error: \n"+e.responseText.replace(/&quot;/g, "'"));
    }


  },



  docxReplaceVariables : function(){

    //The Template is populated with the dataSources coming from the Dynamic Action, then the template is rendereed
    //If any kind of error occurs during the render process, it is caught and thrown
    console.log("Starts docxReplaceVariables Method");
    //console.log("dataSources: "+dataSources);
	  //console.log(JSON.stringify(dataSources));
    //docxtemplater.setData(JSON.parse(dataSources));
    //console.log("JSON: "+JSON.parse(dataSources));

    try{

      console.log(dataSources);
      amanda_docx_printer.docxHTMLFinder();
      console.log(dataSources);

      docxtemplater.setData(dataSources);


      //The docxtemplater tries to render the template, if it has errors, then it will stop and throw the error...
      docxtemplater.render();

      //The docxOut variable is populated with the new content that is generated from the just populated and rendered template
      docxOut = docxtemplater.getZip().generate({
        type:"blob",
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      });

      //The docxDownloader method is called, passing the docxOut variable, which contains the new docx
      amanda_docx_printer.docxDownloader(docxOut);
    }
    catch (e) {
      console.log(e);
      apex.message.clearErrors();
      amanda_docx_printer.docxErrorHandler('e', "["+e+"]");
      //amanda_docx_printer.docxErrorHandler('e', e.name+": ["+e.properties.explanation+"]");
    }



  },

  docxReplaceTexts : function(varNode){
    return;
  },

  docxHTMLTranslate : function(varNode){
    //Function that will translate the String variable "varNode" (with HTML format) into an
    //OpenXml format string, then it will return its result
    console.log("Starts docxHTMLTranslate");
		console.log(html2docx.convertContent(varNode).string);
		return html2docx.convertContent(varNode).string;
  },

  docxIsObject : function(obj){
    return obj === Object(obj);
  },

  docxHTMLFinder : function(){
    //Function that finds the HTML content inside the elements of the dataSources
    //When an element is found, then it will call the docxHTMLTranslate function and
    //Replace the value of that element with the returned value
    console.log("Starts docxHTMLFinder");
    var htmlArray = new Array();
	  var counter = 0;

	  for(var i in dataSources)
	  {
  		var inside = dataSources[i];
  		for(var e in inside)
  		{
  		  var realdata = inside[e];
  		  for(var rd in realdata)
  		  {
          var element = String(realdata[rd]);
          if(typeof element != 'undefined')
          {
            if(element.startsWith("<"))
  		      {
      				var oldVal = realdata[rd];
      				realdata[rd] = amanda_docx_printer.docxHTMLTranslate(oldVal);
      				htmlArray[counter] = rd;
    			  }
          }
  		  }
  		}
	  }
  },

  docxDownloader : function(){
    //The saveAs method is used to download the populated template, and a name to the downloaded file is given
    //Using the docxOutName variable
    saveAs(docxOut,docxOutName);
  },

  docxDataSourceRender : async function(queries){
      return apex.server.process(
          "AmandaDocxDataSourceBuilder", {
              x01: queries
          },
          {loadingIndicator: $(trigger)}

      );
  },


  docxAjaxCaller : async function(){

    try{
      dataSources = await amanda_docx_printer.docxDataSourceRender(querycursor);
      amanda_docx_printer.docxReplaceVariables();
    }
    catch(e){
      apex.message.clearErrors();
      amanda_docx_printer.docxErrorHandler('e', e.responseText);
    }

  },


  docxErrorHandler : function(Type,Msg){
    var findSpan = $('#APEX_SUCCESS_MESSAGE').length;
    if (findSpan == 1){
        $('#APEX_SUCCESS_MESSAGE').remove();

        amanda_docx_printer.docxErrorHandler(Type,Msg);
    } else{
            $('#t_Body_content').prepend(
                '<span id="APEX_SUCCESS_MESSAGE" class="apex-page-success"><div class="t-Body-alert">'
                + '<div class="t-Alert t-Alert--defaultIcons t-Alert--success t-Alert--horizontal t-Alert--page t-Alert--colorBG" '
                + 'id="t_Alert_Success" role="alert" style="display: none;"><div class="t-Alert-wrap"><div class="t-Alert-icon">'
                + '<span style="color: white; font-size: 25px;" class="fa fa-check fa-xl fa-anim-flash"></span></div>'
                + '<div class="t-Alert-content"><div class="t-Alert-header"><h2 class="t-Alert-title">'
                + Msg
                + '</h2></div></div><div class="t-Alert-buttons"><button class="t-Button t-Button--noUI t-Button--icon t-Button--closeAlert" '
                + 'type="button" title="Close Notification"><span class="t-Icon icon-close"></span></button></div></div></div></div></span>');

                $('#t_Alert_Success > div > div.t-Alert-buttons > button[title="Close Notification"]')
                    .on('click',function(){
                    $('#APEX_SUCCESS_MESSAGE').remove();
                });

            if(Type=='s'){
                $.when(
                    $('#t_Alert_Success').show(),
                    $('#t_Alert_Success').fadeIn(300).delay(2000).fadeOut(400)
                ).then(
                        function(){
                            $('#APEX_SUCCESS_MESSAGE').remove();
                        });

            }else if(Type=='e'){
                    $('#t_Alert_Success').attr('style','background-color: #e95b54;');
                    $('#t_Alert_Success div div.t-Alert-icon span').removeClass('fa-check');
                    $('#t_Alert_Success div div.t-Alert-icon span').addClass('fa-close');
                    $('#t_Alert_Success').show();
            }
    }
  },



  init : function(){
    //The init method is the first one to be called when the Dynamic Action is triggered
    //The daThis variable receives all the properties and attributes coming from the Dynamic Action
    //Then this variable is passed to the docxSelector method which starts the whole process
    var daThis = this;
    trigger = "#"+this.triggeringElement.id;
    amanda_docx_printer.docxSelector(daThis.action);
  }
}
