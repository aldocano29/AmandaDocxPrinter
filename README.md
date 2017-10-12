# Oracle APEX Dynamic Action Plugin - AmandaDocxPrinter
Dynamic Action Plugin to Merge data from a query into a docx Template.


## Changelog

#### 1.0.0 - Initial Release


## Install

- Import Plugin File **dynamic_action_plugin_com_amandasoft_docxprinter.sql** from the main directory into your Application.
- Run the **PK_AMANDASOFT_PLUGINS_SPEC.sql** and **PK_AMANDASOFT_PLUGINS_BODY.sql** files from the **plslq** directory to create the Package (inside the Application Schema) used by this plugin.
- Create an Ajax Callback Application Process (Shared Components) named **AmandaDocxDataSourceBuilder** using as PLSQL code the content of **/src/AmandaDocxDataSourceRender.sql**.


## Plugin Settings

Available Plugin Settings :
- **Template** - the name of the docx template inside your Application's Static Files (required)
- **Process Type** - the process that will be used during the Plugin Call "Replace Variables with Datasource(s)" (required)
- **Output Docx Name** - the name for the output docx. You have to include the .docx at the end of the Output Name (optional)
- **DataSources** - the item(s) holding the queries used to bring the data you will merge with the templates (required)
- **Template Validator - Result** - just to be used inside the Demo Application (optional)
- **DataSource Validator - Result** - just to be used inside the Demo Application (optional)
- **DataSource Validator - ITEM(s)** - just to be used inside the Demo Application (optional)



## How to use
- Create a hidden but unprotected item to hold the Query you will use to bring data from your Schema (If the query references other items as parameters, do not use :P_ITEM insted use &P_ITEM.)
- Create a new Dynamic Action based on the Plugin
- Add the name of your Docx Template (located inside your Static Application Files) but without the "#APP_IMAGES#" reference.
- Add the Process Type "Replace Variables with Datasource(s)"
- Add the item(s) holding the Query/Queries to the DataSources (If more than one is needed, separate them by commas)

## Demo Application
- Inside **/plsql** you will find the Demo Application export file.
- Also the Demo Application is at [https://apex.oracle.com/pls/apex/f?p=118977]
- Credentials: demo/demo

## Related info
Based upon the DocxTemplater javascript library by Edgar Hipp.
[https://github.com/open-xml-templating/docxtemplater]


## Preview
## ![](https://github.com/aldocano29/AmandaDocxPrinter/blob/master/img/Preview.png)
