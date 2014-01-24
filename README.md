xslt-in-the-browser
===================

Javascript code for loading XML, or even zipped XML (like in XLSX), in the browser, merge XML files, perform XSLT and display the result.

Copying all files onto a web-server should make the demo work. A copy is published at the link below, but I won't promise to keep it up-to-date or even alive:

http://users.telenet.be/vicmortelmans/xslt-in-the-browser/demo.html

My use case for this toolset is to have an XLSX document on a Document Managment System, which contains tabular data managed by users familiar with MS Office, and to allow this data to be published in a web context in a more interactive way and with custom formatting, from within the same DMS (meaning: realying on files hosted on a static webserver). Another requirement is that the data should not leave the intranet, because then I know a lot better solutions involving Google Drive for example.

This is what the files are about:
* applyXslt.js - contains the javascript code for this project. There's a main function applyXslt() which you should call when your html is loaded (see the example in demo.html), as it's responsible for replacing the placeholders which you put in your html by the actual processed data
* dataMS.xlsx - datafile for demo purpose only
* dataMS.xml - datafile for demo purpose only
* jquery-1.8.2.js - 3rd party code
* js-inflate.min.js - 3rd party code
* js-unzip.min.js - 3rd party code
* msxls2xml.xslt - stylesheet that transforms Microsoft 2003 Spreadsheet XML format into a more semantical XML format (actually, an HTML table with speciall classes)
* msxlsx2xml.xslt - stylesheet that transforms Microsoft Office Open XML format (XLSX) into a more semantical XML format
* demo.html - webpage for demo purpose only and where some information on the different types of placeholders is stored
