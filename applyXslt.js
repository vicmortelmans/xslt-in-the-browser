// jQuery plugin to replace $.parseXML() in order to use MSXML 6.0 on IE (i.o. MSXML 3.0)
(function($){
  $.extend({
    parseXML_msxml6: function(string) {
      if (window.DOMParser) {
        var xml = $.parseXML(string); // jQuery OK for non-IE-browsers
      }
      else {
        var xml = new ActiveXObject("Msxml2.FreeThreadedDOMDocument.6.0");
        xml.async = false;
        xml.loadXML(string);
      }
      return xml;
    }
  });
})(jQuery);

// jQuery plugin to transform an XML document into a string
(function($){
  $.extend({
    serializeXML: function(jqueryxml) {
      var xmlstring;
      //code for IE
      if (window.ActiveXObject) {
        xmlstring = jqueryxml.xml;
      }
      // code for Chrome, Mozilla, Firefox, Opera, etc.
      else {
        xmlstring = (new XMLSerializer()).serializeToString(jqueryxml);
      }
      return xmlstring;
    }
  });
})(jQuery);

// jQuery plugin to do process each individual text node in a document tree
$.fn.processText = function( process ) {  
  return this.each(function(){  
    var node = this.firstChild,  
      val,  
      new_val,  
      remove = [];  
    if ( node ) {  
      do {  
        if ( node.nodeType === 3 ) {  
          val = node.nodeValue;  
          new_val = process(val);  
          if ( new_val !== val ) {  
            node.nodeValue = new_val;  
          }  
        }  
        else if ( node.childNodes != null ) {
          $(node).processText(process);
        }
      } while ( node = node.nextSibling );  
    }  
    remove.length && $(remove).remove();  
  });  
};  

// support cross-domain ajax calls
jQuery.support.cors = true;

var applyXslt = function() {
  // collects all root .xslt nodes, converts them to XML documents and calls for processing
  var rootxsltnodes = $('.xslt').not('.xslt *');
  rootxsltnodes.each(function() {
    var htmlstring = $('<div/>').append($(this).clone()).html(); // clone() because otherwise this is removed from the DOM, for mysterious reasons
    var xmldocument = $.parseXML_msxml6(htmlstring);
    process(xmldocument,this); // will replace the DOM element 'this' by the result of processing the 'xml'
  });
}

var process = function(xml,dom) {
  // precondition: xml is an XML document copying the content of the dom DOM element
  // precondition: xml/dom has .xslt as root node
  var firstleafxsltnode = $(xml).find('.xslt').filter(function(index){
    return $(this).children('.xslt').length == 0;
  }).first(); // should always give result!
  if (firstleafxsltnode.is('.document')) {
    $(dom).append('<p>reading data...</p>');
    xmldocument(xml,dom,firstleafxsltnode);
  }
  else if (firstleafxsltnode.is('.documentcrossdomain')) {
    $(dom).append('<p>reading data...</p>');
    xmldocumentcrossdomain(xml,dom,firstleafxsltnode);
  }
  else if (firstleafxsltnode.is('.stylesheet')) {
    $(dom).append('<p>transforming data...</p>');
    stylesheet(xml,dom,firstleafxsltnode);
  }
  else if (firstleafxsltnode.is('.unzip')) {
    $(dom).append('<p>loading data...</p>');
    unzip(xml,dom,firstleafxsltnode);
  }
  else if (firstleafxsltnode.is('.decrypt')) {
    $(dom).append('<p>decoding data..</p>');
    decryptxml(xml,dom,firstleafxsltnode);
  }
}

var processresult = function(xml,dom,firstleafxsltnode,resultxmlstring){
  if (firstleafxsltnode.parents('.xslt').length == 0) { // is root node (note: .is('not(".xslt *")') won't work in IE) 
    var html = $('<div/>').append(resultxmlstring).children().first();
    $(dom).replaceWith(html); // DONE completely
    $(document).trigger('xsltdone');
  }
  else {
    var firstleafxsltnodeparent = firstleafxsltnode.parent();
    var resultxml = $.parseXML_msxml6(resultxmlstring);
    firstleafxsltnodeparent.get(0).replaceChild(resultxml.documentElement, firstleafxsltnode.get(0));
    setTimeout(function() {process(xml,dom)}, 1); // to prevent IE from raising error about script running too long
  }   
}

var xmldocument = function(xml,dom,n) {
  // returns an XML string !!
  $.ajax({
    url: n.attr("url"),
    dataType: "xml",
    success: function(data) {
      processresult(xml,dom,n,$.serializeXML(data));
    }
  });
} 

var xmldocumentcrossdomain = function(xml,dom,n) {
  // returns an XML string !!
  var yqlquery = "use 'http://github.com/vicmortelmans/yql-tables/raw/master/data/xmlstring.xml' as xmlstring; select * from xmlstring where url='" + n.attr("url") + "'";
  var yqlurl = "https://query.yahooapis.com/v1/public/yql?q=" + encodeURIComponent(yqlquery) + "&format=json";
  $.ajax({
    url: yqlurl,
    dataType: "jsonp",
    success: function(data) {
      processresult(xml,dom,n,data.query.results.result);
    }
  });
} 

var unzip = function(xml,dom,n) {
  // returns an XML string !!
  var zipfile = load_binary_resource(n.attr("url"));
  var unzipper = new JSUnzip(zipfile);
  unzipper.readEntries();
  if (n.attr("filename")) {
    for (var i = 0; i < unzipper.entries.length; i++) {
      var entry = unzipper.entries[i];
      if (entry.fileName.match(n.attr("filename"))) {
        var uncompressed = uncompress(entry);
        var result = uncompressed;
      }
    }
  }
  else {
    var output = '<div class="xlsx zip">';
    for (var i = 0; i < unzipper.entries.length; i++) {
      var entry = unzipper.entries[i];
      if (entry.fileName.match('\.xml$')) {
        var uncompressed = uncompress(entry);
        output += '<div class="xlsx file" name="' + entry.fileName + '">';
        output += uncompressed.replace(/^<\?xml.*\?>/,'');
        output += '</div>'
      }
    }
    output += '</div>';
    var result = output;
  }  
  processresult(xml,dom,n,result);
}

var decryptxml = function(xml,dom,n) {
  var input = n.children().first(); // assumed is only one child!
  var key = n.attr('key');
  var deletefirstrow = n.attr('deletefirstrow');
  if (deletefirstrow == 'yes') {  // this is a workaround for spreadsheets having dummy column headers
                                  // which was only needed when applying the semantic-xml-xslt before
                                  // downloading the data
    input.find('row').first().remove();
  }
  input.processText(function(string) {
    return decrypt(unescape($.trim(string) || ""),key);
  });
  processresult(xml,dom,n,$.serializeXML(input.get(0)));
}

var uncompress = function(entry) {
  // Uncompressed
  if (entry.compressionMethod === 0) {
    var uncompressed = entry.data; 
  }
  // Deflated
  else if (entry.compressionMethod === 8) {
    var uncompressed = JSInflate.inflate(entry.data);
  }
  return uncompressed;
}

var stylesheet = function(xml,dom,n) {
  // precondition: n is always a leaf .xslt node
  // returns an XML string !!
  var xmltree = n.children().first().get(0);
  if (n.is('.debug')) {
    $('body').append('<textarea rows="20" cols="150"></textarea>');
    var xmlstring = $.serializeXML(xmltree);
    $('textarea').last().val(xmlstring);
  }
  var xslresponse = $.ajax({
    async: false,
    url: n.attr("url"),
    dataType: "text xml"
  });
  // code for IE
  var output;
  if (window.ActiveXObject) {
    var xsl = new ActiveXObject("Msxml2.FreeThreadedDOMDocument.6.0");
    xsl.setProperty("AllowXsltScript", true);
    xsl.loadXML(xslresponse.responseText);
    var xslt = new ActiveXObject("Msxml2.XSLTemplate.6.0");
    xslt.stylesheet = xsl;
    var xslproc = xslt.createProcessor();
    var xmldoc = new ActiveXObject("Msxml2.FreeThreadedDOMDocument.6.0");
    var xmlclone = xmltree.cloneNode(true);
    xmldoc.appendChild(xmlclone);
    xslproc.input = xmldoc;
    $.each(n.get(0).attributes, function(i, attrib){
      xslproc.addParameter(attrib.name, attrib.value);
    });
    xslproc.transform();
    output = xslproc.output;
  }
  // code for Chrome, Mozilla, Firefox, Opera, etc.
  else if (document.implementation && document.implementation.createDocument) {
    var xsl = xslresponse.responseXML;
    var xsltProcessor = new XSLTProcessor();
    xsltProcessor.importStylesheet(xsl);
    $.each(n.get(0).attributes, function(i, attrib){
      xsltProcessor.setParameter(null, attrib.name, attrib.value);
    });
    output = $.serializeXML(xsltProcessor.transformToFragment(xmltree,document));
  }
  processresult(xml,dom,n,output);
}

function load_binary_resource(url) {  
  // code for IE
  if (window.ActiveXObject) {
    var req = new window.XMLHttpRequest();
    req.open("GET", url, false);
    req.setRequestHeader("Accept-Charset","x-user-defined");
    req.send();
    if (req.readyState == 4) {
      if (req.status == 200) {
        var fileContents = "";
        fileContents = GetIEByteArray_ByteStr(req.responseBody);
        return fileContents;
      }
      else {
        alert("download failed, status " + req.status);
      }
    }
  }
  // code for Chrome, Mozilla, Firefox, Opera, etc.
  else if (document.implementation && document.implementation.createDocument) {
    var req = new XMLHttpRequest();  
    req.open('GET', url, false);  
    //XHR binary charset opt by Marcus Granado 2006 [http://mgran.blogspot.com]  
    req.overrideMimeType('text/plain; charset=x-user-defined');  
    req.send(null);  
    if (req.status != 200) return '';  
    return req.responseText;  
  }
}  

function GetIEByteArray_ByteStr(IEByteArray) {
  var ByteMapping = {};
  for ( var i = 0; i < 256; i++ ) {
    for ( var j = 0; j < 256; j++ ) {
      ByteMapping[ String.fromCharCode( i + j * 256 ) ] = 
        String.fromCharCode(i) + String.fromCharCode(j);
    }
  }
  var rawBytes = IEBinaryToArray_ByteStr(IEByteArray);
  var lastChr = IEBinaryToArray_ByteStr_Last(IEByteArray);
  return rawBytes.replace(/[\s\S]/g, 
    function( match ) { return ByteMapping[match]; }) + lastChr;
}
 
var IEBinaryToArray_ByteStr_Script = 
  "<!-- IEBinaryToArray_ByteStr -->\r\n"+
  "<script type='text/vbscript'>\r\n"+
  "Function IEBinaryToArray_ByteStr(Binary)\r\n"+
  "    IEBinaryToArray_ByteStr = CStr(Binary)\r\n"+
  "End Function\r\n"+
  "Function IEBinaryToArray_ByteStr_Last(Binary)\r\n"+
  "	Dim lastIndex\r\n"+
  "	lastIndex = LenB(Binary)\r\n"+
  "	if lastIndex mod 2 Then\r\n"+
  "		IEBinaryToArray_ByteStr_Last = Chr( AscB( MidB( Binary, lastIndex, 1 ) ) )\r\n"+
  "	Else\r\n"+
  "		IEBinaryToArray_ByteStr_Last = "+'""'+"\r\n"+
  "	End If\r\n"+
  "End Function\r\n"+
  "</script>\r\n";
document.write(IEBinaryToArray_ByteStr_Script);

