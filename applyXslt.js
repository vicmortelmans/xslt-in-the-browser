var applyXslt = function(node, returnXML) {
  // returns an XML string !!
  var n = node ? node : $('body');
  if (n.is('.xslt')) {
    if (n.is('.document')) {
      var xml = xmldocument(n);
    }
    else if (n.is('.stylesheet')) {
      var xml = stylesheet(n);
    }
    else if (n.is('.unzip')) {
      var xml = unzip(n);
    }
    if (returnXML) {
      return xml;
    }
    else {
      n.replaceWith(xml);
    }
  }
  else {
    if (n.find('.xslt')) {
      n.children().each(function() {
        applyXslt($(this));
      });
    }
    if (returnXML) {
      var xml = htmlcontents(n);
      return xml;
    }
  }
};

var xmldocument = function(n) {
  // returns an XML string !!
  var xmlresponse = $.ajax({
    async: false,
    url: n.attr("url"),
    dataType: "text html"
  });
  var xml = xmlresponse.responseText;
  return xml;
};

var unzip = function(n) {
  // returns an XML string !!
  var zipfile = load_binary_resource(n.attr("url"));
  var unzipper = new JSUnzip(zipfile);
  unzipper.readEntries();
  if (n.attr("filename")) {
    for (var i = 0; i < unzipper.entries.length; i++) {
      var entry = unzipper.entries[i];
      if (entry.fileName.match(n.attr("filename"))) {
        var uncompressed = uncompress(entry);
        return uncompressed;
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
    return output;
  }  
};

var uncompress = function(entry) {
  // Uncompressed
  if (entry.compressionMethod === 0) {
    var uncompressed = entry.data; 
  }
  // Deflated
  else if (entry.compressionMethod === 8) {
    var uncompressed = readUTF8String(JSInflate.inflate(entry.data));
  }
  return uncompressed;
};

var htmlcontents = function(n) {
  var htmlstring = n.html().replace(/( \w+)=(\w+[ /])/gm,'$1="$2"').replace(/<\?/gm,'<!--?').replace(/\?>/gm,'?-->');
  // removed replace(/(\r\n|\n|\r)/gm,'').
  return htmlstring;
};

var stylesheet = function(n) {
  // returns an XML string !!
  // document as child
  if (n.children().first().is('.xslt')) {
    var xmlstring = applyXslt(n.children().first(), true);
  }
  // html as child
  else {    
    applyXslt(n.children().first()); // recursion to children; we assume only one child
    var xmlstring = htmlcontents(n); 
  }
  var xslresponse = $.ajax({
    async: false,
    url: n.attr("url"),
    dataType: "text xml"
  });
  // code for IE
  if (window.ActiveXObject) {
    var xml = new ActiveXObject("Msxml2.FreeThreadedDOMDocument.6.0");
    xml.async = false;
    xml.loadXML(xmlstring); 
    var xsl = new ActiveXObject("Msxml2.FreeThreadedDOMDocument.6.0");
    xsl.setProperty("AllowXsltScript", true);
    xsl.loadXML(xslresponse.responseText);
    var xslt = new ActiveXObject("Msxml2.XSLTemplate.6.0");
    xslt.stylesheet = xsl;
    var xslproc = xslt.createProcessor();
    xslproc.input = xml;
    $.each(n.get(0).attributes, function(i, attrib){
      xslproc.addParameter(attrib.name, attrib.value);
    });
    xslproc.transform();
    var output = xslproc.output;
  }
  // code for Chrome, Mozilla, Firefox, Opera, etc.
  else if (document.implementation && document.implementation.createDocument) {
    var xml = $.parseXML(xmlstring);
    var xsl = xslresponse.responseXML;
    var xsltProcessor = new XSLTProcessor();
    xsltProcessor.importStylesheet(xsl);
    $.each(n.get(0).attributes, function(i, attrib){
      xsltProcessor.setParameter(null, attrib.name, attrib.value);
    });
    var output = (new XMLSerializer()).serializeToString(xsltProcessor.transformToFragment(xml,document));
  }
  return output;
};

var load_binary_resource = function(url) {  
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
}; 

var GetIEByteArray_ByteStr = function(IEByteArray) {
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
};

// following function copied from http://snipplr.com/view/31206/ and patched
var readUTF8String = function(bytes) {
  var ix = 0;
  if( bytes.slice(0,3) == "\xEF\xBB\xBF") {
    ix = 3;
  }
  var string = "";
  for( ; ix < bytes.length; ix++ ) {
    var byte1 = bytes.charCodeAt(ix);
    if( byte1 < 0x80 ) {
      string += String.fromCharCode(byte1);
    } else if( byte1 >= 0xC2 && byte1 < 0xE0 ) {
      var byte2 = bytes.charCodeAt(++ix);
      string += String.fromCharCode(((byte1&0x1F)<<6) + (byte2&0x3F));
    } else if( byte1 >= 0xE0 && byte1 < 0xF0 ) {
      var byte2 = bytes.charCodeAt(++ix);
      var byte3 = bytes.charCodeAt(++ix);
      string += String.fromCharCode(((byte1&0xFF)<<12) + ((byte2&0x3F)<<6) + (byte3&0x3F));
    } else if( byte1 >= 0xF0 && byte1 < 0xF5) {
      var byte2 = bytes.charCodeAt(++ix);
      var byte3 = bytes.charCodeAt(++ix);
      var byte4 = bytes.charCodeAt(++ix);
      var codepoint = ((byte1&0x07)<<18) + ((byte2&0x3F)<<12)+ ((byte3&0x3F)<<6) + (byte4&0x3F);
      codepoint -= 0x10000;
      string += String.fromCharCode(
        (codepoint>>10) + 0xD800,
        (codepoint&0x3FF) + 0xDC00
      );
    }
  }
  return string;
};

var IEBinaryToArray_ByteStr_Script = 
  "<!-- IEBinaryToArray_ByteStr -->\r\n"+
  "<script type='text/vbscript'>\r\n"+
  "Function IEBinaryToArray_ByteStr(Binary)\r\n"+
  "	IEBinaryToArray_ByteStr = CStr(Binary)\r\n"+
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

