// jQuery plugin to replace $.parseXML() in order to use MSXML 6.0 on IE (i.o. MSXML 3.0)
(function($){
  $.extend({
    parseXML_msxml6: function(string) {
      // code for Chrome, Mozilla, Firefox, Opera, etc.
      if (!isIE()) {
        var xml = $.parseXML(string); // jQuery OK for non-IE-browsers
      }
      //code for IE
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
      // code for Chrome, Mozilla, Firefox, Opera, etc.
      if (!isIE()) {
        xmlstring = (new XMLSerializer()).serializeToString(jqueryxml);
        // remove XHTML namespace declaration added by Firefox
        xmlstring = xmlstring.replace(' xmlns="http:\/\/www.w3.org/1999/xhtml"', '');
      }
      //code for IE
      else {
        xmlstring = jqueryxml.xml;
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

// MAIN ENTRY POINT

var applyXslt = function() {
  // collects all root .xslt nodes, converts them to XML documents and calls for processing
  var rootxsltnodes = $('.xslt').not('.xslt *');
  rootxsltnodes.each(function() {
    var htmlstring = $('<div/>').append($(this).clone()).html(); // clone() because otherwise this is removed from the DOM, for mysterious reasons
    var xmldocument = $.parseXML_msxml6(htmlstring);
    process(xmldocument,this); // will replace the DOM element 'this' by the result of processing the 'xml'
  });
};

var process = function(xml,dom) {
  // precondition: xml is an XML document copying the content of the dom DOM element
  // precondition: xml/dom has .xslt as root node
  var firstleafxsltnode = $(xml).find('.xslt').filter(function(index){
    return $(this).children('.xslt').length == 0;
  }).first(); // should always give result!
  if (firstleafxsltnode.is('.document')) {
    $('<span>').html('Downloading...').appendTo('.status');
    xmldocument(xml,dom,firstleafxsltnode);
  }
  else if (firstleafxsltnode.is('.documentcrossdomain')) {
    $('<span>').html('Downloading...').appendTo('.status');
    xmldocumentcrossdomain(xml,dom,firstleafxsltnode);
  }
  else if (firstleafxsltnode.is('.stylesheet')) {
    $('<span>').html('Applying stylesheets...').appendTo('.status');
    stylesheet(xml,dom,firstleafxsltnode);
  }
  else if (firstleafxsltnode.is('.unzip')) {
    $('<span>').html('Unpacking...').appendTo('.status');
    unzip(xml,dom,firstleafxsltnode);
  }
  else if (firstleafxsltnode.is('.decrypt')) {
    $('<span>').html('Decrypting...').appendTo('.status');
    decryptxml(xml,dom,firstleafxsltnode);
  }
  else if (firstleafxsltnode.is('.tsv')) {
    $('<span>').html('Parsing TSV data...').appendTo('.status');
    tsv2xml(xml,dom,firstleafxsltnode);
  }
};

var processresult = function(xml,dom,firstleafxsltnode,resultxmlstring){
  $('<span>').html('Please wait...').appendTo('.status');
  if (firstleafxsltnode.parents('.xslt').length == 0) { // is root node (note: .is('not(".xslt *")') won't work in IE) 
    var html = $('<div/>').append(resultxmlstring).children().first();
    if (html.hasClass("just-a-container")) {
      // XSLT is not allowed to return a set of elements, but your XSLT can assign this class
      // to the top-level element, in order to have it wasted overhere 
      $(dom).replaceWith(html.children());
    } else {
      $(dom).replaceWith(html);
    }
    //$("body").trigger('create'); // this activates jQuery Mobile
    //setTimeout(function() {$("body").trigger('create')}, 1);
    $(document).trigger('xsltdone');
    $('.status').remove();
    // listen for this event elsewhere, to know when XSLT results are available
    // DONE completely
  }
  else {
    var firstleafxsltnodeparent = firstleafxsltnode.parent();
    var resultxml = $.parseXML_msxml6(resultxmlstring);
    firstleafxsltnodeparent.get(0).replaceChild(resultxml.documentElement, firstleafxsltnode.get(0));
    setTimeout(function() {process(xml,dom)}, 1); // to prevent IE from raising error about script running too long
  }   
};

var xmldocument = function(xml,dom,n) {
  // returns an XML string !!
  $.ajax({
    url: n.attr("url"),
    dataType: "xml",
    success: function(data) {
      processresult(xml,dom,n,$.serializeXML(data));
    }
  });
};

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
}; 

var asyncUnzip = function(unzipper,filename,output) {
  var asyncUnzippedContent = function(unzipper,filename) {
    var def = $.Deferred();
    setTimeout(function(){
      var unzippedcontent = unzipper.file(filename);
      def.resolve(unzippedcontent);
    });
    return def.promise();
  };
  var asyncAsText = function(unzippedcontent){
    var def = $.Deferred();
    setTimeout(function(){
      unzippedcontent.asTextAsync(function(unzippedcontentastext) {
        output += '<div class="xlsx file" name="' + filename + '">';
        output += unzippedcontentastext.replace(/^<\?xml.*\?>/,'');  
        output += '</div>';
        def.resolve(output);
      });
    });
    return def.promise();
  };
  return asyncUnzippedContent(unzipper,filename).then(asyncAsText);
};

var unzip = function(xml,dom,n) {
  // returns an XML string !!
  var zipfile = load_binary_resource(n.attr("url"));
  var unzipper = new JSZip(zipfile);
  // to get one specified file out of the zipfile:
  if (n.attr("filename")) {
    for (var filename in unzipper.files) {
      if (filename.match(n.attr("filename"))) {
        var result = unzipper.file(filename).asText();
      }
    }
    processresult(xml,dom,n,result);
  }
  // to get all XML files out of the zipfile:
  else {
    var output = '<div class="xlsx zip">';
    var unzippingJob; 
    for (var filename in unzipper.files) {
      if (filename.match('\.xml$')) {
        if (!unzippingJob) {
          // start first job
          unzippingJob = asyncUnzip(unzipper,filename,output);
        } else {
          (function(unzipper,filename){
            unzippingJob = unzippingJob.then(function(output){
              // start next job
              return asyncUnzip(unzipper,filename,output)
            });
          }(unzipper,filename));
        }
      }
    }
    unzippingJob.done(function(output){
      // process result of final job
      output += '</div>';
      processresult(xml,dom,n,output);
    });
  }  
};

var tsv2xml = function(xml,dom,n) {
  // returns an XML string !!
  $.ajax({
    url: n.attr("url"),
    dataType: "text",
    success: function(data) {
      // use this document for creating XML
      // code for Chrome, Mozilla, Firefox, Opera, etc.
      if (!isIE()) {
        var rows = document.implementation.createDocument(null, 'rows', null);
      }
      // code for IE
      else {
        var rows = new ActiveXObject("Microsoft.XMLDOM");
        var root = el(rows, 'rows', []);
        rows.appendChild(root);
      }
      var lines = data.split('\n');
      var header = lines[0];
      var columns = header.split('\t');
      for (var c = 0; c < columns.length; c++) {
        columns[c] = columns[c].replace(/ /g,'_');
      }
      for (var i = 1; i < lines.length; i++) {
        var children = [];
        var fields = lines[i].split('\t');
        for (var c = 0; c < columns.length; c++) {
          var child = el(rows, columns[c], [fields[c]]);
          children.push(child);
        }
        var record = el(rows, 'row', children);
        rows.documentElement.appendChild(record);
      }
      processresult(xml,dom,n,$.serializeXML(rows));
    }
  });

};

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
};

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
};

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
  // code for Chrome, Mozilla, Firefox, Opera, etc.
  if (!isIE()) {
    var xsl = xslresponse.responseXML;
    var xsltProcessor = new XSLTProcessor();
    xsltProcessor.importStylesheet(xsl);
    $.each(n.get(0).attributes, function(i, attrib){
      xsltProcessor.setParameter(null, attrib.name, attrib.value);
    });
    var output = $.serializeXML(xsltProcessor.transformToFragment(xmltree,document));
    processresult(xml,dom,n,output);
  }
  // code for IE
  else {
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
    var output = new String(xslproc.output);
    processresult(xml,dom,n,output.toString()); // string-stuff to accomodate jQuery in IE
  }
};

function load_binary_resource(url) {  
  // code for Chrome, Mozilla, Firefox, Opera, etc.
  if (!isIE()) {
    var req = new XMLHttpRequest();  
    req.open('GET', url, false);  
    //XHR binary charset opt by Marcus Granado 2006 [http://mgran.blogspot.com]  
    req.overrideMimeType('text/plain; charset=x-user-defined');  
    req.send(null);  
    if (req.status != 200) return '';  
    return req.responseText;  
  }
  // code for IE
  else {
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

/* this code was online, but new browsers don't take scripts from github raw 
 * https://raw.github.com/gist/1393077/2b43c9ead04658fbff7444dcafb54558ab521195/gistfile1.js
 */

// use (16 chars of) 'password' to encrypt 'plaintext'

function encrypt(plaintext, password) {
  var v = new Array(2), k = new Array(4), s = "", i;

  plaintext = escape(plaintext);  // use escape() so only have single-byte chars to encode 

  // build key directly from 1st 16 chars of password
  for (var i=0; i<4; i++) k[i] = Str4ToLong(password.slice(i*4,(i+1)*4));

  for (i=0; i<plaintext.length; i+=8) {  // encode plaintext into s in 64-bit (8 char) blocks
    v[0] = Str4ToLong(plaintext.slice(i,i+4));  // ... note this is 'electronic codebook' mode
    v[1] = Str4ToLong(plaintext.slice(i+4,i+8));
    code(v, k);
    s += LongToStr4(v[0]) + LongToStr4(v[1]);
  }

  return escCtrlCh(s);
  // note: if plaintext or password are passed as string objects, rather than strings, this
  // function will throw an 'Object doesn't support this property or method' error
}

// use (16 chars of) 'password' to decrypt 'ciphertext' with xTEA

function decrypt(ciphertext, password) {
  var v = new Array(2), k = new Array(4), s = "", i;

  for (var i=0; i<4; i++) k[i] = Str4ToLong(password.slice(i*4,(i+1)*4));

  ciphertext = unescCtrlCh(ciphertext);
  for (i=0; i<ciphertext.length; i+=8) {  // decode ciphertext into s in 64-bit (8 char) blocks
    v[0] = Str4ToLong(ciphertext.slice(i,i+4));
    v[1] = Str4ToLong(ciphertext.slice(i+4,i+8));
    decode(v, k);
    s += LongToStr4(v[0]) + LongToStr4(v[1]);
  }

  // strip trailing null chars resulting from filling 4-char blocks:
  s = s.replace(/\0+$/, '');

  return unescape(s);
}


function code(v, k) {
  // Extended TEA: this is the 1997 revised version of Needham & Wheeler's algorithm
  // params: v[2] 64-bit value block; k[4] 128-bit key
  var y = v[0], z = v[1];
  var delta = 0x9E3779B9, limit = delta*32, sum = 0;

  while (sum != limit) {
    y += (z<<4 ^ z>>>5)+z ^ sum+k[sum & 3];
    sum += delta;
    z += (y<<4 ^ y>>>5)+y ^ sum+k[sum>>>11 & 3];
    // note: unsigned right-shift '>>>' is used in place of original '>>', due to lack 
    // of 'unsigned' type declaration in JavaScript (thanks to Karsten Kraus for this)
  }
  v[0] = y; v[1] = z;
}

function decode(v, k) {
  var y = v[0], z = v[1];
  var delta = 0x9E3779B9, sum = delta*32;

  while (sum != 0) {
    z -= (y<<4 ^ y>>>5)+y ^ sum+k[sum>>>11 & 3];
    sum -= delta;
    y -= (z<<4 ^ z>>>5)+z ^ sum+k[sum & 3];
  }
  v[0] = y; v[1] = z;
}


// supporting functions

function Str4ToLong(s) {  // convert 4 chars of s to a numeric long
  var v = 0;
  for (var i=0; i<4; i++) v |= s.charCodeAt(i) << i*8;
  return isNaN(v) ? 0 : v;
}

function LongToStr4(v) {  // convert a numeric long to 4 char string
  var s = String.fromCharCode(v & 0xFF, v>>8 & 0xFF, v>>16 & 0xFF, v>>24 & 0xFF);
  return s;
}

function escCtrlCh(str) {  // escape control chars which might cause problems with encrypted texts
  return str.replace(/[\0\t\n\v\f\r\xa0'"!]/g, function(c) { return '!' + c.charCodeAt(0) + '!'; });
}

function unescCtrlCh(str) {  // unescape potentially problematic nulls and control characters
  return str.replace(/!\d\d?\d?!/g, function(c) { return String.fromCharCode(c.slice(1,-1)); });
}


// copied from http://stackoverflow.com/questions/3191179/generate-xml-document-in-memory-with-javascript
// and modified to take an array of children rather than a variable number of arguments

// function that creates the XML structure
function el(doc, name, children) {
    var node = doc.createElement(name);

    for(var i = 0; i < children.length; i++) {
        var child = children[i];
        if(typeof child == 'string') {
            child = doc.createTextNode(child);
        }
        node.appendChild(child);
    }

    return node;
};

// copied from http://stackoverflow.com/a/15983064/591336
// to detect IE browser

function isIE () {
  var myNav = navigator.userAgent.toLowerCase();
  return (myNav.indexOf('msie') != -1) ? parseInt(myNav.split('msie')[1]) : false;
}

