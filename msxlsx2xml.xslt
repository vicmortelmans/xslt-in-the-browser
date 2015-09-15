<?xml version="1.0"?>
<xsl:stylesheet version="1.0" xmlns:xhtml="http://www.w3.org/1999/xhtml"
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:exslt="http://exslt.org/common"
  xmlns:os="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:msxsl="urn:schemas-microsoft-com:xslt"
  exclude-result-prefixes="exslt msxsl">

  <!-- IMPORTANT NOTE: in FIREFOX, there's no ROOT ('/') in the XML, at least not where it's expected -->

  <!--msxsl:script language="JScript" implements-prefix="exslt">
   this['node-set'] =  function (x) {
    return x;
    }
  </msxsl:script-->
  
  <xsl:output method="xml" indent="no" encoding="UTF-8"/>

  <xsl:param name="worksheet" select="'Sheet1'"/>
  <xsl:param name="headerrow" select="1"/>
  
  <xsl:key name="rows" match="//div/div/os:worksheet/os:sheetData/os:row" use="generate-id()"/>
  <xsl:key name="fillid-by-style"
    match="//div/div[@name = 'xl/styles.xml']/os:styleSheet/os:cellXfs/os:xf/@fillId"
    use="count(../preceding-sibling::os:xf)"/>
  <xsl:key name="string-by-v"
    match="//div/div[@name = 'xl/sharedStrings.xml']/os:sst/os:si/os:t"
    use="count(../preceding-sibling::os:si)"/>
  <xsl:key name="colors-by-fillid"
    match="//div/div[@name = concat('xl/worksheets/sheet',substring-after(//div/div[@name = 'xl/workbook.xml']/os:workbook/os:sheets/os:sheet[@name = 'Colors']/@r:id,'rId'),'.xml')]/os:worksheet/os:sheetData/os:row/os:c"
    use="key('fillid-by-style',@s)"/>

  <xsl:template match="@*|node()">
    <xsl:apply-templates select="@*|node()"/>
  </xsl:template>

  <xsl:template match="div[contains(@class,'xlsx zip')]"><!-- ROOT element, not to be addressed as '/' in FIREFOX -->
      <xsl:apply-templates select="div[@name = 'xl/workbook.xml']/os:workbook/os:sheets/os:sheet"/>
  </xsl:template>

  <xsl:template match="os:sheet">
    <xsl:if test="@name = $worksheet">
      <xsl:variable name="sheetname" select="concat('xl/worksheets/sheet',substring-after(current()/@r:id,'rId'),'.xml')"/>
      <table class=" xml table ">
        <xsl:apply-templates mode="headerrow" select="//div/div[@name = $sheetname]/os:worksheet/os:sheetData/os:row[@r = $headerrow]"/>
        <xsl:apply-templates mode="rows" select="//div/div[@name = $sheetname]/os:worksheet/os:sheetData/os:row[@r &gt; $headerrow]">
          <xsl:with-param name="sheetname" select="$sheetname"/>
        </xsl:apply-templates>
      </table>
    </xsl:if>
  </xsl:template>

  <xsl:template match="os:row" mode="headerrow">
    <!-- HEADER ROW -->
    <tr class=" xml row header ">
      <xsl:for-each select="./os:c">
        <xsl:variable name="value">
          <xsl:call-template name="cellvalue">
            <xsl:with-param name="t" select="@t"/>
            <xsl:with-param name="v" select="os:v"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="name">
          <xsl:call-template name="unspace">
            <xsl:with-param name="string" select="$value"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="colorattribute">
          <xsl:call-template name="colorattributeunspaced">
            <xsl:with-param name="s" select="@s"/>
          </xsl:call-template>
        </xsl:variable>
        <th class=" xml cell header {$name} {$colorattribute} ">
          <xsl:if test="$colorattribute != ''">
            <xsl:attribute name="color">
              <xsl:value-of select="$colorattribute"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:value-of select="$value"/>
        </th>
      </xsl:for-each>
    </tr>
  </xsl:template>

  <xsl:template match="os:row" mode="rows">
    <xsl:param name="sheetname"/>
    <!-- DATA ROW -->
    <xsl:variable name="row" select="."/>
    <xsl:variable name="rownr" select="@r"/>
    <xsl:variable name="rowid" select="generate-id()"/>
    <tr class=" xml row ">
      <xsl:for-each select="//div/div[@name = $sheetname]/os:worksheet/os:sheetData/os:row[@r = $headerrow]/os:c">
        <!-- iterating header row cells, because datarows may not have a cell for each column -->
        <xsl:variable name="r"><!-- is the coordinates 'A1',... of the target cell -->
          <xsl:value-of select="substring-before(@r,$headerrow)"/><!-- is the character 'A', 'B', 'C',... indicating the COLUMN -->
          <xsl:value-of select="$rownr"/><!-- is the number 1, 2, 3,... indicating the ROW -->
        </xsl:variable>
        <!--xsl:message>inner loop <xsl:value-of select="$r"/></xsl:message-->
        <xsl:variable name="headercolorattribute">
          <xsl:call-template name="colorattributeunspaced">
            <xsl:with-param name="s" select="@s"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="name">
          <xsl:call-template name="unspace">
            <xsl:with-param name="string">
              <xsl:call-template name="cellvalue">
                <xsl:with-param name="t" select="@t"/>
                <xsl:with-param name="v" select="os:v"/>
              </xsl:call-template>
            </xsl:with-param>
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="colorattribute">
          <xsl:call-template name="colorattributeunspaced">
            <xsl:with-param name="s" select="key('rows',$rowid)/os:c[@r = $r]/@s"/>
          </xsl:call-template>
        </xsl:variable>
        <xsl:variable name="cellvalue">
          <xsl:call-template name="cellvalue">
            <xsl:with-param name="t" select="key('rows',$rowid)/os:c[@r = $r]/@t"/>
            <xsl:with-param name="v" select="key('rows',$rowid)/os:c[@r = $r]/os:v"/>
          </xsl:call-template>
        </xsl:variable>
        <td class=" xml cell data {$name} {$colorattribute} ">
          <xsl:if test="$colorattribute != ''">
            <xsl:attribute name="color">
              <xsl:value-of select="$colorattribute"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="$headercolorattribute != ''">
            <xsl:attribute name="{$headercolorattribute}">
              <xsl:value-of select="$name"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:if test="$cellvalue != ''">
            <xsl:attribute name="value">
              <xsl:value-of select="$cellvalue"/>
            </xsl:attribute>
          </xsl:if>
          <xsl:value-of select="$cellvalue"/>
        </td>
      </xsl:for-each> 
    </tr>    
  </xsl:template>

  <xsl:template name="unspace">
    <xsl:param name="string" select="."/>
    <xsl:value-of select="translate(normalize-space($string),' ','_')"/>
  </xsl:template>

  <xsl:template name="cellvalue">
    <xsl:param name="t" select="'s'"/>
    <xsl:param name="v"/>
    <xsl:choose>
      <xsl:when test="$t = 's'">
        <xsl:value-of select="key('string-by-v',$v)"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="$v"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="colorattributeunspaced">
    <xsl:param name="s"/>
    <xsl:variable name="f" select="key('fillid-by-style',$s)"/>
    <xsl:if test="$f != '0'">
      <xsl:call-template name="unspace">
        <xsl:with-param name="string">
          <xsl:for-each select="key('colors-by-fillid',$f)">
            <xsl:call-template name="cellvalue">
              <xsl:with-param name="v" select="os:v"/>
            </xsl:call-template>
            <!-- reminder: this won't work if the contents of the color field is NOT a shared string (e.g. a number) -->
          </xsl:for-each>
        </xsl:with-param>
      </xsl:call-template>
    </xsl:if>
  </xsl:template>

</xsl:stylesheet>
