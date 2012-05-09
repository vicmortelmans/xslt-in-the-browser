<?xml version="1.0"?>
<xsl:stylesheet version="1.0" xmlns="http://www.w3.org/1999/xhtml"
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:exslt="http://exslt.org/common"
  xmlns:ss="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  xmlns:msxsl="urn:schemas-microsoft-com:xslt"
  exclude-result-prefixes="exslt msxsl">

  <msxsl:script language="JScript" implements-prefix="exslt">
   this['node-set'] =  function (x) {
    return x;
    }
  </msxsl:script>
  
  <xsl:output method="html" indent="yes"/>

  <xsl:param name="worksheet" select="'Sheet1'"/>
  <xsl:param name="headerrow" select="1"/>
  
  <xsl:template match="@*|node()">
    <xsl:apply-templates select="@*|node()"/>
  </xsl:template>

  <xsl:template match="ss:sheets/ss:sheet">
    <xsl:if test="@name = $worksheet">
      <table class=" xml table ">
        <xsl:apply-templates mode="rows" select="//div[@class = 'xlsx file'][@name = concat('xl/worksheets/sheet',current()/@sheetId,'.xml')]//ss:row[count(preceding-sibling::ss:row) = $headerrow - 1]"/>
        <xsl:apply-templates mode="rows" select="//div[@class = 'xlsx file'][@name = concat('xl/worksheets/sheet',current()/@sheetId,'.xml')]//ss:row[count(preceding-sibling::ss:row) &gt; $headerrow - 1]"/>
      </table>
    </xsl:if>
  </xsl:template>

  <xsl:template match="ss:row" mode="rows">
    <xsl:choose>
      <xsl:when test="count(preceding-sibling::ss:row) = $headerrow - 1">
        <tr class=" xml row header ">
          <xsl:for-each select="./ss:c">
            <xsl:variable name="value">
              <xsl:call-template name="cellvalue"/>
            </xsl:variable>
            <xsl:variable name="name">
              <xsl:call-template name="unspace">
                <xsl:with-param name="string" select="$value"/>
              </xsl:call-template>
            </xsl:variable>
            <xsl:variable name="colorattribute">
              <xsl:call-template name="unspace">
                <xsl:with-param name="string">
                  <xsl:call-template name="colorattribute">
                    <xsl:with-param name="fillid">
                      <xsl:call-template name="fillid"/>
                    </xsl:with-param>
                  </xsl:call-template>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:variable>
            <th class=" xml cell header {$name} {$colorattribute} ">
              <xsl:value-of select="$value"/>
            </th>
          </xsl:for-each>
        </tr>
      </xsl:when>
      <xsl:otherwise>
        <xsl:variable name="row" select="."/>
        <tr class=" xml row ">
          <xsl:for-each select=".//ancestor::ss:worksheet//ss:row[count(preceding-sibling::ss:row) = $headerrow - 1]/ss:c">
            <xsl:variable name="r">
              <xsl:value-of select="substring-before(@r,../@r)"/>
              <xsl:value-of select="exslt:node-set($row)/@r"/>
            </xsl:variable>
            <xsl:variable name="name">
              <xsl:call-template name="unspace">
                <xsl:with-param name="string">
                  <xsl:call-template name="cellvalue"/>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:variable>
            <xsl:variable name="colorattribute">
              <xsl:call-template name="unspace">
                <xsl:with-param name="string">
                  <xsl:call-template name="colorattribute">
                    <xsl:with-param name="fillid">
                      <xsl:call-template name="fillid">
                        <xsl:with-param name="style" select="exslt:node-set($row)/ss:c[@r = $r]/@s"/>
                      </xsl:call-template>
                    </xsl:with-param>
                  </xsl:call-template>
                </xsl:with-param>
              </xsl:call-template>
            </xsl:variable>
            <td class=" xml cell data {$name} {$colorattribute} ">
              <xsl:call-template name="cellvalue">
                <xsl:with-param name="c" select="exslt:node-set($row)/ss:c[@r = $r]"/>
              </xsl:call-template>
            </td>
          </xsl:for-each> 
        </tr>    
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="unspace">
    <xsl:param name="string" select="."/>
    <xsl:value-of select="translate($string,' ','')"/>
  </xsl:template>

  <xsl:template name="cellvalue">
    <xsl:param name="c" select="."/>
    <xsl:choose>
      <xsl:when test="exslt:node-set($c)/@t = 's'">
        <xsl:value-of select="//div[@class = 'xlsx file'][@name = 'xl/sharedStrings.xml']//ss:si[exslt:node-set($c)/ss:v + 1]/ss:t/text()"/>
      </xsl:when>
      <xsl:otherwise>
        <xsl:value-of select="exslt:node-set($c)/ss:v/text()"/>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template name="colorattribute">
    <xsl:param name="fillid"/>
    <xsl:for-each select="//div[@class = 'xlsx file'][@name = concat('xl/worksheets/sheet',//ss:sheets/ss:sheet[@name = 'Colors']/@sheetId,'.xml')]//ss:c">
      <xsl:variable name="colorfillid">
        <xsl:call-template name="fillid"/>
      </xsl:variable>
      <xsl:if test="$colorfillid = $fillid">
        <xsl:call-template name="cellvalue"/>
      </xsl:if>
    </xsl:for-each>
  </xsl:template>

  <xsl:template name="fillid">
    <xsl:param name="style" select="@s"/>
    <xsl:value-of select="//div[@class = 'xlsx file'][@name = 'xl/styles.xml']//ss:cellXfs/ss:xf[count(preceding-sibling::ss:xf) = $style]/@fillId"/>
  </xsl:template>


</xsl:stylesheet>