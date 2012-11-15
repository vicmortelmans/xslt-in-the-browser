<?xml version="1.0"?>
<xsl:stylesheet version="1.0" xmlns="http://www.w3.org/1999/xhtml"
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:exslt="http://exslt.org/common"
  xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
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

  <xsl:template match="ss:Worksheet">
    <xsl:if test="@ss:Name = $worksheet">
      <table class="xml table">
        <xsl:apply-templates select=".//ss:Row[count(preceding-sibling::ss:Row) = $headerrow - 1]"/>
        <xsl:apply-templates select=".//ss:Row[count(preceding-sibling::ss:Row) &gt; $headerrow - 1]"/>
      </table>
    </xsl:if>
  </xsl:template>

  <xsl:template match="ss:Row">
    <xsl:choose>
      <xsl:when test="count(preceding-sibling::ss:Row) = $headerrow - 1">
        <tr class="xml row header">
          <xsl:for-each select="./ss:Cell">
            <xsl:variable name="name">
              <xsl:call-template name="unspace">
                <xsl:with-param name="string" select="."/>
              </xsl:call-template>
            </xsl:variable>
            <th class="xml cell header {$name}">
              <xsl:value-of select="."/>
            </th>
          </xsl:for-each>
        </tr>
      </xsl:when>
      <xsl:otherwise>
        <xsl:variable name="cells" select="./ss:Cell"/>
        <tr class="xml row">
          <xsl:for-each select="//ss:Worksheet[@ss:Name = $worksheet]//ss:Row[count(preceding-sibling::ss:Row) = $headerrow - 1]/ss:Cell">
            <xsl:variable name="name">
              <xsl:call-template name="unspace">
                <xsl:with-param name="string" select="."/>
              </xsl:call-template>
            </xsl:variable>
            <xsl:variable name="index" select="count(preceding-sibling::ss:Cell) + 1"/>
            <td class="xml cell data {$name}">
              <xsl:value-of select="exslt:node-set($cells)[(count(preceding-sibling::ss:Cell) + 1 = $index) or (@ss:Index and @ss:Index = $index)]"/>
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

</xsl:stylesheet>

