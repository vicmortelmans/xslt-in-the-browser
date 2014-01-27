<?xml version="1.0"?>
<xsl:stylesheet version="1.0"
  xmlns:x="http://www.w3.org/1999/xhtml"
  xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
  xmlns:exslt="http://exslt.org/common"
  xmlns:msxsl="urn:schemas-microsoft-com:xslt"
  exclude-result-prefixes="exslt msxsl">

  <msxsl:script language="JScript" implements-prefix="exslt">
   this['node-set'] =  function (x) {
    return x;
    }
  </msxsl:script>
  
  <xsl:output method="html" indent="yes"/>

  <xsl:template match="/div">
    <!-- copy header row from the first table -->
    <!-- return the rows from all tables, sorted alphabetically -->
    <table>
      <xsl:apply-templates select="(//x:tr[contains(@class,'header')])[1]" mode="copy"/> 
      <xsl:apply-templates select="//x:tr[not(contains(@class,'header'))]" mode="copy">
        <xsl:sort select="x:td[contains(@class,'Article')]"/>
      </xsl:apply-templates>
    </table>
  </xsl:template>

  <xsl:template match="@*|node()">
    <xsl:apply-templates select="@*|node()"/>
  </xsl:template>

  <xsl:template match="@*|node()" mode="copy">
    <xsl:copy>
      <xsl:apply-templates select="@*|node()" mode="copy"/>
    </xsl:copy>
  </xsl:template>

</xsl:stylesheet>
