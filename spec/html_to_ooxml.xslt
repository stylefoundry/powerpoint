<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
                version="1.0"
                exclude-result-prefixes="java msxsl ext w o v WX aml w10">


  <xsl:output method="xml" encoding="utf-8" omit-xml-declaration="yes" indent="yes" />

  <xsl:template match="div[not(ancestor::td) and not(ancestor::th) and not(ancestor::p) and not(descendant::div) and not(descendant::p) and not(descendant::h1) and not(descendant::h2) and not(descendant::h3) and not(descendant::h4) and not(descendant::h5) and not(descendant::h6) and not(descendant::table) and not(descendant::li)]">
    <xsl:comment>Divs should create a p if nothing above them has and nothing below them will</xsl:comment>
    <a:p>
       <xsl:call-template name="text-alignment" />
       <xsl:apply-templates />
    </a:p>
  </xsl:template>

  <xsl:template match="div">
    <xsl:apply-templates />
  </xsl:template>

  <!-- TODO: make this prettier. Headings shouldn't enter in template from L51 -->
  <xsl:template match="body/h1|body/h2|body/h3|body/h4|body/h5|body/h6|h1|h2|h3|h4|h5|h6">
    <xsl:variable name="length" select="string-length(name(.))"/>
    <a:p>
      <a:r>
        <a:t xml:space="preserve"><xsl:value-of select="."/></a:t>
      </a:r>
    </a:p>
  </xsl:template>

  <xsl:template match="p">
    <a:p>
      <a:pPr marL="0" indent="0">
        <a:spcAft>
          <a:spcPts val="600"/>
        </a:spcAft>
        <a:buNone/>
      </a:pPr>
      <xsl:call-template name="text-alignment" />
      <xsl:apply-templates />
    </a:p>
  </xsl:template>

  <xsl:template match="li">
    <a:p>
      <a:r>
        <a:t xml:space="preserve"><xsl:value-of select="."/></a:t>
      </a:r>
    </a:p>
  </xsl:template>

  <xsl:template match="span[not(ancestor::td) and (preceding-sibling::h1 or preceding-sibling::h2 or preceding-sibling::h3 or preceding-sibling::h4 or preceding-sibling::h5 or preceding-sibling::h6 or preceding-sibling::table or preceding-sibling::p or preceding-sibling::ol or preceding-sibling::ul or preceding-sibling::div or following-sibling::h1 or following-sibling::h2 or following-sibling::h3 or following-sibling::h4 or following-sibling::h5 or following-sibling::h6 or following-sibling::table or following-sibling::p or following-sibling::ol or following-sibling::ul or following-sibling::div)]
    |a[not(ancestor::td) and (preceding-sibling::h1 or preceding-sibling::h2 or preceding-sibling::h3 or preceding-sibling::h4 or preceding-sibling::h5 or preceding-sibling::h6 or preceding-sibling::table or preceding-sibling::p or preceding-sibling::ol or preceding-sibling::ul or preceding-sibling::div or following-sibling::h1 or following-sibling::h2 or following-sibling::h3 or following-sibling::h4 or following-sibling::h5 or following-sibling::h6 or following-sibling::table or following-sibling::p or following-sibling::ol or following-sibling::ul or following-sibling::div)]
    |small[not(ancestor::td) and (preceding-sibling::h1 or preceding-sibling::h2 or preceding-sibling::h3 or preceding-sibling::h4 or preceding-sibling::h5 or preceding-sibling::h6 or preceding-sibling::table or preceding-sibling::p or preceding-sibling::ol or preceding-sibling::ul or preceding-sibling::div or following-sibling::h1 or following-sibling::h2 or following-sibling::h3 or following-sibling::h4 or following-sibling::h5 or following-sibling::h6 or following-sibling::table or following-sibling::p or following-sibling::ol or following-sibling::ul or following-sibling::div)]
    |strong[not(ancestor::td) and (preceding-sibling::h1 or preceding-sibling::h2 or preceding-sibling::h3 or preceding-sibling::h4 or preceding-sibling::h5 or preceding-sibling::h6 or preceding-sibling::table or preceding-sibling::p or preceding-sibling::ol or preceding-sibling::ul or preceding-sibling::div or following-sibling::h1 or following-sibling::h2 or following-sibling::h3 or following-sibling::h4 or following-sibling::h5 or following-sibling::h6 or following-sibling::table or following-sibling::p or following-sibling::ol or following-sibling::ul or following-sibling::div)]
    |em[not(ancestor::td) and (preceding-sibling::h1 or preceding-sibling::h2 or preceding-sibling::h3 or preceding-sibling::h4 or preceding-sibling::h5 or preceding-sibling::h6 or preceding-sibling::table or preceding-sibling::p or preceding-sibling::ol or preceding-sibling::ul or preceding-sibling::div or following-sibling::h1 or following-sibling::h2 or following-sibling::h3 or following-sibling::h4 or following-sibling::h5 or following-sibling::h6 or following-sibling::table or following-sibling::p or following-sibling::ol or following-sibling::ul or following-sibling::div)]
    |i[not(ancestor::td) and (preceding-sibling::h1 or preceding-sibling::h2 or preceding-sibling::h3 or preceding-sibling::h4 or preceding-sibling::h5 or preceding-sibling::h6 or preceding-sibling::table or preceding-sibling::p or preceding-sibling::ol or preceding-sibling::ul or preceding-sibling::div or following-sibling::h1 or following-sibling::h2 or following-sibling::h3 or following-sibling::h4 or following-sibling::h5 or following-sibling::h6 or following-sibling::table or following-sibling::p or following-sibling::ol or following-sibling::ul or following-sibling::div)]
    |b[not(ancestor::td) and (preceding-sibling::h1 or preceding-sibling::h2 or preceding-sibling::h3 or preceding-sibling::h4 or preceding-sibling::h5 or preceding-sibling::h6 or preceding-sibling::table or preceding-sibling::p or preceding-sibling::ol or preceding-sibling::ul or preceding-sibling::div or following-sibling::h1 or following-sibling::h2 or following-sibling::h3 or following-sibling::h4 or following-sibling::h5 or following-sibling::h6 or following-sibling::table or following-sibling::p or following-sibling::ol or following-sibling::ul or following-sibling::div)]
    |u[not(ancestor::td) and (preceding-sibling::h1 or preceding-sibling::h2 or preceding-sibling::h3 or preceding-sibling::h4 or preceding-sibling::h5 or preceding-sibling::h6 or preceding-sibling::table or preceding-sibling::p or preceding-sibling::ol or preceding-sibling::ul or preceding-sibling::div or following-sibling::h1 or following-sibling::h2 or following-sibling::h3 or following-sibling::h4 or following-sibling::h5 or following-sibling::h6 or following-sibling::table or following-sibling::p or following-sibling::ol or following-sibling::ul or following-sibling::div)]">
    <xsl:comment>
        In the following situation:

        div
          h2
          span
            textnode
            span
              textnode
          p

        The div template will not create a a:p because the div contains a h2. Therefore we need to wrap the inline elements span|a|small in a p here.
      </xsl:comment>
    <a:p>
      <xsl:apply-templates />
    </a:p>
  </xsl:template>

  <xsl:template match="text()[not(parent::td) and (preceding-sibling::h1 or preceding-sibling::h2 or preceding-sibling::h3 or preceding-sibling::h4 or preceding-sibling::h5 or preceding-sibling::h6 or preceding-sibling::table or preceding-sibling::p or preceding-sibling::ol or preceding-sibling::ul or preceding-sibling::div or following-sibling::h1 or following-sibling::h2 or following-sibling::h3 or following-sibling::h4 or following-sibling::h5 or following-sibling::h6 or following-sibling::table or following-sibling::p or following-sibling::ol or following-sibling::ul or following-sibling::div)]">
    <xsl:comment>
        In the following situation:

        div
          h2
          textnode
          p

        The div template will not create a a:p because the div contains a h2. Therefore we need to wrap the textnode in a p here.
      </xsl:comment>
    <a:p>
      <a:r>        
        <a:rPr lang="en-GB" sz="1800" dirty="0" smtClean="0"/>
        <a:t xml:space="preserve"><xsl:value-of select="."/></a:t>
      </a:r>
    </a:p>
  </xsl:template>

  <xsl:template match="span[contains(concat(' ', @class, ' '), ' h ')]">
    <xsl:comment>
        This template adds MS Word highlighting ability.
      </xsl:comment>
    <xsl:variable name="color">
      <xsl:choose>
        <xsl:when test="./@data-style='pink'">magenta</xsl:when>
        <xsl:when test="./@data-style='blue'">cyan</xsl:when>
        <xsl:when test="./@data-style='orange'">darkYellow</xsl:when>
        <xsl:otherwise><xsl:value-of select="./@data-style"/></xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:choose>
      <xsl:when test="preceding-sibling::h1 or preceding-sibling::h2 or preceding-sibling::h3 or preceding-sibling::h4 or preceding-sibling::h5 or preceding-sibling::h6 or preceding-sibling::table or preceding-sibling::p or preceding-sibling::ol or preceding-sibling::ul or preceding-sibling::div or following-sibling::h1 or following-sibling::h2 or following-sibling::h3 or following-sibling::h4 or following-sibling::h5 or following-sibling::h6 or following-sibling::table or following-sibling::p or following-sibling::ol or following-sibling::ul or following-sibling::div">
        <a:p>
          <a:r>
            <a:rPr>
              <a:highlight a:val="{$color}"/>
            </a:rPr>
            <a:t xml:space="preserve"><xsl:value-of select="."/></a:t>
          </a:r>
        </a:p>
      </xsl:when>
      <xsl:otherwise>
        <a:r>
          <a:rPr>
            <a:highlight a:val="{$color}"/>
          </a:rPr>
          <a:t xml:space="preserve"><xsl:value-of select="."/></a:t>
        </a:r>
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template match="div[contains(concat(' ', @class, ' '), ' -page-break ')]">
    <a:p>
      <a:r>
        <a:br a:type="page" />
      </a:r>
    </a:p>
    <xsl:apply-templates />
  </xsl:template>

  <xsl:template match="details" />

  <xsl:template name="tableborders">
    <xsl:variable name="border">
      <xsl:choose>
        <xsl:when test="contains(concat(' ', @class, ' '), ' table-bordered ')">6</xsl:when>
        <xsl:when test="not(@border)">0</xsl:when>
        <xsl:otherwise><xsl:value-of select="./@border * 6"/></xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:variable name="bordertype">
      <xsl:choose>
        <xsl:when test="$border=0">none</xsl:when>
        <xsl:otherwise>single</xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <a:tblBorders>
      <a:top a:val="{$bordertype}" a:sz="{$border}" a:space="0" a:color="auto"/>
      <a:left a:val="{$bordertype}" a:sz="{$border}" a:space="0" a:color="auto"/>
      <a:bottom a:val="{$bordertype}" a:sz="{$border}" a:space="0" a:color="auto"/>
      <a:right a:val="{$bordertype}" a:sz="{$border}" a:space="0" a:color="auto"/>
      <a:insideH a:val="{$bordertype}" a:sz="{$border}" a:space="0" a:color="auto"/>
      <a:insideV a:val="{$bordertype}" a:sz="{$border}" a:space="0" a:color="auto"/>
    </a:tblBorders>
  </xsl:template>

  <xsl:template match="table">
    <a:tbl>
      <a:tblPr>
        <a:tblStyle a:val="TableGrid"/>
        <a:tblW a:w="0" a:type="auto"/>
        <xsl:call-template name="tableborders"/>
        <a:tblLook a:val="0600" a:firstRow="0" a:lastRow="0" a:firstColumn="0" a:lastColumn="0" a:noHBand="1" a:noVBand="1"/>
      </a:tblPr>
      <a:tblGrid>
        <a:gridCol a:w="2310"/>
        <a:gridCol a:w="2310"/>
      </a:tblGrid>
      <xsl:apply-templates />
    </a:tbl>
  </xsl:template>

  <xsl:template match="tbody">
    <xsl:apply-templates />
  </xsl:template>

  <xsl:template match="thead">
    <xsl:choose>
      <xsl:when test="count(./tr) = 0">
        <a:tr><xsl:apply-templates /></a:tr>
      </xsl:when>
      <xsl:otherwise>
        <xsl:apply-templates />
      </xsl:otherwise>
    </xsl:choose>
  </xsl:template>

  <xsl:template match="tr">
    <xsl:if test="string-length(.) > 0">
      <a:tr>
        <xsl:apply-templates />
      </a:tr>
    </xsl:if>
  </xsl:template>

  <xsl:template match="th">
    <a:tc>
      <a:p>
        <a:r>
          <a:rPr>
            <a:b />
          </a:rPr>
          <a:t xml:space="preserve"><xsl:value-of select="."/></a:t>
        </a:r>
      </a:p>
    </a:tc>
  </xsl:template>

  <xsl:template match="td">
    <a:tc>
      <xsl:call-template name="block">
        <xsl:with-param name="current" select="." />
        <xsl:with-param name="class" select="@class" />
        <xsl:with-param name="style" select="@style" />
      </xsl:call-template>
    </a:tc>
  </xsl:template>

  <xsl:template name="block">
    <xsl:param name="current" />
    <xsl:param name="class" />
    <xsl:param name="style" />
    <xsl:if test="count($current/*|$current/text()) = 0">
      <a:p/>
    </xsl:if>
    <xsl:for-each select="$current/*|$current/text()">
      <xsl:choose>
        <xsl:when test="name(.) = 'table'">
          <xsl:apply-templates select="." />
          <a:p/>
        </xsl:when>
        <xsl:when test="contains('|p|h1|h2|h3|h4|h5|h6|ul|ol|', concat('|', name(.), '|'))">
          <xsl:apply-templates select="." />
        </xsl:when>
        <xsl:when test="descendant::table|descendant::p|descendant::h1|descendant::h2|descendant::h3|descendant::h4|descendant::h5|descendant::h6|descendant::li">
          <xsl:call-template name="block">
            <xsl:with-param name="current" select="."/>
          </xsl:call-template>
        </xsl:when>
        <xsl:otherwise>
          <a:p>
            <xsl:call-template name="text-alignment">
              <xsl:with-param name="class" select="$class" />
              <xsl:with-param name="style" select="$style" />
            </xsl:call-template>
            <xsl:apply-templates select="." />
          </a:p>
        </xsl:otherwise>
      </xsl:choose>
    </xsl:for-each>
  </xsl:template>

  <xsl:template match="text()">
    <xsl:if test="string-length(.) > 0">
      <a:r>
        <xsl:if test="ancestor::i or ancestor::em">
          <a:rPr>
            <a:i />
          </a:rPr>
        </xsl:if>
        <xsl:if test="ancestor::b or ancestor::strong">
          <a:rPr>
            <a:b />
          </a:rPr>
        </xsl:if>
        <xsl:if test="ancestor::u">
          <a:rPr>
            <a:u a:val="single"/>
          </a:rPr>
        </xsl:if>
        <a:rPr lang="en-GB" sz="1800" dirty="0" smtClean="0"/>
        <a:t xml:space="preserve"><xsl:value-of select="."/></a:t>
      </a:r>
    </xsl:if>
  </xsl:template>

  <xsl:template match="*">
    <xsl:apply-templates/>
  </xsl:template>

  <xsl:template name="text-alignment">
    <xsl:param name="class" select="@class" />
    <xsl:param name="style" select="@style" />
    <xsl:variable name="alignment">
      <xsl:choose>
        <xsl:when test="contains(concat(' ', $class, ' '), ' center ') or contains(translate(normalize-space($style),' ',''), 'text-align:center')">center</xsl:when>
        <xsl:when test="contains(concat(' ', $class, ' '), ' right ') or contains(translate(normalize-space($style),' ',''), 'text-align:right')">right</xsl:when>
        <xsl:when test="contains(concat(' ', $class, ' '), ' left ') or contains(translate(normalize-space($style),' ',''), 'text-align:left')">left</xsl:when>
        <xsl:when test="contains(concat(' ', $class, ' '), ' justify ') or contains(translate(normalize-space($style),' ',''), 'text-align:justify')">both</xsl:when>
        <xsl:otherwise></xsl:otherwise>
      </xsl:choose>
    </xsl:variable>
    <xsl:if test="string-length(normalize-space($alignment)) > 0">
      <a:pPr>
        <a:jc a:val="{$alignment}"/>
      </a:pPr>
    </xsl:if>
  </xsl:template>
</xsl:stylesheet>