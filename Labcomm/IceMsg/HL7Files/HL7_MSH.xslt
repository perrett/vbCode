<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
<xsl:output method="text"/>

<xsl:template name="MSH">
<xsl:apply-templates/>
</xsl:template >

<xsl:template match="/">
<xsl:text>MSH</xsl:text><xsl:value-of select="MSH.1"/><xsl:value-of select="MSH.2"/>
	<xsl:value-of select="MSH.2/HD.1"/><xsl:text>|</xsl:text>
	<xsl:value-of select="MSH.3/HD.1"/><xsl:text>_</xsl:text><xsl:value-of select="MSH.3/HD.2"/><xsl:text>|</xsl:text>
	<xsl:value-of select="MSH.4/HD.1"/><xsl:text>||</xsl:text>
	<xsl:value-of select="MSH.6/HD.1"/><xsl:text>|</xsl:text>
	<xsl:value-of select="MSH.7/TS.1"/><xsl:text>||</xsl:text>
	<xsl:value-of select="MSH.9/MSG.1"/>^<xsl:value-of select="MSH.9/MSG.2"/><xsl:text>|</xsl:text>
	<xsl:value-of select="MSH.10"/><xsl:text>|</xsl:text>
	<xsl:value-of select="MSH.11/PT.1"/><xsl:text>|</xsl:text>
	<xsl:value-of select="MSH.12/VID.1"/><xsl:text>|</xsl:text>
</xsl:template>

</xsl:stylesheet>
