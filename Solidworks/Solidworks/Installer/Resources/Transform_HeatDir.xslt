<?xml version="1.0" ?>
<xsl:stylesheet version="1.0"
				xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
				xmlns:wix="http://schemas.microsoft.com/wix/2006/wi">

	<!-- Copy all attributes and elements to the output. -->
	<xsl:output method="xml"
				indent="yes"
				omit-xml-declaration="yes"/>
	<xsl:strip-space elements="*"/>

	<xsl:template match="@* | node()">
		<xsl:copy>
			<xsl:apply-templates select="@* | node()"/>
		</xsl:copy>
	</xsl:template>

	<!-- Create searches for the directories to remove. -->
	<xsl:key name="git-directory"
			 match="wix:Directory[contains(@Name,'.git')]"
			 use="@Id" />
	<xsl:key name="vs-directory"
			 match="wix:Directory[contains(@Name,'.vs')]"
			 use="@Id" />
	<xsl:key name="pdb-component"
			 match="wix:Component[contains(wix:File/@Source, '.pdb')]"
			 use="@Id" />

	<xsl:template match="*[self::wix:Directory]
						[key('git-directory', @Id)]" />
	<xsl:template match="*[self::wix:Directory]
						[key('vs-directory', @Id)]" />
	<xsl:template match="*[self::wix:Component or self::wix:ComponentRef]
						[key('pdb-component', @Id)]" />
</xsl:stylesheet>