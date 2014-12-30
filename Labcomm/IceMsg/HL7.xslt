<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:a="urn:hl7-org:v2xml">
<xsl:output method="text"/>
<xsl:template match="a:ORU_R01">
<xsl:apply-templates/>
<xsl:apply-templates select="ORU_R01"/>
<xsl:apply-templates select="ORU_R01.PATIENT_RESULT"/>
</xsl:template>

<xsl:template match="a:MSH">
<xsl:text>MSH</xsl:text><xsl:value-of select="a:MSH.1"/><xsl:value-of select="a:MSH.2"/>
	<xsl:value-of select="a:MSH.2/a:HD.1"/><xsl:text>|</xsl:text>
	<xsl:value-of select="a:MSH.3/a:HD.1"/><xsl:text>|</xsl:text>
	<xsl:value-of select="a:MSH.4/a:HD.1"/>^<xsl:value-of select="a:MSH.4/a:HD.2"/><xsl:text>||</xsl:text>
	<xsl:value-of select="a:MSH.6/a:HD.1"/><xsl:text>|</xsl:text>
	<xsl:value-of select="a:MSH.7/a:TS.1"/><xsl:text>||</xsl:text>
	<xsl:value-of select="a:MSH.9/a:MSG.1"/>^<xsl:value-of select="a:MSH.9/a:MSG.2"/><xsl:text>|</xsl:text>
	<xsl:value-of select="a:MSH.10"/><xsl:text>|</xsl:text>
	<xsl:value-of select="a:MSH.11/a:PT.1"/><xsl:text>|</xsl:text>
	<xsl:value-of select="a:MSH.12/a:VID.1"/>|
</xsl:template>

<xsl:template match="a:ORU_R01.PATIENT">
	<xsl:apply-templates select="a:PID"/>
	<xsl:apply-templates select="a:NTE"/>
	<xsl:apply-templates select="a:ORU_R01.PATIENT_VISIT"/>
	<xsl:apply-templates select="a:ORU_R01.ORDER_OBSERVATION"/>
</xsl:template>

<xsl:template match="a:PID">
<xsl:text>PID|</xsl:text><xsl:value-of select="a:PID.1"/><xsl:text>||</xsl:text>
	<xsl:value-of select="a:PID.3/a:CX.1"/><xsl:text>^</xsl:text>
	<xsl:value-of select="a:PID.3/a:CX.2"/><xsl:text>^</xsl:text>
	<xsl:value-of select="a:PID.3/a:CX.3"/><xsl:text>^</xsl:text>
	<xsl:value-of select="a:PID.3/a:CX.4"/><xsl:text>^</xsl:text>
	<xsl:value-of select="a:PID.3/a:CX.5"/><xsl:text>||</xsl:text>
	<xsl:value-of select="a:PID.5/a:XPN.1/a:FN.1"/>^<xsl:value-of select="a:PID.5/a:XPN.2"/>^<xsl:value-of select="a:PID.5/a:XPN.3"/>^<xsl:value-of select="a:PID.5/a:XPN.4"/>^<xsl:value-of select="a:PID.5/a:XPN.5"/>^<xsl:value-of select="a:PID.5/a:XPN.6"/>^<xsl:value-of select="a:PID.5/a:XPN.7"/><xsl:text>||</xsl:text>
	<xsl:value-of select="a:PID.7/a:TS.1"/><xsl:text>|</xsl:text>
	<xsl:value-of select="a:PID.8"/><xsl:text>||||</xsl:text>
	<xsl:value-of select="a:PID.11/a:XAD.1/a:SAD"/>^<xsl:value-of select="a:PID.11/a:XAD.2"/>^<xsl:value-of select="a:PID.11/a:XAD.3"/><xsl:text>|</xsl:text>
	<xsl:apply-templates select="a:NTE"/> 
</xsl:template>

<xsl:template match="a:PV1">
PV1|<xsl:value-of select="a:PV1.1"/><xsl:text>|</xsl:text>
<xsl:value-of select="a:PV1.2"/><xsl:text>|</xsl:text>
<xsl:value-of select="a:PV1.3/a:PL.4/a:HD.1"/><xsl:text>|</xsl:text>
<xsl:value-of select="a:PV1.4"/><xsl:text>|</xsl:text>
<xsl:value-of select="a:PV1.5"/><xsl:text>|</xsl:text>
<xsl:value-of select="a:PV1.6"/><xsl:text>|</xsl:text>
<xsl:value-of select="a:PV1.7/a:XCN.1"/><xsl:text>^</xsl:text>
<xsl:value-of select="a:PV1.7/a:XCN.2"/><xsl:text>|</xsl:text>
<xsl:value-of select="a:PV1.8/a:XCN.1"/><xsl:text>^</xsl:text>
<xsl:value-of select="a:PV1.8/a:XCN.2"/><xsl:text>|</xsl:text>
<xsl:text>|||||||||||</xsl:text><xsl:value-of select="a:PV1.19"/>
<xsl:apply-templates select="a:NTE"/>
</xsl:template>

<xsl:template match="a:OBR">
OBR|<xsl:value-of select="a:OBR.1"/><xsl:text>|</xsl:text>
<xsl:value-of select="a:OBR.2"/><xsl:text>|</xsl:text>
<xsl:value-of select="a:OBR.3/a:EI.1"/>^<xsl:value-of select="a:OBR.3/a:EI.2"/><xsl:text>|</xsl:text>
<xsl:apply-templates select="a:OBR.4/a:CE.1"/><xsl:text>^</xsl:text>
<xsl:apply-templates select="a:OBR.4/a:CE.2"/><xsl:text>^</xsl:text>
<xsl:apply-templates select="a:OBR.4/a:CE.3"/><xsl:text>^</xsl:text>
<xsl:apply-templates select="a:OBR.4/a:CE.4"/><xsl:text>^</xsl:text>
<xsl:apply-templates select="a:OBR.4/a:CE.5"/><xsl:text>^</xsl:text>
<xsl:apply-templates select="a:OBR.4/a:CE.6"/><xsl:text>|</xsl:text>
<xsl:value-of select="a:OBR.5"/><xsl:text>|</xsl:text>
<xsl:value-of select="a:OBR.6"/><xsl:text>|</xsl:text>
<xsl:value-of select="a:OBR.7/a:TS.1"/><xsl:text>|</xsl:text>
<xsl:value-of select="a:OBR.8"/><xsl:text>|</xsl:text>
<xsl:value-of select="a:OBR.9"/><xsl:text>|</xsl:text>
<xsl:value-of select="a:OBR.10"/><xsl:text>|</xsl:text>
<xsl:value-of select="a:OBR.11"/><xsl:text>|</xsl:text>
<xsl:value-of select="a:OBR.12"/><xsl:text>|</xsl:text>
<xsl:value-of select="a:OBR.13"/><xsl:text>|</xsl:text>
<xsl:value-of select="a:OBR.14/a:TS.1"/><xsl:text>|</xsl:text>
<xsl:value-of select="a:OBR.15/a:SPS.1/a:CE.1"/>^<xsl:value-of select="a:OBR.15/a:SPS.2"/><xsl:text>|</xsl:text>
<xsl:value-of select="a:OBR.16/a:XCN.1"/><xsl:text>^</xsl:text>
<xsl:value-of select="a:OBR.16/a:XCN.2/a:FN.1"/><xsl:text>|</xsl:text>

<xsl:value-of select="a:OBR.17"/><xsl:text>|</xsl:text>
<xsl:value-of select="a:OBR.18"/><xsl:text>|</xsl:text>
<xsl:value-of select="a:OBR.19"/><xsl:text>|</xsl:text>
<xsl:value-of select="a:OBR.20"/><xsl:text>|</xsl:text>
<xsl:value-of select="a:OBR.21"/><xsl:text>|</xsl:text>
<xsl:value-of select="a:OBR.22/a:TS.1"/><xsl:text>|</xsl:text>
<xsl:value-of select="a:OBR.23"/><xsl:text>|</xsl:text>
<xsl:value-of select="a:OBR.24"/><xsl:text>|</xsl:text>
<xsl:value-of select="a:OBR.25"/><xsl:text>|</xsl:text>
<xsl:value-of select="a:OBR.26"/><xsl:text>|</xsl:text>
<xsl:value-of select="a:OBR.27"/><xsl:text>|</xsl:text>
<xsl:apply-templates select="a:NTE"/>
<xsl:apply-templates select="a:ORU_R01.OBSERVATION"/>
</xsl:template>

<xsl:template match="a:OBX">
OBX|<xsl:value-of select="a:OBX.1"/><xsl:text>|</xsl:text>
<xsl:value-of select="a:OBX.2"/><xsl:text>|</xsl:text>
<xsl:apply-templates select="a:OBX.3/a:CE.1"/><xsl:text>^</xsl:text>
<xsl:apply-templates select="a:OBX.3/a:CE.2"/><xsl:text>^</xsl:text>
<xsl:apply-templates select="a:OBX.3/a:CE.3"/><xsl:text>^</xsl:text>
<xsl:apply-templates select="a:OBX.3/a:CE.4"/><xsl:text>^</xsl:text>
<xsl:apply-templates select="a:OBX.3/a:CE.5"/><xsl:text>^</xsl:text>
<xsl:apply-templates select="a:OBX.3/a:CE.6"/><xsl:text>|</xsl:text>

<xsl:value-of select="a:OBX.4"/><xsl:text>|</xsl:text>
<xsl:value-of select="a:OBX.5"/><xsl:text>|</xsl:text>
<xsl:value-of select="a:OBX.6/a:CE.1"/><xsl:text>|</xsl:text>
<xsl:value-of select="a:OBX.7"/><xsl:text>|</xsl:text>
<xsl:value-of select="a:OBX.8"/><xsl:text>|</xsl:text>
<xsl:value-of select="a:OBX.9"/><xsl:text>|</xsl:text>
<xsl:value-of select="a:OBX.10"/><xsl:text>|</xsl:text>
<xsl:value-of select="a:OBX.11"/><xsl:text>|</xsl:text>
<xsl:value-of select="a:OBX.12"/><xsl:text>|</xsl:text>
<xsl:value-of select="a:OBX.13"/><xsl:text>|</xsl:text>
<xsl:value-of select="a:OBX.14/a:TS.1"/><xsl:text>|</xsl:text>
<xsl:apply-templates select="a:NTE"/>
</xsl:template>

<xsl:template match="a:NTE">
NTE|<xsl:value-of select="a:NTE.1"/>|<xsl:value-of select="a:NTE.2"/>|<xsl:value-of select="a:NTE.3"/>
</xsl:template>


</xsl:stylesheet>

