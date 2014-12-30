<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
<xsl:output method="text"/>
<xsl:template match="ORU_R01">
<xsl:apply-templates select="ORU_R01"/>
<xsl:apply-templates select="ORU_R01.PATIENT_RESULT"/>
</xsl:template>

<xsl:template match="ORU_R01.PATIENT">
	<xsl:apply-templates select="PID"/>
	<xsl:apply-templates select="NTE"/>
	<xsl:apply-templates select="ORU_R01.PATIENT_VISIT"/>
	<xsl:apply-templates select="ORU_R01.ORDER_OBSERVATION"/>
</xsl:template>

<xsl:template match="PID">
PID|<xsl:value-of select="PID.1"/><xsl:text>||</xsl:text>
	<xsl:value-of select="PID.3/PD.1/CX.1"/><xsl:text>^^^</xsl:text>
	<xsl:value-of select="PID.3/PD.1/CX.4"/><xsl:text>^</xsl:text>
	<xsl:value-of select="PID.3/PD.1/CX.5"/><xsl:text>~</xsl:text>
	<xsl:value-of select="PID.3/PD.2/CX.1"/><xsl:text>^^^</xsl:text>
	<xsl:value-of select="PID.3/PD.2/CX.4"/><xsl:text>^</xsl:text>
	<xsl:value-of select="PID.3/PD.2/CX.5"/><xsl:text>|</xsl:text>
	<xsl:value-of select="PID.3/PD.3/CX.1"/><xsl:text>|</xsl:text>
	<xsl:value-of select="PID.5/XPN.1/FN.1"/>^<xsl:value-of select="PID.5/XPN.2"/>^<xsl:value-of select="PID.5/XPN.3"/>^<xsl:value-of select="PID.5/XPN.4"/>^<xsl:value-of select="PID.5/XPN.5"/>^<xsl:value-of select="PID.5/XPN.6"/>^<xsl:value-of select="PID.5/XPN.7"/><xsl:text>||</xsl:text>
	<xsl:value-of select="PID.7/TS.1"/><xsl:text>|</xsl:text>
	<xsl:value-of select="PID.8"/><xsl:text>|||</xsl:text>
	<xsl:value-of select="PID.11/XAD.1/SAD.1"/>^<xsl:value-of select="PID.11/XAD.2"/>^<xsl:value-of select="PID.11/XAD.3"/><xsl:text>|</xsl:text><xsl:value-of select="PID.12"/>
	<xsl:apply-templates select="NTE"/> 
</xsl:template>

<xsl:template match="PV1">
PV1|<xsl:value-of select="PV1.1"/><xsl:text>|</xsl:text>
<xsl:value-of select="PV1.2"/><xsl:text>|</xsl:text>
<xsl:value-of select="PV1.3/PL.4/HD.2"/><xsl:text>|</xsl:text>
<xsl:value-of select="PV1.4"/><xsl:text>|</xsl:text>
<xsl:value-of select="PV1.5"/><xsl:text>|</xsl:text>
<xsl:value-of select="PV1.6"/><xsl:text>|</xsl:text>
<xsl:value-of select="PV1.7/XCN.1"/><xsl:text>^</xsl:text>
<xsl:value-of select="PV1.7/XCN.2"/><xsl:text>|</xsl:text>
<xsl:value-of select="PV1.8/XCN.1"/><xsl:text>^</xsl:text>
<xsl:value-of select="PV1.8/XCN.2"/><xsl:text>|</xsl:text>
<xsl:text>|||||||||||</xsl:text><xsl:value-of select="PV1.19"/>
<xsl:apply-templates select="NTE"/>
</xsl:template>

<xsl:template match="OBR">
OBR|<xsl:value-of select="OBR.1"/><xsl:text>|</xsl:text>
<xsl:value-of select="OBR.2"/><xsl:text>|</xsl:text>
<xsl:value-of select="OBR.3/EI.1"/>^<xsl:value-of select="OBR.3/EI.2"/><xsl:text>|</xsl:text>
<xsl:apply-templates select="OBR.4/CE.1"/><xsl:text>^</xsl:text>
<xsl:apply-templates select="OBR.4/CE.2"/><xsl:text>^</xsl:text>
<xsl:apply-templates select="OBR.4/CE.3"/><xsl:text>^</xsl:text>
<xsl:apply-templates select="OBR.4/CE.4"/><xsl:text>^</xsl:text>
<xsl:apply-templates select="OBR.4/CE.5"/><xsl:text>^</xsl:text>
<xsl:apply-templates select="OBR.4/CE.6"/><xsl:text>|</xsl:text>
<xsl:value-of select="OBR.5"/><xsl:text>|</xsl:text>
<xsl:value-of select="OBR.6"/><xsl:text>|</xsl:text>
<xsl:value-of select="OBR.7/TS.1"/><xsl:text>|</xsl:text>
<xsl:value-of select="OBR.8"/><xsl:text>|</xsl:text>
<xsl:value-of select="OBR.9"/><xsl:text>|</xsl:text>
<xsl:value-of select="OBR.10"/><xsl:text>|</xsl:text>
<xsl:value-of select="OBR.11"/><xsl:text>|</xsl:text>
<xsl:value-of select="OBR.12"/><xsl:text>|</xsl:text>
<xsl:value-of select="OBR.13"/><xsl:text>|</xsl:text>
<xsl:value-of select="OBR.14/TS.1"/><xsl:text>|</xsl:text>
<xsl:value-of select="OBR.15/SPS.1/CE.1"/>^<xsl:value-of select="OBR.15/SPS.2"/><xsl:text>|</xsl:text>
<xsl:value-of select="OBR.16/XCN.1"/><xsl:text>^</xsl:text>
<xsl:value-of select="OBR.16/XCN.2/FN.1"/><xsl:text>|</xsl:text>

<xsl:value-of select="OBR.17"/><xsl:text>|</xsl:text>
<xsl:value-of select="OBR.18"/><xsl:text>|</xsl:text>
<xsl:value-of select="OBR.19"/><xsl:text>|</xsl:text>
<xsl:value-of select="OBR.20"/><xsl:text>|</xsl:text>
<xsl:value-of select="OBR.21"/><xsl:text>|</xsl:text>
<xsl:value-of select="OBR.22/TS.1"/><xsl:text>|</xsl:text>
<xsl:value-of select="OBR.23"/><xsl:text>|</xsl:text>
<xsl:value-of select="OBR.24"/><xsl:text>|</xsl:text>
<xsl:value-of select="OBR.25"/><xsl:text>|</xsl:text>
<xsl:value-of select="OBR.26"/><xsl:text>|</xsl:text>
<xsl:value-of select="OBR.27"/><xsl:text>|</xsl:text>
<xsl:apply-templates select="NTE"/>
<xsl:apply-templates select="ORU_R01.OBSERVATION"/>
</xsl:template>

<xsl:template match="OBX">
OBX|<xsl:value-of select="OBX.1"/><xsl:text>|</xsl:text>
<xsl:value-of select="OBX.2"/><xsl:text>|</xsl:text>
<xsl:apply-templates select="OBX.3/CE.1"/><xsl:text>^</xsl:text>
<xsl:apply-templates select="OBX.3/CE.2"/><xsl:text>^</xsl:text>
<xsl:apply-templates select="OBX.3/CE.3"/><xsl:text>^</xsl:text>
<xsl:apply-templates select="OBX.3/CE.4"/><xsl:text>^</xsl:text>
<xsl:apply-templates select="OBX.3/CE.5"/><xsl:text>^</xsl:text>
<xsl:apply-templates select="OBX.3/CE.6"/><xsl:text>|</xsl:text>

<xsl:value-of select="OBX.4"/><xsl:text>|</xsl:text>
<xsl:value-of select="OBX.5"/><xsl:text>|</xsl:text>
<xsl:value-of select="OBX.6/CE.1"/><xsl:text>|</xsl:text>
<xsl:value-of select="OBX.7"/><xsl:text>|</xsl:text>
<xsl:value-of select="OBX.8"/><xsl:text>|</xsl:text>
<xsl:value-of select="OBX.9"/><xsl:text>|</xsl:text>
<xsl:value-of select="OBX.10"/><xsl:text>|</xsl:text>
<xsl:value-of select="OBX.11"/><xsl:text>|</xsl:text>
<xsl:value-of select="OBX.12"/><xsl:text>|</xsl:text>
<xsl:value-of select="OBX.13"/><xsl:text>|</xsl:text>
<xsl:value-of select="OBX.14/TS.1"/><xsl:text>|</xsl:text>
<xsl:apply-templates select="NTE"/>
</xsl:template>

<xsl:template match="NTE">
NTE|<xsl:value-of select="NTE.1"/>|<xsl:value-of select="NTE.2"/>|<xsl:value-of select="NTE.3"/>
</xsl:template>

</xsl:stylesheet>

