CREATE TABLE EDI_HL7_Params(
	ParameterId		varchar(50) NULL,
	ParameterValue	varchar(50) NULL
)
GO

INSERT INTO EDI_HL7_Params (ParameterId, ParameterValue) Values ('XSLT', 'hl7Body.xslt')
INSERT INTO EDI_HL7_Params (ParameterId, ParameterValue) Values ('Start', '11')
INSERT INTO EDI_HL7_Params (ParameterId, ParameterValue) Values ('End', '28')
INSERT INTO EDI_HL7_Params (ParameterId, ParameterValue) Values ('Escape', '92')
INSERT INTO EDI_HL7_Params (ParameterId, ParameterValue) Values ('Field', '124')
INSERT INTO EDI_HL7_Params (ParameterId, ParameterValue) Values ('Component', '94')
INSERT INTO EDI_HL7_Params (ParameterId, ParameterValue) Values ('SubComp', '38')
INSERT INTO EDI_HL7_Params (ParameterId, ParameterValue) Values ('Repetition', '126')
INSERT INTO EDI_HL7_Params (ParameterId, ParameterValue) Values ('Sender', 'ICE')
INSERT INTO EDI_HL7_Params (ParameterId, ParameterValue) Values ('ProcId', 'P')
INSERT INTO EDI_HL7_Params (ParameterId, ParameterValue) Values ('Version', '2.4')
INSERT INTO EDI_HL7_Params (ParameterId, ParameterValue) Values ('Country', 'uk')
