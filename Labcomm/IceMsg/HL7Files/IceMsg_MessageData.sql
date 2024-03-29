USE [ICE]
GO
/****** Object:  StoredProcedure [dbo].[ICEMSG_MessageData]    Script Date: 10/07/2010 15:48:30 ******/
SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO
ALTER  PROCEDURE [dbo].[ICEMSG_MessageData]
@LTIndex int,
@NatCode varchar(6),
@Spec varchar(6)
AS
SELECT em.*,
	File_Extension,
	EDI_OrgCode,
	RTrim(Organisation_Name) AS Organisation_Name,
	RTrim(er.EDI_Name) AS EDI_Name,
	EDI_Encryption,
	EDI_Specialty,
	EDI_Hold_Output,
	el.EDI_LTS_Index,
	err.Ref_Index,
	err.Link_Interchange_Nos,
	err.EDI_Trader_Account + err.EDI_Free_Part AS ReceiverID,
	EDI_Trader_Code + el.EDI_Free_Part AS SenderId,
	Case
		When UseLabReadCodes is Null Then 0
		Else UseLabReadCodes 
	End as UseLabReadCodes,
	ein.EDI_Last_Interchange as Last_Interchange,
	Conformance_Frequency,
	Conformance_Nat_Code,
	Conformance_Total
FROM EDI_Msg_Types em
	INNER JOIN EDI_Msg_Formats
	ON EDI_Msg_Format = Type + ',' + Version

	INNER JOIN Organisation
	ON Organisation = Organisation_National_Code

	INNER JOIN EDI_Recipients er
	ON EDI_Org_NatCode = EDI_NatCode

	INNER JOIN EDI_Recipient_Ref err
		INNER JOIN EDI_Interchange_No ein
		ON err.Ref_Index = ein.Ref_Index
	ON er.Ref_Index= err.Ref_Index

	INNER JOIN EDI_Local_Trader_Settings el
	ON el.EDI_LTS_Index = @LTIndex

	INNER JOIN EDI_Loc_Specialties es
	ON er.EDI_NatCode = es.EDI_Nat_Code
		AND em.EDI_Msg_Format = es.EDI_Msg_Format
		AND em.Organisation = es.Organisation
WHERE EDI_Org_NatCode = @NatCode
	AND es.EDI_LTS_Index = @LTIndex
	AND EDI_Korner_Code = @Spec
