SELECT DISTINCT sh.Date_Added, sh.Service_ImpExp_Id, ImpExp_File, Status_Flag, Warning_Flag, Error_Found FROM Service_ImpExp_HeadeRS sh LEFT JOIN Service_ImpExp_Messages sm ON sh.Service_ImpExp_Id = sm.Service_ImpExp_Id WHERE Service_Type = 1 AND (Case When sm.EDI_LTS_Index is null Then sh.EDI_LTS_Index When sh.EDI_LTS_Index=0 then 0 When sm.EDI_LTS_Index = 0 then sh.EDI_LTS_Index When sm.EDI_LTS_Index <> sh.EDI_LTS_Index then sm.EDI_LTS_Index Else sm.EDI_LTS_Index End) = 1  AND (datepart(yyyy, sh.Date_Added)='2007'  AND datepart(m, sh.Date_Added) = '7'  AND datepart(d, sh.Date_Added) = '12') ORDER BY sh.Service_ImpExp_Id