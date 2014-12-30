Attribute VB_Name = "mGuid"
Option Explicit

Private Type Guid
Data1 As Long
Data2 As Integer
Data3 As Integer
Data4(7) As Byte
End Type

Private Declare Function CoCreateGuid Lib "OLE32.DLL" (pGuid As Guid) As Long

' See http://support.microsoft.com/kb/176790
Public Function GetGuid() As String

    Dim udtGUID As Guid

    If (CoCreateGuid(udtGUID) = 0) Then
    GetGuid = _
        String(8 - Len(Hex$(udtGUID.Data1)), "0") & Hex$(udtGUID.Data1) & _
        String(4 - Len(Hex$(udtGUID.Data2)), "0") & Hex$(udtGUID.Data2) & _
        String(4 - Len(Hex$(udtGUID.Data3)), "0") & Hex$(udtGUID.Data3) & _
        IIf((udtGUID.Data4(0) < &H10), "0", "") & Hex$(udtGUID.Data4(0)) & _
        IIf((udtGUID.Data4(1) < &H10), "0", "") & Hex$(udtGUID.Data4(1)) & _
        IIf((udtGUID.Data4(2) < &H10), "0", "") & Hex$(udtGUID.Data4(2)) & _
        IIf((udtGUID.Data4(3) < &H10), "0", "") & Hex$(udtGUID.Data4(3)) & _
        IIf((udtGUID.Data4(4) < &H10), "0", "") & Hex$(udtGUID.Data4(4)) & _
        IIf((udtGUID.Data4(5) < &H10), "0", "") & Hex$(udtGUID.Data4(5)) & _
        IIf((udtGUID.Data4(6) < &H10), "0", "") & Hex$(udtGUID.Data4(6)) & _
        IIf((udtGUID.Data4(7) < &H10), "0", "") & Hex$(udtGUID.Data4(7))
    End If

End Function

Public Function FormatGuid(ByVal UnformattedGuid As String) As String
    
    Debug.Assert Len(UnformattedGuid) = 32
    
    FormatGuid = _
        Mid$(UnformattedGuid, 1, 8) & "-" & _
        Mid$(UnformattedGuid, 9, 4) & "-" & _
        Mid$(UnformattedGuid, 13, 4) & "-" & _
        Mid$(UnformattedGuid, 17, 4) & "-" & _
        Mid$(UnformattedGuid, 21, 12)
        
    Debug.Assert Len(FormatGuid) = 36
        
End Function

Public Function UnformatGuid(ByVal FormattedGuid As String) As String

    UnformatGuid = Replace(FormattedGuid, "-", "")

End Function
