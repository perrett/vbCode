Attribute VB_Name = "StatusFlags"
Option Explicit

Public Enum enumInvResStatus
   IS_RC_NONE = &H1
   IS_RC_NA = &H2
   IS_RC_DELETED = &H4
   IS_RC_REMOVED = &H8
   IS_TEST_SUPPRESSED = &H20010
   IS_TEST_INACTIVE = &H20
   IS_INV_SUPPRESSED = &H20040
   IS_INV_INACTIVE = &H80
   IS_SAMPLE_SUPPRESSED = &H20100
   IS_NO_ACK = &H100000
   IS_PARSE_FAIL = &H2000
End Enum

Public Enum enumReportStatus
   RS_ACK_RECEIVED = &H100
   RS_WARNING_LOGGED = &H800
   RS_CONFORMANCE = &H1000
   RS_PARSE_FAIL = &H2000
   RS_ACK_ERROR = &H4000
   RS_READ_CODE_COMMENT = &H8000
   RS_GENERAL = &H10000
   RS_SUPPRESSION = &H20000
   RS_DATA_INTEGRITY = &H40000
End Enum

Public Enum enumMessageStatus
   MS_HTML = &H1
   MS_ASTM = &H2
   MS_RSR = &H4
   MS_EDI2 = &H8
   MS_EDI3 = &H10
   MS_HL7 = &H20
   MS_XML = &H40
   MS_DSCH = &H80
   MS_ACK_RECEIVED = &H100
   MS_ACK_REJECT_ALL = &H200
   MS_ACK_REJECT_PART = &H400
   MS_ACK_CRYPTO = &H800
   MS_CONFORMANCE = &H1000
   MS_PARSE_FAIL = 32768   '  &H8000
   MS_ACK_FAIL = &H10000
   MS_NO_OUTPUT = &H20000
   MS_DATA_INTEGRITY = &H40000
   MS_DTS_FAIL = &H80000
   MS_CRYPT_FAIL = &H100000
   MS_REQUEUE = &H10000000
   MS_ACK_TO_LIMS = &H20000000
   MS_AWAIT_ACK = &H40000000
   MS_MSGOK = &H80000000
End Enum

