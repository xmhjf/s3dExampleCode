Attribute VB_Name = "CommonEnumsAndConstants"
'*******************************************************************
'
'Copyright (C) 2015 Intergraph Corporation. All rights reserved.
'
'File : CommonEnumsAndConstants.bas
'
'Author : PYK
'
'Description :
'    Common Constants useful for plant and marine content
'
'History:
'
' 11/12/15   PYK    Created bas file for Common constants and Enums [TR-CP-283069].

'********************************************************************
Option Explicit

Public Const E_FAIL = -2147467259

' Different types and type categories of members are defined as Enums
Public Enum MemberCategoryAndType
  HandRailElement = 6
    HRPost = 602
    HRAttachment = 607
    HRTopRail = 603
  
  StairElement = 7
    StrStringer = 703
    StrSupport = 704
    StrAttachment = 706
  
  LadderElement = 8
    LElement = 801
    LRung = 802
    LRail = 803
    LHoop = 804
    LAttachment = 805
End Enum
