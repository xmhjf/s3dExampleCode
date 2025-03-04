Attribute VB_Name = "MfgFilterGuids"
'----------------------------------------------------------------------------
' Copyright (C) 2000-2001, Intergraph Corporation.  All rights reserved.
'
' Project:
'    GSCAD structure manufactoring
'
' File:
'   Guids.bas
'
' Description:
'   Definition of GUIDs used in GSCAD Structural Manufacturing
'   (included the GSCAD Define Shrinkage application)
'
' Programmed by:
'
'
' History:
'   2001-01-16  oss\och                         Creation
'   22.01.2000  Alex Tchipanini(Maersk data)    Creation of Shrinkage version
'   02.11.2001  oss\pfb                         Merging of the two files
'----------------------------------------------------------------------------


'----------------------------------------------------------------------------
' Declarations


Option Explicit

'----------------------------------------------------------------------------
' Planning environment
'
Public Const IID_IJAEBoundVolume As String = "{6FEC5F51-72F4-11D2-8A6A-00A0C9065DF6}"
Public Const IID_IJAssembly As String = "{B447C9B6-FB74-11D1-8A49-00A0C9065DF6}"
Public Const IID_IJAssemblyChild As String = "{B447C9B4-FB74-11D1-8A49-00A0C9065DF6}"
Public Const IID_IJBlock As String = "{F4EC2FE1-0FF3-11D2-9A02-00A0C94472D6}"
Public Const IID_IJIntersection As String = "{FC359E53-119E-11D2-8A4C-00A0C9065DF6}"
Public Const IID_IJMergeVolume As String = "{3CE6CF16-A669-11D3-8077-0090276F4295}"
Public Const IID_IJPlaneSurface As String = "{1DA75D11-0FEB-11D2-8A4C-00A0C9065DF6}"
Public Const IID_IJSplitVolume As String = "{9DEBEE34-5C2E-11D2-B01A-00A0C9295A96}"
Public Const IID_IJTrimVolume As String = "{D5CC4516-A6FB-11D3-8078-0090276F4295}"
Public Const IID_IJVolume As String = "{5247C240-1010-11D2-AFD0-00A0C9295A96}"
Public Const IID_IJVolumeNotify As String = "{EAC0FD4A-286E-11d3-8051-0090276F429E}"
Public Const IID_IJConstMargin As String = "{6A889585-29A6-11D5-BFE5-00902770756B}"

'----------------------------------------------------------------------------
' StrMfg environment
'
Public Const IID_IJMfgPlateInit_AE As String = "{A859F29D-65A0-11D5-815E-0090276F4297}"
Public Const IID_IJMfgPlateDevelopment_AE As String = "{98EA95B3-6A31-11D5-815F-0090276F4297}"
Public Const IID_IJMfgProcessPlate2d_AE As String = "{F4F7B5C2-8588-11D5-8164-0090276F4297}"
Public Const IID_IJMfgPlatePart As String = "{BCA241EE-F5E1-47A8-90DA-17141F9D39BC}"
Public Const IID_IJMfgProfilePart As String = "{1BEB9DD4-3B5D-4571-AEFA-4DC8B9C21434}"
Public Const IID_IJScalingShr As String = "{DE77050C-3300-11D5-BA1A-0090276F4279}"
Public Const IID_IJMfgGeomCol3d As String = "{DB90858B-45B6-11D5-AE21-0004AC96EFFB}"
Public Const IID_IJMfgGeomCol2d As String = "{E6B9C8C8-4AC2-11d5-8151-0090276F4297}"
Public Const IID_IJMfgProfileInit_AE As String = "{D0D32CA2-D1A3-4137-BF78-D60E01C5127E}"
Public Const IID_IJMfgPlateSettingDataInit As String = "{FC4603F0-B2C0-41FB-BB61-15767DCDA619}"
Public Const IID_IJMfgPlateRule_AE As String = "{02C92354-AE04-46CA-B982-A2015A0BF1CE}"
Public Const IID_IJMfgProfileSettingDataInit As String = "{138A949C-887A-406d-80CA-25F6B2706E75}"
Public Const IID_IJMfgProfileRule_AE As String = "{CBDFE280-759F-4157-8C15-EB316C3F36C9}"
Public Const IID_IJMfgProfileDevelopment_AE As String = "{8C689E8D-3402-45d7-B88C-CEDE7BC6AA2F}"

'----------------------------------------------------------------------------
' Other environments
'

Public Const IID_IJPlane As String = "{4317C6B3-D265-11D1-9558-0060973D4824}"
Public Const IID_IJPlate As String = "{53CF4EA0-91BF-11D1-BE56-080036B3A103}"
Public Const IID_IJPlateEntity As String = "{BA8D7A80-9180-11D1-BE56-080036B3A103}"
Public Const IID_IJPlatePart As String = "{780F26C2-82E9-11D2-B339-080036024603}"

Public Const IID_IJSurface As String = "{7D82F810-D270-11d1-9558-0060973D4824}"
Public Const IID_IJPort As String = "{5CF7C404-546D-11D2-B328-080036024603}"
Public Const IID_IJLine As String = "{6C5B8BF2-C81F-11d1-9555-0060973D4824}"
Public Const IID_IJFrame As String = "{80F0A54E-4541-11D1-82D1-0800367F3D03}"
Public Const IID_IJProfile As String = "{0D54249E-7CE9-11D3-B351-0050040EFC17}"
Public Const IID_IJStiffener As String = "{CCBC03F1-223A-11D2-B310-080036024603}"
Public Const IID_IJAppConnectionType As String = "{A6EC992A-902E-11D2-B33D-080036024603}"
Public Const IID_IJAppConnection As String = "{5CF7C402-546D-11D2-B328-080036024603}"
Public Const IID_IJMfgDefinition As String = "{8E96B943-3F93-11D5-BFF5-00902770756B}"
Public Const IID_IJDVector = "{901C8E01-428A-11D1-B832-000000000000}"
Public Const IID_IJProfilePart = "{69F3E7BF-40A0-11D2-B324-080036024603}"
Public Const IID_IJProfileSystem = "{F178F347-F4C7-11D3-9D06-00105AA5BAEB}"
Public Const IID_IJStiffenerPart As String = "{E0B23CD6-7CEB-11D3-B351-0050040EFC17}"
Public Const IID_IJStructGeometry As String = "{6034AD40-FA0B-11d1-B2FD-080036024603}"

Public Const IHFrame As String = "{D21CA530-4556-11d1-82D1-0800367F3D03}" 'kst
Public Const IID_IHFrameSystem As String = "{967E8C33-45F3-11D1-82D3-0800367F3D03}" 'kst
Public Const IHFrameAxis As String = "{80F0A557-4541-11D1-82D1-0800367F3D03}"

Public Const IID_IJShpStrDesignChild = "{982B1A6F-728F-4A6D-8EF8-907282EEF279}"

'- ModelBody --------------------------------------------------------'
Public Const IJDModelBody As String = "{9628D999-CEC5-11D1-B3BE-080036D85603}"

'- copy from the MFgFilterCriteriaKeys.bas --------------------------'

'- MfgParent ------------------------------------------------------'
Public Const IJMfgParent As String = "{063547CD-06F8-11D3-AC33-00104BCC3AC6}"

'- ProductionGeometry ------------------------------------------------------'
Public Const IJDProductionGeometry As String = "{B1E8D472-A966-11D2-A4BE-080036B9C303}"

'- EdgeSet ------------------------------------------------------'
Public Const IJDEdgeSet As String = "{B1E8D473-A966-11D2-A4BE-080036B9C303}"

'- PinJig ------------------------------------------------------'
Public Const IJPinJig As String = "{FE221533-5879-11D5-B86E-0000E2300200}"

'- MfgPin ------------------------------------------------------'
Public Const IJMfgPin As String = "{FE221546-5879-11D5-B86E-0000E2300200}"

'- MfgPlate ----------------------------------------------------'
'Public Const IJMfgPlatePart As String = "{BCA241EE-F5E1-47A8-90DA-17141F9D39BC}"

'- MfgTemplate -------------------------------------------------'
Public Const IJDMfgTemplateSet As String = "{0D5FB0AA-7C0B-4DC3-9F7C-583741D6F542}"

'- MfgProfile --------------------------------------------------'
'Public Const IJMfgProfilePart As String = "{1BEB9DD4-3B5D-4571-AEFA-4DC8B9C21434}"

'- Block ------------------------------------------------------------'
'Public Const IJBlock As String = "{F4EC2FE1-0FF3-11D2-9A02-00A0C94472D6}"

'- Fabrication Margin-----------------------------------------'
'Public Const IJConstMargin As String = "{6A889585-29A6-11D5-BFE5-00902770756B}"

'- Fabrication Margin-----------------------------------------'
Public Const IJObliqueMargin As String = "{A446B3C5-2A8C-11D5-BFE7-00902770756B}"

'- Scaling shrinkage -----------------------------------------'
'Public Const IJScalingShr As String = "{DE77050C-3300-11D5-BA1A-0090276F4279}"

'- Assembly Margin -------------------------------------------'
Public Const IJAssyMarginParent As String = "{A998DAF5-7AB0-4964-8513-188788F40677}"

'- IJMfgGeom3D ----------------------------------------------'
Public Const IJMfgGeom3D As String = "{DB90858C-45B6-11D5-AE21-0004AC96EFFB}"

'- IJMfgMarkingLinesData ----------------------------------------------'
Public Const IJMfgMarkingLinesData As String = "{F3C87DFA-BEE2-4442-94DC-F81BA64EFE8F}"

'-IJMarkingLines_AE -----------------------------------------------------'
Public Const IJMarkingLines_AE As String = "{666DFFE4-015C-4F3F-8087-A0DA1C60A64D}"

 