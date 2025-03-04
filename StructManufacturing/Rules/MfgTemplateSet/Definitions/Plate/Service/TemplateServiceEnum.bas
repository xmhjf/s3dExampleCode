Attribute VB_Name = "TemplateServiceEnum"
'************************************************************************************************************
'Copyright (C) 2011, Intergraph Limited. All rights reserved.
'
'Abstract:
'   Bas file with enumerators used in this project
'
'Description:
'History :
'   Raman Dubbareddy         11/28/2011      Added new bas file
'************************************************************************************************************


Public Enum enumPlateType
    SymToCenterLine = 0
    PerpendicularXY ' Next to Stem/Stern
    Box
    MostFlatPlate
    NormalPlate
End Enum

Public Enum enumDirection
    eXDir = 0
    eYDir
    eZDir
End Enum

Public Enum enumGroupType
    Primary = 1
    Secondary
    NotInPrimaryOrSecondary
End Enum

'Public Enum enumTemplateType
'    FrameType = 0
'    CenterLineType
'    PerpendicularType
'    Stem_SternType
'    PerpendicularXYType
'End Enum
'
 
