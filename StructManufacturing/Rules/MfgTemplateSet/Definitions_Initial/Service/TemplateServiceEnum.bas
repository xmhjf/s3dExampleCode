Attribute VB_Name = "TemplateServiceEnum"
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
 
