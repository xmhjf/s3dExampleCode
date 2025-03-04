Attribute VB_Name = "StrMfgDataErrors"

Public Enum enumTemplateDataErrors
    TPL_FailedToValidateProcessSettings = 3014 'Error in Validation of ProcessSettings
    TPL_FailedToGetSurfaceFromPlatePart = 3015 'Template Service routine InitSettings:Failed to get Surface from Plate Part
    TPL_CBP_FailedToGetSurfaceFromPlatePart = 3101 'Template Service routine Create Base Plane: Failed to get Surface from Plate Part
    TPL_CBP_FailedToGetEgesFromPlatePart = 3102 'Template Service routine Create Base Plane: Failed to get edges from Plate Part
    TPL_CBP_FailedToIntersectSurfaceWithPlane = 3103 'Template Service routine Create Base Plane: Failed to intersect Surface with Plane
    TPL_CBP_FaliedToCreateBasePlane = 3104 'Template Service routine Create Base Plane: Failed to Create Base Plane
    TPL_CBCL_FailedToGetSurfaceFromPlatePart = 3111 'Template Service routine Create BaseControl Line: Failed to get Surface from Plate Part
    TPL_CBCL_FailedToGetEgesFromPlatePart = 3112 'Template Service routine Create BaseControl Line: Failed to get edges from Plate Part
    TPL_CBCL_FailedToIntersectSurfaceWithPlane = 3113 'Template Service routine Create BaseControl Line: Failed to intersect Surface with Plane
    TPL_CBCL_FaliedToCreateBaseControlLine = 3114 'Template Service routine Create BaseControl Line: Failed to create Base Control Line
    TPL_TC_FaliedToProjectComplexStringToSurf = 3122 'Template Service routine Template Contours: Failed to project Complex String to Surface
    TPL_TC_FaliedToGetIntersectionBetweenCurves = 3123 'Template Service routine Template Contours: Failed to get Intersection between curves
    TPL_TC_FailedToIntersectSurfaceWithPlane = 3124 'Template Service routine Template Contours: Failed to intersect Surface with Plane
    TPL_TC_FailedToGetEdgesFromPlatePart = 3125 'Template Service routine Template Contours: Failed to get edges from Plate Part
    TPL_TC_FailedToIntersectCurveWithPlane = 3126 'Template Service routine Template Contours: Failed to intersect curve with Plane
    TPL_TC_FailedToDefTemplatePlaneFromCornerPts = 3127 'Template Service routine Template Contours: Failed to define Template Plane from corner points
    TPL_TC_FailedToCreateBottomLines = 3128 'Template Service routine Template Contours: Failed to Create Bottom Lines
    TPL_TC_FailedToApplyExtToLocationMarkLines = 3129 'Template Service routine Template Contours: Failed to Apply Extension to Template Location Mark Lines
    TPL_TC_FailedToApplyMargin = 3130 'Template Service routine Template Contours: Failed to Apply Margin
    TPL_TC_FailedToApplyOffset = 3131 'Template Service routine Template Contours: Failed to Apply Offset
    TPL_TC_FailedToCreateTemplateContours = 3132 'Template Service routine Template Contours: Failed to Create Template Contours
    TPL_TC_FailedToGetFrameSystem = 3133 'Template Service routine Template Contours: Failed to get Frame System
    TPL_VPS_InvalidDirectionForFrameType = 3141 'If Type is Frame, Direction should be Transversal
    TPL_VPS_InvalidPositionEvenForAlongFrameOrientation = 3142 'If Orientation is AlongFrame, PositionEven should be NotUsed
    TPL_VPS_InvalidPositionFramesForAlongFrameOrientation = 3143 'If Orientation is AlongFrame, PositionFrames should be Along Frame
    TPL_VPS_InvalidOrientationForPerpendicularType = 3144 'If Type is Perpendicular, Orientation should be Perpendicular
    TPL_VPS_InvalidPositionEvenAndPositionFrame = 3145 'If PositionFrame is chosen, PositionEven should be NotUsed
    
End Enum


 
