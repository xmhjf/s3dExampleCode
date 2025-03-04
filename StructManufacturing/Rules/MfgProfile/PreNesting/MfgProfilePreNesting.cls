VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MfgProfile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2009, Intergraph Corp.  All rights reserved.
'
' Project: StrMfgPrenestingRules
' Module: MfgProfile
'
' Description:  To customize the bevel correction for profile free ends.
'  History:
'       StructMfg         February 5th, 2010   created
'
'*******************************************************************************
Option Explicit

Implements IJMfgProfilePreNest
Private Sub IJMfgProfilePreNest_GetProfilePreNestEndCorrection( _
    ByVal oProfilePart As Object, _
    ByVal oMfgGeom2D As Object, _
    ByVal oEndPointOrLC As Object, _
    ByVal oBoundingPartObj As Object, _
    ByVal oSplittingRefPlane As Object, _
    dWebEndCorrection As Double, _
    dFlangeEndCorrection As Double)

    ' oProfilePart - Input - Profile Part to be unfolded
    ' oMfgGeom2D - Input - MfgGeom2d Object on which this bevel correction is being applied.
    ' oEndPointOrLC - Input - Profile End Which at which correction is being calculated. This argument could be a logical connection or End Point.
    ' oBoundingPartObj - Input - Bounding Object
    ' oSplittingRefPlane - Input - If the profile is split by a reference plane, ref plane is sent
    ' dWebEndCorrection - Input/Output - This argument is input as well as output. dWebEndCorrection is the bevel correction value
    '                  that will be applied on this profile end for Web Contours. User can change it in this program.
    ' dFlangeEndCorrection - Input/Output - This argument is input as well as output. dFlangeEndCorrection is the bevel correction value
    '                  that will be applied on this profile end for Flange Contours. User can change it in this program.


'''    Profile End Condition           oEndPointOrLC       oBoundingPartObj        oSplittingRefPlane
'''    =====================           =============       ================        ==================
'''    Free End                        Position            Nothing                 Nothing
'''    Bounded By Reference Plane      Position            Reference Plane         Nothing
'''    Split By a Reference Plane      LogicalConnection   Adjacent Profile Part   Reference Plane
'''    Bounded by Plate                LogicalConnection   Bounding Plate Part     Nothing
'''    Bounded by Profile              LogicalConnection   Bounding Profile Part   Nothing
'''
'''

    ' Following code will be executed, only If a profile is split by a reference plane
    If (Not oBoundingPartObj Is Nothing) Then
        If ((TypeOf oBoundingPartObj Is IJProfilePart) And (Not oSplittingRefPlane Is Nothing)) Then
            Dim oSDThisProfilePart As New StructDetailObjects.ProfilePart
            Set oSDThisProfilePart.Object = oProfilePart
            
            ' Seam Point Split (Longitudinal Profile): Length is based on the Butt Global of X axis.
            ' Seam Point Split (Transversal, Vertical Profile): Butt Normal of the Landing Curve.
            If ((oSDThisProfilePart.ProfileType = sptTransversal) Or (oSDThisProfilePart.ProfileType = sptVertical)) Then
                dWebEndCorrection = 0
                dFlangeEndCorrection = dWebEndCorrection
            End If
                    
            ' If the profile has any Flange, apply the difference of web lengths as the bevel correcrtion.
            If (oSDThisProfilePart.FlangeThickness > 0.001) Then
                Dim dBoundingProfileHeight As Double, dThisProfileHeight As Double
                Dim oSDBoundingProfilePart As New StructDetailObjects.ProfilePart
                
                dThisProfileHeight = oSDThisProfilePart.WebLength
                
                Set oSDBoundingProfilePart.Object = oBoundingPartObj
                dBoundingProfileHeight = oSDBoundingProfilePart.WebLength
                
                ' Apply the bevel correction, only if this profile's web length is bigger than the bounding profile's
                If Abs(dThisProfileHeight - dBoundingProfileHeight) > 0.001 Then
                    If dThisProfileHeight > dBoundingProfileHeight Then
                        dWebEndCorrection = dWebEndCorrection + (Abs(dThisProfileHeight - dBoundingProfileHeight))
                        dFlangeEndCorrection = dWebEndCorrection
                    End If
                End If
                Set oSDBoundingProfilePart = Nothing
            End If
            Set oSDThisProfilePart = Nothing
        End If
    End If
End Sub
