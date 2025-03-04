VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PositionFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2002, Intergraph Corp.  All rights reserved.
'
' Project: MfgTemplateProcess
' Module:
'
' Description:  Determines the process settings for the Template
'
' Author:
'
' Comments:
' 2004.04.22    MJV     Included correct error handling
' 2006.09.11    Ram     Modified to use MfgFrameSystem
'*******************************************************************************
Option Explicit
Private Const MODULE = "MfgTemplateProcess.PositionFrame"
Private Const PART_SURFACE = 0
Private Const MOLDED_SURFACE = 1
Private Const PART_SURFACE_BASE = 2

Implements IJMfgTemplatePositionFrameRule
 
Private Function IJMfgTemplatePositionFrameRule_GetPositionFrame(ByVal pFrameSysDsp As Object, ByVal pPlatePart As IJPlatePart, ByVal pProcessSettings As Object, ByVal pMfgTemplateSetDisp As Object) As IJElements
Const METHOD = "IJMfgTemplatePositionFrameRule_GetPositionFrame"
On Error GoTo ErrorHandler

    Dim sERROR As String
    sERROR = ""

    Dim oGeomHelper As MfgGeomHelper
    Set oGeomHelper = New MfgGeomHelper

    Dim oProcessSettings As IJMfgTemplateProcessSettings
    Set oProcessSettings = pProcessSettings

    Dim oPlateSideRule   As IJDMfgPlateUpSideRule
    Dim oSettingsHelper  As MfgSettingsHelper
    Dim Upside As enumPlateSide
    Dim strProgId As String

    Set oSettingsHelper = oProcessSettings

    strProgId = oSettingsHelper.GetProgIDFromAttr("Side")
    Set oPlateSideRule = SP3DCreateObject(strProgId)
    Upside = oPlateSideRule.GetPlateUpSide(pPlatePart)

    Dim eThicknessSide As PlateThicknessSide
    If (Upside = BaseSide) Then
        eThicknessSide = PlateBaseSide
    ElseIf (Upside = OffsetSide) Then
        eThicknessSide = PlateOffsetSide
    End If
    
    Dim eSurfaceType As eStrMfgSurfaceType
    Dim oMfgTemplateSet As IJDMfgTemplateSet
    Dim oSurfaceBody    As IJSurfaceBody
    Dim lSurfaceType As Long
    Set oMfgTemplateSet = pMfgTemplateSetDisp
    lSurfaceType = oMfgTemplateSet.SurfaceType
    If lSurfaceType = PART_SURFACE Then
        eSurfaceType = TRUE_PART
    ElseIf lSurfaceType = MOLDED_SURFACE Then
        eSurfaceType = TRUE_MOLD
    ElseIf lSurfaceType = PART_SURFACE_BASE Then
        On Error Resume Next
        
        Dim oStructGeom As IJStructGeometry
        Set oStructGeom = pPlatePart
        Set oSurfaceBody = oStructGeom
        
        On Error GoTo ErrorHandler
        
        If oSurfaceBody Is Nothing Then
            eSurfaceType = TRUE_MOLD
            eThicknessSide = PlateSideUnspecified
        End If
    End If

    If oSurfaceBody Is Nothing Then
        ' True or False  doesn't care...it will be used only for determining the direction of Reference Planes
        Set oSurfaceBody = oGeomHelper.GetSurfaceFromPlate(pPlatePart, eSurfaceType, eThicknessSide, 0)
    End If

    Dim bIsTemplateOnFrame As Boolean
    bIsTemplateOnFrame = True
    Dim strTempPosition As String
    strTempPosition = oProcessSettings.TemplatePositionEven
    If strTempPosition = "PositionEven" Then
        bIsTemplateOnFrame = False
    End If

    Dim strTemplateType As String
    strTemplateType = oProcessSettings.TemplateType

    Dim oDirectionVec As DVector
    Set oDirectionVec = New DVector

    Dim oFramesColl As IJElements
    Set oFramesColl = New JObjectCollection
    Dim oTempFrameColl As IJElements
    Set oTempFrameColl = New JObjectCollection
    Dim strDirection As String
    strDirection = oProcessSettings.TemplateDirection
    If strTemplateType = "Frame" Then 'Frame
        oDirectionVec.Set 1, 0, 0
        oGeomHelper.GetReferencePlanesInRange oSurfaceBody, oDirectionVec, oTempFrameColl
        oFramesColl.AddElements oTempFrameColl
    ElseIf strTemplateType = "Box" Or strTemplateType = "UserDefined Box" Or _
              strTemplateType = "UserDefined Box With Edges" Then
        ''Get primary and secondary direction vectors
       If strDirection = "Longitudinal" Then 'X - Direction(Buttock)
           oDirectionVec.Set 0, 1, 0
       ElseIf strDirection = "Transversal" Then 'Y - Direction(Frame)
           oDirectionVec.Set 1, 0, 0
       Else 'Z - Direction(WaterLine)
           oDirectionVec.Set 0, 0, 1
       End If
       
       oGeomHelper.GetReferencePlanesInRange oSurfaceBody, oDirectionVec, oTempFrameColl
       oFramesColl.AddElements oTempFrameColl
        
        'Get Frames in secondary direction
        Dim oParallelAxisVec As IJDVector
        Dim oSecondaryDirVec As IJDVector
        Set oParallelAxisVec = GetParallelAxis(oSurfaceBody)
        Set oSecondaryDirVec = oParallelAxisVec.Cross(oDirectionVec)
        
       oGeomHelper.GetReferencePlanesInRange oSurfaceBody, oSecondaryDirVec, oTempFrameColl
       oFramesColl.AddElements oTempFrameColl
          
   Else
        ''Check the Direction Longitudinal, Transversal, Waterline
       If strDirection = "Longitudinal" Then 'X - Direction(Buttock)
           oDirectionVec.Set 0, 1, 0
       ElseIf strDirection = "Transversal" Then 'Y - Direction(Frame)
           oDirectionVec.Set 1, 0, 0
       Else 'Z - Direction(WaterLine)
           oDirectionVec.Set 0, 0, 1
       End If
        oGeomHelper.GetReferencePlanesInRange oSurfaceBody, oDirectionVec, oTempFrameColl
        oFramesColl.AddElements oTempFrameColl
    End If


    ''''''''''''''''Sort Frames

    Dim oFrameSet As IJElements
    Set oFrameSet = New JObjectCollection

    Dim nIndex As Long
    Dim oFrame As IHFrame
    Dim oFrameAxis As IHFrameAxis
    Dim oFrameSystem As IHFrameSystem

    If Not pFrameSysDsp Is Nothing Then
        For nIndex = 1 To oFramesColl.Count
            Set oFrame = oFramesColl.Item(nIndex)
            Set oFrameAxis = oFrame.FrameAxis
            Set oFrameSystem = oFrameAxis.FrameSystem

            If oFrameSystem Is pFrameSysDsp Then
                oFrameSet.Add oFramesColl.Item(nIndex)
            End If
        Next nIndex
    End If

      ' Send the data back
    Set IJMfgTemplatePositionFrameRule_GetPositionFrame = oFrameSet

  ' *************************************************************
  ' * End Example
  ' *************************************************************


    Set oGeomHelper = Nothing
    Set oProcessSettings = Nothing
    Set oSettingsHelper = Nothing
    Set oPlateSideRule = Nothing
    Set oSurfaceBody = Nothing
    Set oDirectionVec = Nothing
    Set oFramesColl = Nothing
    Set oFrameSet = Nothing

 Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 3010, , "RULES")
End Function
Private Function GetParallelAxis(ByVal oSurfaceBody As IJSurfaceBody) As IJDVector
Const METHOD = "GetParallelAxis"
    On Error GoTo ErrorHandler
    
     Dim oCenter             As IJDPosition
     Dim oNormal             As IJDVector
     Dim oFinalNormal        As IJDVector
     Dim oTopoLocate         As IJTopologyLocate
     Dim oTemplateHelper     As MfgTemplateHelper
     
     Set oTopoLocate = New TopologyLocate
     oTopoLocate.FindApproxCenterAndNormal oSurfaceBody, oCenter, oNormal
     Set oTopoLocate = Nothing
    
     Dim oXVector As IJDVector
     Dim oYVector As IJDVector
     Dim oZVector As IJDVector
     
     Set oXVector = New DVector
     Set oYVector = New DVector
     Set oZVector = New DVector
     
     oXVector.Set 1, 0, 0
     oYVector.Set 0, 1, 0
     oZVector.Set 0, 0, 1
     
     'normalize the vectors
     oXVector.Length = 1#
     oYVector.Length = 1#
     oZVector.Length = 1#
     
     Dim dX As Double
     Dim dY As Double
     Dim dZ As Double
     
     dX = Abs(oXVector.Dot(oNormal))
     dY = Abs(oYVector.Dot(oNormal))
     dZ = Abs(oZVector.Dot(oNormal))
     
     If (dX > dY) And (dX > dZ) Then
         Set GetParallelAxis = oXVector
     ElseIf (dY > dX) And (dY > dZ) Then
         Set GetParallelAxis = oYVector
     ElseIf (dZ > dX) And (dZ > dY) Then
         Set GetParallelAxis = oZVector
     End If

Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function


