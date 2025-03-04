VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FittingMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2011 Intergraph Corporation  All rights reserved.
'
'  Project  : StrMfgTemplateMarkingTube
'
'  Abstract : Creates Fitting marks on Template
'
'  History  :
'      Siva   20th October 2011    created
'
'******************************************************************
Option Explicit
Private Const MODULE = "StrMfgTemplateMarkingTube.FittingMark"

Implements IJMfgTemplateMarkingRule

Private Function IJMfgTemplateMarkingRule_CreateMarks(ByVal oMfgTemplate As IJMfgTemplate, ByVal oReferenceObjColl As IJElements) As IJElements
Const METHOD = "IJMfgTemplateMarkingRule_CreateMarks"
On Error GoTo ErrorHandler

    Dim oTopCurve2D  As IJCurve
    
    ' Get the 2D top line from the input reference collection
    Dim iIndex  As Long
    For iIndex = 1 To oReferenceObjColl.Count
        
        Dim oTempObj    As Object
        Set oTempObj = oReferenceObjColl.Item(iIndex)
        
        If TypeOf oTempObj Is IJMfgGeom2d Then
            Dim oGeom2D As IJMfgGeom2d
            Set oGeom2D = oTempObj
            
            If oGeom2D.GetGeometryType = STRMFG_TopLine Then
                Set oTopCurve2D = oGeom2D.GetGeometry
                Exit For
            End If
        End If
    Next
    
    Dim oOffsetVec  As IJDVector
    Set oOffsetVec = New DVector
    
    oOffsetVec.Set 0, -1, 0
    oOffsetVec.length = TEMPLATE_FITTING_MARK_OFFSET

    ' Get the start and end points of the 2D top curve
    Dim dStartX As Double, dStartY As Double, dStartZ As Double, dEndX As Double, dEndY As Double, dEndZ As Double
    oTopCurve2D.EndPoints dStartX, dStartY, dStartZ, dEndX, dEndY, dEndZ
    
    Dim oStartPos As IJDPosition
    Set oStartPos = New DPosition
    
    Dim oEndPos As IJDPosition
    Set oEndPos = New DPosition
    
    oStartPos.Set dStartX, dStartY, dStartZ
    oEndPos.Set dEndX, dEndY, dEndZ
    
    Dim oNewStartPos As IJDPosition
    Dim oNewEndPos As IJDPosition
    
    ' Offset the top curve by the offset to create the fitting mark
    Set oNewStartPos = oStartPos.Offset(oOffsetVec)
    Set oNewEndPos = oEndPos.Offset(oOffsetVec)
    
    Dim oFittingMark    As IJCurve
    Set oFittingMark = CreateLineFromPoints(oNewStartPos, oNewEndPos)
    
    Dim oTempElems      As IJElements
    Set oTempElems = New JObjectCollection
    
    ' Create Geom2D for the fitting mark
    Dim oMfgGeom2D As IJMfgGeom2d
    Set oMfgGeom2D = CreateGeom2D(oFittingMark, STRMFG_FITTING_MARK, Nothing, "Fitting Mark")
    
    oTempElems.Add oMfgGeom2D
    
    Set IJMfgTemplateMarkingRule_CreateMarks = oTempElems
    Exit Function

ErrorHandler:
   Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 2004, , "RULES")
End Function

