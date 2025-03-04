VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RefCurveMarks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************************
' Copyright (C) 2009, Intergraph Corp.  All rights reserved.
'
' Project: MfgCSOptions
' Module: RefCurveMarks
'
' Description:  Create Reference Curve Marks on the copied Plate
'
' Author: Manasa.J
'
' Comments:
' User         Date        Comments
' Manasa.J     9.12.09     Initial creation
'*******************************************************************************
Option Explicit
Private Const MODULE = "MfgCSOptions.RefCurveMarks"

Implements IJDMfgCopyStructureMarksRule
Private Function IJDMfgCopyStructureMarksRule_CreateMarks(ByVal oInputPlates As IJElements, ByVal oNewPlates As IJElements) As IJElements

    Const METHOD = "IJDMfgCopyStructureMarksRule_CreateMarks"
    On Error GoTo ErrorHandler
    
    Set IJDMfgCopyStructureMarksRule_CreateMarks = New JObjectCollection
    
    Dim lOldPlatesCount As Long
    Dim lNewPlatesCount As Long
    
    Dim oRefCurveColl As IJDObjectCollection
    Set oRefCurveColl = New JObjectCollection
    
    Dim lMarkingSide As Long
    lMarkingSide = 1111 'molded side
                    
    For lOldPlatesCount = 1 To oInputPlates.Count
       
            Dim oSDPlateWrapper   As StructDetailObjects.PlatePart
            Set oSDPlateWrapper = New StructDetailObjects.PlatePart
            Set oSDPlateWrapper.object = oInputPlates.Item(lOldPlatesCount)
                       
            Dim side As String
            side = oSDPlateWrapper.AlternateMoldedSide
            
            Dim oRefCurveDataColl As Collection
            If side = "Base" Then
                oSDPlateWrapper.ReferenceCurves PlateBaseSide, oRefCurveDataColl
            Else
                oSDPlateWrapper.ReferenceCurves PlateOffsetSide, oRefCurveDataColl
            End If
            
            Dim i As Long
            For i = 1 To oRefCurveDataColl.Count
                Dim oRefCurveData   As IJRefCurveData
                Set oRefCurveData = oRefCurveDataColl.Item(i)
                
                If Not oRefCurveData Is Nothing Then
                    
                    Dim oActualReferenceCurve   As IJRefCurveOnSurface
                    Set oActualReferenceCurve = oRefCurveData.ParentReferenceCurve
                    
                    Dim RefCurveType As JSRCOS_TYPES
                    RefCurveType = oRefCurveData.Type
                    
                    Dim lMarkingType As Long
                    
                    Select Case RefCurveType
                        Case JSRCOS_KNUCKLE
                            lMarkingType = STRMFG_KNUCKLE_MARK
                        Case JSRCOS_REFERENCE
                            lMarkingType = STRMFG_NAVALARCHLINE
                        Case JSRCOS_TANGENT
                            lMarkingType = STRMFG_NAVALARCHLINE
                        Case JSRCOS_UNKNOWN
                            lMarkingType = STRMFG_NAVALARCHLINE
                        Case Else
                            GoTo NextRefCurve
                    End Select
                    
                    If oRefCurveColl.Contains(oActualReferenceCurve) = False Then
                        For lNewPlatesCount = 1 To oNewPlates.Count
                            Dim oMarkingLineObj As Object
                            Set oMarkingLineObj = CreateMarkingLine(oNewPlates.Item(lNewPlatesCount), oActualReferenceCurve, Nothing, oActualReferenceCurve, lMarkingSide, lMarkingType)
                            
                            If Not oMarkingLineObj Is Nothing Then
                                IJDMfgCopyStructureMarksRule_CreateMarks.Add oMarkingLineObj
                            End If
                        Next lNewPlatesCount
                        
                        oRefCurveColl.Add oActualReferenceCurve
                    End If
                    
                    Set oActualReferenceCurve = Nothing
                   
                End If
                
NextRefCurve:
                Set oRefCurveData = Nothing
            Next i
            
            Set oSDPlateWrapper = Nothing
           
    Next lOldPlatesCount
    
    Exit Function

ErrorHandler:
   Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1015, , "RULES")
End Function
