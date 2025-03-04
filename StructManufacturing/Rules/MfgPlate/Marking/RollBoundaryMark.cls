VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RollBoundaryMark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*******************************************************************
'  Copyright (C) 2002 Intergraph Corporation  All rights reserved.
'
'  Project: MfgPlateMarking
'
'  Abstract:    Rule for creating the Roll Boundary Markings in the MfgPlate command.
'               "CreateBeforeUnfold" generate 3d markings which have to be unfolded,
'               and "CreateAfterUnfold" generate 2d markings which are not unfolded
'
'  History:
'      TBH        feb. 6. 2002    created
'      MJV        2004.04.23      Included correct error handling
'
'******************************************************************

Option Explicit

Implements IJDMfgSystemMarkingRule

Const NAME_FOR_ROLL_BOUNDARY As String = "RE"
' Const NAME_FOR_ROLL_BOUNDARY As String = "ROLL BOUNDARY"

Private Const MODULE = "MfgPlateMarking.RollBoundaryMark"
Private Const IID_IJMfgGeom2d As String = "{E6B9C8CA-4AC2-11D5-8151-0090276F4297}"

Private Sub Class_Initialize()
    PlMrkHelpers.Initialize
End Sub

Private Sub Class_Terminate()
    PlMrkHelpers.UnInitialize
End Sub

Private Function IJDMfgSystemMarkingRule_CreateAfterUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol2d

Const METHOD = "RollboundaryMark: IJDMfgSystemMarkingRule_CreateAfterUnfold"
On Error GoTo ErrorHandler
        
    Dim oPlateWrapper As MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper = New MfgRuleHelpers.PlatePartHlpr
    Set oPlateWrapper.object = Part
    
    Dim oMfgPlateWrapper As MfgRuleHelpers.MfgPlatePartHlpr
    Dim oMfgPart As IJMfgPlatePart
    If oPlateWrapper.PlateHasMfgPart(oMfgPart) Then
        Set oMfgPlateWrapper = New MfgRuleHelpers.MfgPlatePartHlpr
        Set oMfgPlateWrapper.object = oMfgPart
    Else
        Exit Function
    End If
    
    Dim oGeomCol2d As IJMfgGeomCol2d
    Set oGeomCol2d = oMfgPlateWrapper.GetFinal2dGeometries

    If oGeomCol2d Is Nothing Then
        'Since there is nothing to be marked you can exit the function after cleanup
        GoTo CleanUp
    End If
    
    Dim lSwagedPlate As Long
    ' Default is not Swaged
    lSwagedPlate = 0
    
    ' Retrieve default values for custom attributes from the plate
    If TypeOf Part Is IJDAttributes Then
        Dim oAttributeMetaData As IJDAttributeMetaData
        Dim oObject             As IJDObject
        
        Set oObject = Part
        Set oAttributeMetaData = oObject.ResourceManager
        
        Set oObject = Nothing
        ' Get the interface IID of the custom interface IJUASwagePlate
        ' This call might fail if IJUASwagePlate is not bulkloaded
        Dim varIID As Variant
        On Error Resume Next
        varIID = oAttributeMetaData.IID("IJUASwagePlate")
        On Error GoTo ErrorHandler
        
        Dim oAttributes As IJDAttributes
        Set oAttributes = Part
        Dim varCol As Variant
        For Each varCol In oAttributes
            If varCol = varIID Then
                Dim oAttributesCol As IJDAttributesCol
                Set oAttributesCol = oAttributes.CollectionOfAttributes(varIID)
                ' Cycle through all the properties on the custom interface IJUAAssemblyChild
                Dim oAttribute As IJDAttribute
                For Each oAttribute In oAttributesCol
                    Dim sName As String
                    sName = oAttribute.AttributeInfo.Name
                    If sName = "SwagePlate" Then
                        ' Fill in the current value
                        lSwagedPlate = CLng(oAttribute.Value)
                    End If
                Next oAttribute
            End If
        Next varCol
    End If
    
    Dim oGeom2d As IJMfgGeom2d
    Dim oSystemMark As IJMfgSystemMark
    Dim oMarkingInfo As MarkingInfo
    Dim i As Long

    If lSwagedPlate > 0 Then
        ' This is a Swaged plate

        ' We convert 8 roll boundaries to 1 mark
        Dim oGeomObject As IJDObject
        Dim oCollection As IJElements
        Set oCollection = New JObjectCollection
        
        For i = oGeomCol2d.Getcount To 1 Step -1
            Set oGeom2d = oGeomCol2d.GetGeometry(i)
            If oGeom2d.GetGeometryType = STRMFG_ROLL_BOUNDARIES_MARK Then
                oCollection.Add oGeom2d
                If oCollection.Count = 8 Then
                    ' Check if they are all within some range of each other
                    Dim dLowX As Double, dLowY As Double, dLowZ As Double
                    Dim dHighX As Double, dHighY As Double, dHighZ As Double
                    Dim dAddStartX As Double, dAddStartY As Double, dAddStartZ As Double
                    Dim dAddEndX As Double, dAddEndY As Double, dAddEndZ As Double
                    Dim oComplexString As IJComplexString
                    Dim bInit As Boolean
                    bInit = False
                    
                    Dim oTestGeom2d As IJMfgGeom2d
                    For Each oTestGeom2d In oCollection
                        Set oComplexString = oTestGeom2d.GetGeometry
                        
                        Dim oCurve As IJCurve
                        Set oCurve = oComplexString
                        Dim dStartX As Double, dStartY As Double, dStartZ As Double
                        Dim dEndX As Double, dEndY As Double, dEndZ As Double
                        oCurve.EndPoints dStartX, dStartY, dStartZ, dEndX, dEndY, dEndZ
                        Set oComplexString = Nothing
                        
                        If Not bInit Then
                            bInit = True
                            dLowX = dStartX
                            dHighX = dStartX
                            dLowY = dStartY
                            dHighY = dStartY
                            dLowZ = dStartZ
                            dHighZ = dStartZ
                            dAddStartX = dStartX
                            dAddStartY = dStartY
                            dAddStartZ = dStartZ
                            dAddEndX = dEndX
                            dAddEndY = dEndY
                            dAddEndZ = dEndZ
                        Else
                            dAddStartX = dAddStartX + dStartX
                            dAddStartY = dAddStartY + dStartY
                            dAddStartZ = dAddStartZ + dStartZ
                            dAddEndX = dAddEndX + dEndX
                            dAddEndY = dAddEndY + dEndY
                            dAddEndZ = dAddEndZ + dEndZ
                        End If
                        
                        If dStartX > dHighX Then dHighX = dStartX
                        If dStartX < dLowX Then dLowX = dStartX
                        If dStartY > dHighY Then dHighY = dStartY
                        If dStartY < dLowY Then dLowY = dStartY
                        If dStartZ > dHighZ Then dHighZ = dStartZ
                        If dStartZ < dLowZ Then dLowZ = dStartZ
                    Next
                    Dim dLength As Double
                    dLength = Sqr((dHighX - dLowX) * (dHighX - dLowX) + (dHighY - dLowY) * (dHighY - dLowY) + (dHighZ - dLowZ) * (dHighZ - dLowZ))
                    If dLength < 0.3 Then
                        ' We found 8 lines to replace with 1
                        Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
                        oSystemMark.SetMarkingSide UpSide
                        Dim oFoundGeom2d As IJMfgGeom2d
                        Set oFoundGeom2d = oCollection.Item(1)
                        oSystemMark.Set2dGeometry oFoundGeom2d
                    
                        'QI for the MarkingInfo object on the SystemMark
                        Set oMarkingInfo = oSystemMark
                        oMarkingInfo.Name = "BS"
                        
                        Set oComplexString = oFoundGeom2d.GetGeometry
                        Dim oObjCurve As Object
                        oComplexString.GetCurve 1, oObjCurve
                        
                        Dim oLine As IJLine
                        Set oLine = oObjCurve
                        oLine.DefineBy2Points dAddStartX / 8, dAddStartY / 8, dAddStartZ / 8, dAddEndX / 8, dAddEndY / 8, dAddEndZ / 8
                        oComplexString.SetCurve 1, oObjCurve
                        
                        Set oLine = Nothing
                        Set oObjCurve = Nothing
                        Set oComplexString = Nothing
                        Set oFoundGeom2d = Nothing
                    Else
                        Set oGeomObject = oCollection.Item(1)
                        oGeomCol2d.RemoveGeometry oGeomObject
                        ' remove persistent object
                        oGeomObject.Remove
                        Set oGeomObject = Nothing
                    End If
                    
                    oCollection.Remove 1
                    
                End If
                
            End If
            Set oGeom2d = Nothing
            Set oSystemMark = Nothing
            Set oMarkingInfo = Nothing
        Next i
        
        ' Get rid of remaining marking lines
        While oCollection.Count > 0
            Set oGeomObject = oCollection.Item(1)
            oGeomCol2d.RemoveGeometry oGeomObject
            ' remove persistent object
            oGeomObject.Remove
            Set oGeomObject = Nothing
            
            oCollection.Remove 1
        Wend
        Set oCollection = Nothing
        
    Else
        ' This is not a Swaged plate
        For i = 1 To oGeomCol2d.Getcount
            Set oGeom2d = oGeomCol2d.GetGeometry(i)
            If oGeom2d.GetGeometryType = STRMFG_ROLL_BOUNDARIES_MARK Then
                Dim oRelationHelper As IMSRelation.DRelationHelper
                Dim oCollectionHelper As IMSRelation.DCollectionHelper
                                        
                Set oRelationHelper = oGeom2d
                Set oCollectionHelper = oRelationHelper.CollectionRelations(IID_IJMfgGeom2d, "SystemMark2dParent")
                If oCollectionHelper.Count > 0 Then
                    Set oSystemMark = oCollectionHelper.Item(1)
                    oSystemMark.SetMarkingSide UpSide
                Else
                    Set oSystemMark = m_oSystemMarkFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
                    oSystemMark.SetMarkingSide UpSide
                    oSystemMark.Set2dGeometry oGeom2d
                End If
            
                'QI for the MarkingInfo object on the SystemMark
                Set oMarkingInfo = oSystemMark
                oMarkingInfo.Name = NAME_FOR_ROLL_BOUNDARY
                
                Set oCollectionHelper = Nothing
                Set oRelationHelper = Nothing
            End If
            Set oGeom2d = Nothing
            Set oSystemMark = Nothing
            Set oMarkingInfo = Nothing
        Next i
    End If
        
CleanUp:
    Set oPlateWrapper = Nothing
    Set oMfgPlateWrapper = Nothing
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1026, , "RULES")
    GoTo CleanUp
End Function

' ***********************************************************************************
' Function IJDMfgSystemMarkingRule_CreateBeforeUnfold
'
' Description:  User Created Roll Boundary Lines (from the marking line command) will be processed here.
'
' ***********************************************************************************
Private Function IJDMfgSystemMarkingRule_CreateBeforeUnfold(ByVal Part As Object, ByVal UpSide As Long, ByVal bSelectiveRecompute As Boolean, ByVal ReferenceObjColl As GSCADMfgRulesDefinitions.JCmnShp_CollectionAlias) As GSCADMfgRulesDefinitions.IJMfgGeomCol3d
Const METHOD = "RollboundaryMark: IJDMfgSystemMarkingRule_CreateBeforeUnfold"
On Error GoTo ErrorHandler

    Dim oGeomCol3d As IJMfgGeomCol3d
    Set oGeomCol3d = m_oGeom3dColFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
 
    CreateAPSMarkings STRMFG_ROLL_BOUNDARIES_MARK, ReferenceObjColl, oGeomCol3d
    
    'Return the 3d collection
    Set IJDMfgSystemMarkingRule_CreateBeforeUnfold = oGeomCol3d
    
Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 1026, , "RULES")
End Function

