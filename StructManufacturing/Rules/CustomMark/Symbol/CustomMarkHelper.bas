Attribute VB_Name = "CustomMarkHelper"
'*******************************************************************
'  Copyright (C) 2002 Intergraph.  All rights reserved.
'
'  Project:
'
'  Abstract:    CustomMarkHelper.bas
'
'  History:
'       Siva         April 06, 2010
'******************************************************************
Option Explicit

'---------------------------------------------------
'   Filter Criteria for the inputs
'---------------------------------------------------
Public Const IJPlateSystem = "{E0B23CD4-7CEB-11d3-B351-0050040EFC17}"
Public Const IHFrame = "{D21CA530-4556-11d1-82D1-0800367F3D03}"
Public Const IJSeam = "{02C1327F-2C31-11D2-8329-0800367F3D03}"
Public Const IJAppConnectionType = "{A6EC992A-902E-11D2-B33D-080036024603}"
Public Const IJStiffenerSystem = "{E0B23CD5-7CEB-11d3-B351-0050040EFC17}"
Public Const IJPinJig = "{FE221533-5879-11D5-B86E-0000E2300200}"
Public Const IJPlatePart = "{780F26C2-82E9-11D2-B339-080036024603}"
Public Const RELATION_INSERTED_IN_TRANS = &H200
Private Const REPRESENTATION = "Simple" ' used to named the simple physical representation

'--------------------------------------------------------------------------------------------------
' Abstract : The purpose of this routine is to calculate the intersection between any two given
'            objects. The call will be delegated to the G&T PlaceIntersectionObject routine
'--------------------------------------------------------------------------------------------------
Public Function GetIntersection(pIntersectedObject As Object, pIntersectingObject As Object) As Object
On Error GoTo ErrorHandler
Const METHOD = "GetIntersection"

    ' Find the intersection.
    Dim oGeometryIntersector    As IMSModelGeomOps.DGeomOpsIntersect
    Set oGeometryIntersector = New IMSModelGeomOps.DGeomOpsIntersect
    
    On Error Resume Next 'Needed for continuing with next skid mark if intersection fails
    Dim oIntersectionUnknown    As IUnknown        ' Resultant intersection.
    oGeometryIntersector.PlaceIntersectionObject Nothing, pIntersectedObject, pIntersectingObject, Nothing, oIntersectionUnknown
    
    On Error GoTo ErrorHandler
    Set GetIntersection = oIntersectionUnknown
    Set oGeometryIntersector = Nothing
    Set oIntersectionUnknown = Nothing

    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'********************************************************************
' Routine: GetPlatePartFromPlate
'
' Abstract: Gets the Child PlatePart of the Plate system given
'********************************************************************
Public Function GetAllPartsFromSystem(oInputSys As Object) As IJElements
    Const METHOD = "GetAllPartsFromSystem"
    On Error GoTo ErrorHandler
    
    Dim oPartColl  As IJElements
    Set oPartColl = New JObjectCollection

    Dim colChildren As IJDTargetObjectCol
    Dim oMFSystem As IJSystem
    Dim oPart As Object
    
    'Get the children from the plate system
    Set oMFSystem = oInputSys
    Set colChildren = oMFSystem.GetChildren
    
    'Get the child which supports IJPlatePart/IJStiffenerPart
    Dim i As Integer, j As Integer
    
    For i = 1 To colChildren.Count
        If TypeOf colChildren.Item(i) Is IJPlatePart Or TypeOf colChildren.Item(i) Is IJStiffenerPart Then
            oPartColl.Add colChildren.Item(i)
        Else
            
            On Error Resume Next ' colChildren.Item(i) need not support IJSystem when connected system is split
            
            Dim oLeafSystem As IJSystem
            Set oLeafSystem = colChildren.Item(i)
            
            On Error GoTo ErrorHandler
            
            If Not oLeafSystem Is Nothing Then
            
                Dim oTargetObjColl As IJDTargetObjectCol
                Set oTargetObjColl = oLeafSystem.GetChildren
                
                For j = 1 To oTargetObjColl.Count
                    On Error Resume Next
                    Set oPart = oTargetObjColl.Item(j)
                    
                    If Not oPart Is Nothing Then
                        If TypeOf oPart Is IJPlatePart Or TypeOf oPart Is IJStiffenerPart Then
                            oPartColl.Add oPart
                        End If
                        
                    Else
                        On Error GoTo ErrorHandler
                    End If
                Next
                
            End If
        End If
    Next i
   
    
    Set GetAllPartsFromSystem = oPartColl
    
    'Cleanup
    Set colChildren = Nothing
    Set oMFSystem = Nothing
    Set oPart = Nothing
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'////////////////////////////////////////////////////////////////////
'********************************************************************
'Method: TrackText
'
'Interface: private sub
'
'Abstract:
'********************************************************************
Public Sub TrackText(strText As String, nIndent As Integer)
    On Error GoTo SkipTrack
    
'    TR-CP-227633  Few delivered MFG rules cannot be recompiled on an End user machine
'    Removing the functionality since GSTracker.exe is not delivered in EUS.



'    Dim oTracker As GSTracker.Tracker
'    Set oTracker = New GSTracker.Tracker
'    On Error GoTo KillTrack
'    If nIndent < 0 Then oTracker.Outdent
'    oTracker.UseFontDefaults
'    oTracker.ForegroundColor = RGB(0, 255, 255)
'    oTracker.WriteLn strText
'    If nIndent > 0 Then oTracker.Indent
'
'KillTrack:
'    Set oTracker = Nothing

SkipTrack:

End Sub

Public Sub ReconnectOldOutputsToSymbol(pSymbolOccurrence As Object, pOutputColl As Object)
    Const METHOD = "ReconnectOldOutputs"
    On Error GoTo ErrorHandler
    Dim oRefProxyColl As IJDObjectCollection
    
    ' Please note that ref proxy connected to Marking AE is treated as the actual
    ' output collection connected to symbol through Flavor is deleted when symbol is triggered.

    Set oRefProxyColl = GetReferenceProxyCollection(pSymbolOccurrence)
    
    If oRefProxyColl Is Nothing Then
        Exit Sub
    ElseIf oRefProxyColl.Count = 0 Then
        Exit Sub
    Else
        ' The Marking Line CS is cloned and added as output of the symbol
        
        Dim oPOM As IJDPOM
        Set oPOM = GetPOM
        
        Dim oOutPutCollection   As IJDOutputCollection
        Set oOutPutCollection = pOutputColl
        
        Dim oResourceManager    As Object
        Set oResourceManager = oOutPutCollection.ResourceManager
    
        Dim oObj As Object
        For Each oObj In oRefProxyColl
        
            Dim oRelation As IJDAssocRelation
            Set oRelation = oObj
            
            Dim oTargetObjCol               As IJDTargetObjectCol
            Set oTargetObjCol = oRelation.CollectionRelations("IJGeometry", "RefInput_DEST")
                
            If Not oTargetObjCol Is Nothing Then
                Dim lCount As Long
                For lCount = 1 To oTargetObjCol.Count
                    Dim oOutputUnk As Object
                    Set oOutputUnk = oTargetObjCol.Item(lCount)
                    
                    Dim oMfgmarking_AE As IJMfgMarkingLines_AE
                    Set oMfgmarking_AE = oOutputUnk
                    
                    Dim oPart As Object
                    Set oPart = oMfgmarking_AE.GetMfgMarkingPart
                    
                    'Add the complex string to symbol output collection with part OID as output name
                    'This is needed to eleminate the swapping of ref proxies with Marking line AEs
                    Dim strPartOID As String
                    strPartOID = GetOID(oPOM, oPart)
                    
                    'Remove the curly braces from the OID
                    strPartOID = Replace(strPartOID, "{", "")
                    strPartOID = Replace(strPartOID, "}", "")
                    
                    Dim oCurve As IJCurve
                    Set oCurve = oMfgmarking_AE.GeometryAsComplexString
                    If oCurve.Length > 0.0001 Then
                        Dim oPersistCS As IJComplexString
                        Set oPersistCS = CreatePersistableCS(oCurve, oResourceManager)
                        InitNewOutput pOutputColl, strPartOID
                        
                        pOutputColl.AddOutput strPartOID, oPersistCS
                    End If
                Next
            End If
        Next
        
    End If

    Exit Sub
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

'---------------------------------------------------------------------------------------
' Procedure : GetActiveConnection
' Purpose   : Gets the Active Connection (what else?!)
'---------------------------------------------------------------------------------------
'
Public Function GetPOM() As IJDPOM
    Const METHOD As String = "GetPOM"
    On Error GoTo ErrorHandler

    Dim oCmnAppGenericUtil As IJDCmnAppGenericUtil
    Set oCmnAppGenericUtil = New CmnAppGenericUtil
    
    Dim oActiveConnection As IJDAccessMiddle
    oCmnAppGenericUtil.GetActiveConnection oActiveConnection
    
    Dim jContext As IJContext
    Dim oDBTypeConfiguration As IJDBTypeConfiguration
    
    'Get the middle context
    Set jContext = GetJContext()
    
    'Get IJDBTypeConfiguration from the Context.
    Set oDBTypeConfiguration = jContext.GetService("DBTypeConfiguration")
    
    'Get the Model DataBase ID given the database type
    Dim strConnectionName As String
    strConnectionName = oDBTypeConfiguration.get_DataBaseFromDBType("Model")
    
    Set jContext = Nothing
    Set oDBTypeConfiguration = Nothing

    
    Set GetPOM = oActiveConnection.GetResourceManager(strConnectionName)
    
    Set oActiveConnection = Nothing
    Set oCmnAppGenericUtil = Nothing

    Exit Function

ErrorHandler:
    Err.Raise Err.Number, , Err.Description
End Function

Public Function GetOID(oPOM As IJDPOM, objDB As Object) As String

    'Retrive the OID from the Manufacturing Object
    Dim oMoniker As IMoniker
    Set oMoniker = oPOM.GetObjectMoniker(objDB)
    
    GetOID = oPOM.DbIdentifierFromMoniker(oMoniker)
    
End Function

Public Sub InitNewOutput(pOC As IJDOutputCollection, name As String)
Const METHOD = "InitNewOutput"
On Error GoTo ErrorHandler
    
    Dim oRep As IJDRepresentation
    Dim oOutputs As IJDOutputs
    Dim oOutput As IJDOutput
    Set oOutput = New DOutput
    Set oRep = pOC.definition.IJDRepresentations.GetRepresentationByName(REPRESENTATION)
    Set oOutputs = oRep

    oOutput.name = name
    oOutput.Description = name
    oOutputs.SetOutput oOutput
    oOutput.Reset
    
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, , Err.Description
End Sub

Public Function CreatePersistableCS(oProjCS As IJComplexString, oResourceManager As Object) As IJComplexString
On Error GoTo ErrorHandler
Const METHOD = "CreatePersistableCS"

    Dim oMfgMGhelper As New MfgMGHelper

    ' create persistent complex string
    Dim oGeometryFactory    As IngrGeom3D.GeometryFactory
    Set oGeometryFactory = New IngrGeom3D.GeometryFactory

    Dim oComplexStrings3d As IComplexStrings3d
    Set oComplexStrings3d = oGeometryFactory.ComplexStrings3d

    Dim oCrvElemets As IJElements
    oProjCS.GetCurves oCrvElemets

    Dim oCS3d As ComplexString3d
    Set oCS3d = oComplexStrings3d.CreateByCurves(oResourceManager, oCrvElemets)
    
    Set CreatePersistableCS = oCS3d
    
    Exit Function
ErrorHandler:
    Err.Raise Err.Number, , Err.Description
End Function

'******************************************************************************
' Routine: GetSymbolOutputCollection
'
' Abstract:
' Please note that ref proxy connected to Marking AE is returned
'******************************************************************************
Private Function GetReferenceProxyCollection(oSymbol As IJDSymbol) As IJDObjectCollection
    Const METHOD = "GetReferenceProxyCollection"
    On Error GoTo ErrorHandler
        
    Dim oTempColl As IJDObjectCollection
    Set oTempColl = New JObjectCollection
    
    Dim oRelation As IJDAssocRelation
    Set oRelation = oSymbol
    
    Dim oTargetObjCol               As IJDTargetObjectCol
    Set oTargetObjCol = oRelation.CollectionRelations("IJProxy", "Owner")
        
    If Not oTargetObjCol Is Nothing Then
        Dim lCount As Long
        For lCount = 1 To oTargetObjCol.Count
            Dim oOutputUnk As Object
            Set oOutputUnk = oTargetObjCol.Item(lCount)
            
            Dim oRelation2 As IJDAssocRelation
            Set oRelation2 = oOutputUnk
            
            Dim oTargetObjCol2               As IJDTargetObjectCol
            Set oTargetObjCol2 = oRelation2.CollectionRelations("IJProxy", "Owner")
                
            If Not oTargetObjCol2 Is Nothing Then
                Dim lCount2 As Long
                For lCount2 = 1 To oTargetObjCol2.Count
                    Dim oOutputUnk2 As Object
                    Set oOutputUnk2 = oTargetObjCol2.Item(lCount2)
                    oTempColl.Add oOutputUnk2
                    
                Next
            End If
        Next
    End If
    
    Set GetReferenceProxyCollection = oTempColl
    
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, , Err.Description
End Function
