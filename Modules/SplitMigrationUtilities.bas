Attribute VB_Name = "SplitMigrationUtilities"
'*******************************************************************
'
'Copyright (C) 2006 Intergraph Corporation. All rights reserved.
'
'File : SplitMigrationUtilities.bas
'
'Author : AS
'
'Description :
'    Module for common split migration utilities. SHuold be used to migrate aggregator with direct inputs
'    or ref coll in case of split in any input.
'
'History:
'   20-Jul-2006 AS  Initial Creation
'   01-Aug-2006 AS  TR#99968 Added support for external point supplied by an object wishing to be migrated. Added support for IJPlane in selectReplacingObject
'*****************************************************************************************************************

Option Explicit
Private Const MODULE = "SplitMigrationUtilities"

Public Sub GetPositionFromRefColl(pReferencesCollection As IJDReferencesCollection, ByRef oPoint As IJPoint)
    Const METHOD = "GetPositionFromRefColl"
    On Error GoTo ErrorHandler
    
    Dim pObj As Object
    Dim iPort As IJPort
    Dim oGeom As Object
    Dim iPoint As IJPoint
    Dim eleCount As Long, ii As Long
    Dim x As Double, y As Double, z As Double
    Dim oGeom3DFactory As New GeometryFactory
    
    If oPoint Is Nothing Then
        Set oPoint = oGeom3DFactory.Points3d.CreateByPoint(Nothing, 0#, 0#, 0#)
        If oPoint Is Nothing Then
            Err.Raise E_FAIL
        End If
    End If

    eleCount = pReferencesCollection.IJDEditJDArgument.GetCount
    For ii = 1 To eleCount
        Set pObj = pReferencesCollection.IJDEditJDArgument.GetEntityByIndex(ii)
        If TypeOf pObj Is IJPort Then
            Set iPort = pObj
            Set oGeom = iPort.Geometry
            If TypeOf oGeom Is IJPoint Then
                Set iPoint = oGeom
                iPoint.GetPoint x, y, z
                oPoint.SetPoint x, y, z
                Exit Sub        'found one.
            End If

        ElseIf TypeOf pObj Is IJPoint Then
            Set iPoint = pObj
            iPoint.GetPoint x, y, z
            oPoint.SetPoint x, y, z
            Exit Sub
        End If
    Next ii
    
    'TODO: in the future, we could support get xyz from the intersection of two line ports.
    Err.Raise E_FAIL            'did NOT find one. x,y,z not set.

ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub GetPositionFromElementsList(pElements As IJElements, ByRef oPoint As IJPoint)
    Const METHOD = "GetPositionFromElementsList"
    On Error GoTo ErrorHandler
    
    Dim pObj As Object
    Dim iPort As IJPort
    Dim oGeom As Object
    Dim iPoint As IJPoint
    Dim eleCount As Long, ii As Long
    Dim x As Double, y As Double, z As Double
    Dim oGeom3DFactory As New GeometryFactory
    
    If oPoint Is Nothing Then
        Set oPoint = oGeom3DFactory.Points3d.CreateByPoint(Nothing, 0#, 0#, 0#)
        If oPoint Is Nothing Then
            Err.Raise E_FAIL
        End If
    End If

    eleCount = pElements.count
    For ii = 1 To eleCount
        Set pObj = pElements(ii)
        If TypeOf pObj Is IJPort Then
            Set iPort = pObj
            Set oGeom = iPort.Geometry
            If TypeOf oGeom Is IJPoint Then
                Set iPoint = oGeom
                iPoint.GetPoint x, y, z
                oPoint.SetPoint x, y, z
                Exit Sub        'found one.
            End If

        ElseIf TypeOf pObj Is IJPoint Then
            Set iPoint = pObj
            iPoint.GetPoint x, y, z
            oPoint.SetPoint x, y, z
            Exit Sub
        End If
    Next ii
    
    'TODO: in the future, we could support get xyz from the intersection of two line ports.
    Err.Raise E_FAIL            'did NOT find one. x,y,z not set.

ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub SelectReplacingObject(pObjectCollectionReplacing As IJDObjectCollection, _
        oPoint As IJPoint, ByRef pObj As Object)

    Const METHOD = "SelectReplacingObject"
    On Error GoTo ErrorHandler

    Dim dist As Double, minDist As Double
    Dim oGeom As Object
    Dim pObjectReplacing As Object
    Dim iPort As IJPort
    Dim iPoint As IJPoint
    Dim iCurve As IJCurve
    Dim iSurface As IJSurface
    Dim iPlane As IJPlane
    Dim ptX As Double, ptY As Double, ptZ As Double
    Dim cvX As Double, cvY As Double, cvZ As Double
    Dim sfX As Double, sfY As Double, sfZ As Double
    Dim pnts1() As Double, pnts2() As Double, pars1() As Double, pars2() As Double
    Dim numPts As Long

    minDist = 100000000#

    For Each pObjectReplacing In pObjectCollectionReplacing

        If TypeOf pObjectReplacing Is IJPort Then
            Set iPort = pObjectReplacing
            Set oGeom = iPort.Geometry
        Else
            Set oGeom = pObjectReplacing
        End If
    
        If TypeOf oGeom Is IJPoint Then     'should only be one replacement !
            Set iPoint = oGeom
            dist = iPoint.DistFromPt(oPoint)    'dist should be zero !
        
        ElseIf TypeOf oGeom Is IJCurve Then
            Set iCurve = oGeom
            iCurve.DistanceBetween oPoint, dist, cvX, cvY, cvZ, ptX, ptY, ptZ

        ElseIf TypeOf oGeom Is IJSurface Then
            Set iSurface = oGeom
            iSurface.DistanceBetween oPoint, dist, sfX, sfY, sfZ, ptX, ptY, ptZ, numPts, pnts1, pnts2, pars1, pars2
        
        ElseIf TypeOf oGeom Is IJPlane Then
            Dim oMemberFactory As New SPSMemberFactory
            Dim oFeatureServices As ISPSMemberFeatureServices
            
            Set oFeatureServices = oMemberFactory.CreateMemberFeatureServices
            oFeatureServices.GetMinDistFromPlane oGeom, oPoint, dist
            Set oMemberFactory = Nothing
        Else
            Err.Raise E_FAIL    ' unknown geometry.  perhaps a compound surface from intelliship.
        End If
        
        If (dist >= 0) And (dist < minDist) Then
            minDist = dist
            Set pObj = pObjectReplacing
        End If

    Next

    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub MigrateMemberObject(pMemberObj As Object, pMigrateHelper As IJMigrateHelper)

    Const METHOD = "MigrateMember"
    On Error GoTo ErrorHandler

    Dim pObjectCollectionReplaced As IJDObjectCollection
    Dim pObjectCollectionReplacing As IJDObjectCollection
    
    Set pObjectCollectionReplaced = New JObjectCollection
    Set pObjectCollectionReplacing = New JObjectCollection
    
    Call pObjectCollectionReplaced.Add(pMemberObj)
    Call pObjectCollectionReplacing.Add(pMemberObj)
                    
    Call pMigrateHelper.ObjectsReplaced(pObjectCollectionReplaced, pObjectCollectionReplacing, False)

    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub

Public Sub MigrateTheAggregator(pAggregatorDescription As IJDAggregatorDescription, pMigrateHelper As IJMigrateHelper)

    'kinda cool name huh ?  Like: Luke, the Alien !
    
    Const METHOD = "MigrateTheAggregator"
    On Error GoTo ErrorHandler

    Dim ii As Long, eleCount As Long
    Dim pElements As IJElements
    Dim oPoint As IJPoint
    Dim pObjOld As Object, pObjNew As Object
    Dim pAppConn As IJAppConnection
    Dim pObjectCollectionReplacing As IJDObjectCollection
    Dim bIsDeleted As Boolean
    Dim oDesignChild As IJDesignChild
    Dim oDesignParent As IJDesignParent
    
    'Get the input ports.
    'If an input is on the RefColl, we assume it is not split.
    Set pAppConn = pAggregatorDescription.CAO
    Call pAppConn.enumPorts(pElements)

    'Get the end position of the part using the ports
    GetPositionFromElementsList pElements, oPoint
        
    'for each port
    'if it is replaced,
    '   select the closest one from the replacing list
    '   replace input relation with the selected one.

    Set oDesignChild = pAppConn
    Set oDesignParent = oDesignChild.GetParent

    eleCount = pElements.count
    For ii = 1 To eleCount
        
        Set pObjOld = pElements(ii)
        pMigrateHelper.ObjectsReplacing pObjOld, pObjectCollectionReplacing, bIsDeleted

        If Not pObjectCollectionReplacing Is Nothing Then
            
            SelectReplacingObject pObjectCollectionReplacing, oPoint, pObjNew
    
            If Not pObjNew Is Nothing Then
                Call pAppConn.removePort(pObjOld)
                Call pAppConn.addPort(pObjNew)
                                
                'if the connection is a child of the object being replaced  then make
                'the replacing object the parent
                If Not oDesignParent Is Nothing Then ' check on the current parent
                    Dim oPort As IJPort
                    Dim oPart  As IJConnectable
                    Set oPort = pObjOld
                    Set oPart = oPort.Connectable
                    If Not oPart Is Nothing Then
                        If TypeOf oPart Is IJDesignParent Then
                            If oPart Is oDesignParent Then ' current parent is replaced
                                Set oPort = pObjNew
                                Set oPart = oPort.Connectable
                                If Not oPart Is Nothing Then
                                    If TypeOf oPart Is IJDesignParent Then
                                        Set oDesignParent = oPart
                                        oDesignParent.AddChild oDesignChild
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If

        End If
    Next ii

    Exit Sub

ErrorHandler:
    HandleError MODULE, METHOD
End Sub

'*************************************************************************
'Sub: MigrateRefColl
'
'Abstract: Replaces the objects in the Ref coll with 'matching' objects. Calling function should use the
'           return arguments to replace the refcoll objects themselves in the right relation
'
'Arguments
'pReferencesCollection - The aggregator
'pMigrateHelper - Migrate Helper
'oReplacing - returns Array of replacing objects which match the old objects- by distance test
'bIsInputMigrated - returns Boolean which is true if a ref coll obuject is indeed replaced
'oPoint - optional argument if the ref coll does not have point. This point will be used to select replacingobject.
'****************
Public Sub MigrateRefColl(pReferencesCollection As IJDReferencesCollection, pMigrateHelper As IJMigrateHelper, oReplacing() As Object, bIsInputMigrated As Boolean, Optional oInPoint As IJPoint)
    Const METHOD = "MigrateRefColl"
    On Error GoTo ErrorHandler
    
    Dim oPoint As IJPoint
    Dim ii As Integer, eleCount As Integer
    Dim pObjectCollectionReplacing As IJDObjectCollection
    Dim bIsDeleted As Boolean
    Dim oOld As Object

    If Not oInPoint Is Nothing Then
        Dim tempX As Double
        Dim tempY As Double
        Dim tempZ As Double
        Dim oGeom3DFactory As New GeometryFactory
        oInPoint.GetPoint tempX, tempY, tempZ
        Set oPoint = oGeom3DFactory.Points3d.CreateByPoint(Nothing, 0#, 0#, 0#)
        oPoint.SetPoint tempX, tempY, tempZ
        Set oGeom3DFactory = Nothing
    Else
        GetPositionFromRefColl pReferencesCollection, oPoint
    End If

    eleCount = pReferencesCollection.IJDEditJDArgument.GetCount
    ReDim oReplacing(1 To eleCount)

    For ii = 1 To eleCount

        Set oOld = pReferencesCollection.IJDEditJDArgument.GetEntityByIndex(ii)
        
        Call pMigrateHelper.ObjectsReplacing(oOld, pObjectCollectionReplacing, bIsDeleted)
    
        If Not pObjectCollectionReplacing Is Nothing Then
            bIsInputMigrated = True
            SelectReplacingObject pObjectCollectionReplacing, oPoint, oReplacing(ii)
        Else
            Set oReplacing(ii) = oOld
        End If
        
        Set oOld = Nothing
        Set pObjectCollectionReplacing = Nothing
        
    Next ii
     
    Exit Sub
ErrorHandler:
    HandleError MODULE, METHOD
End Sub

