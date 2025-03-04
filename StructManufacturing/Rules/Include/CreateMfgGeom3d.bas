Attribute VB_Name = "CreateMfgGeom3d"
'*******************************************************************
'  Copyright (C) 2006 Intergraph.  All rights reserved.
'
'  Project:
'
'  Abstract:    CreateMfgGeom3d.bas
'
'  History:
'     Anand     May 31, 2006  Creation
'******************************************************************

Private Const MODULE As String = "CreateMfgGeom3d"

'---------------------------------------------------------------------------------------
' Procedure : GetActiveConnectionName
' Purpose   : Gets the Name of the Active Connection (what else?!)
'---------------------------------------------------------------------------------------
'
Private Function GetActiveConnectionName() As String
    Const METHOD As String = "GetActiveConnectionName"
    On Error GoTo ErrorHandler

    Dim jContext As IJContext
    Dim oDBTypeConfiguration As IJDBTypeConfiguration
    
    'Get the middle context
    Set jContext = GetJContext()
    
    'Get IJDBTypeConfiguration from the Context.
    Set oDBTypeConfiguration = jContext.GetService("DBTypeConfiguration")
    
    'Get the Model DataBase ID given the database type
    GetActiveConnectionName = oDBTypeConfiguration.get_DataBaseFromDBType("Model")
    
    Set jContext = Nothing
    Set oDBTypeConfiguration = Nothing

    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 5003, , "RULES")
    
End Function

'---------------------------------------------------------------------------------------
' Procedure : GetActiveConnection
' Purpose   : Gets the Active Connection (what else?!)
'---------------------------------------------------------------------------------------
'
Private Function GetActiveConnection() As IJDAccessMiddle
    Const METHOD As String = "GetActiveConnection"
    On Error GoTo ErrorHandler

    Dim oCmnAppGenericUtil As IJDCmnAppGenericUtil
    Set oCmnAppGenericUtil = New CmnAppGenericUtil
    
    Dim oActiveConnection As IJDAccessMiddle
    oCmnAppGenericUtil.GetActiveConnection oActiveConnection
    
    Set GetActiveConnection = oActiveConnection
    
    Set oActiveConnection = Nothing
    Set oCmnAppGenericUtil = Nothing

    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 5004, , "RULES")
End Function

Public Function CreateMfgGeomCol3dObject() As IJMfgGeomCol3d
    Const METHOD As String = "CreateMfgGeomCol3dObject"
    On Error GoTo ErrorHandler
    
    Dim oGeomCol3dFactory    As GSCADMfgGeometry.MfgGeomCol3dFactory
    Set oGeomCol3dFactory = New GSCADMfgGeometry.MfgGeomCol3dFactory
    
    Set CreateMfgGeomCol3dObject = oGeomCol3dFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
    
    Set oGeomCol3dFactory = Nothing
    
    Exit Function
    
ErrorHandler:
    Err.Raise Err.Number, Err.Source, Err.Description
End Function

'---------------------------------------------------------------------------------------
' Procedure : CreateMfgGeom3d
' Purpose   :
' Inputs    :
    ' Geometry as IJComplexString
    ' Geometry Type as Long (ideally StrMfgGeometryType)
    ' Original Object as Object (will be stored as Moniker)
    ' Geometry Collection Holder as IJMfgGeomCol3d
    ' Name of this object as String (if empty, will be name of parent object)
    ' Geometry sub-type (not normally required)
'---------------------------------------------------------------------------------------

Public Function CreateMfgGeom3dObject(oCS As IJComplexString, _
                                      lGeomType As Long, _
                                      oParentObject As Object, _
                                      Optional oGeomCol3d As IJMfgGeomCol3d, _
                                      Optional bstrName As String, _
                                      Optional lGeomSubType As Long) As IJMfgGeom3d
    
    Const METHOD As String = "CreateMfgGeom3d"
    On Error GoTo ErrorHandler

    Dim oGeom3dFactory As GSCADMfgGeometry.MfgGeom3dFactory
    Set oGeom3dFactory = New GSCADMfgGeometry.MfgGeom3dFactory
    
    ' Create the object from factory
    Dim oGeom3d As IJMfgGeom3d
    Set oGeom3d = oGeom3dFactory.Create(GetActiveConnection.GetResourceManager(GetActiveConnectionName))
    
    ' Stuff it with attributes
    oGeom3d.PutGeometry oCS
    oGeom3d.PutGeometrytype lGeomType
    oGeom3d.PutSubGeometryType lGeomSubType
    
    If Not oGeomCol3d Is Nothing Then
        ' Make a relation between newly created object and input parent object
        oGeomCol3d.AddGeometry oGeomCol3d.Getcount + 1, oGeom3d
    End If
    
    ' Set its name
    If TypeOf oGeom3d Is IJNamedItem Then
        Dim oName As IJNamedItem
        Set oName = oGeom3d
        oName.Name = bstrName
        Set oName = Nothing
    End If
    
    Dim oObjUtil As IJDMfgGeomUtilWrapper
    Set oObjUtil = New MfgGeomUtilWrapper
    
    If Not oParentObject Is Nothing Then
        If TypeOf oParentObject Is IJDObject Then
            oGeom3d.PutMoniker oObjUtil.GetMoniker(oParentObject)
        End If
    End If
    
    Set CreateMfgGeom3dObject = oGeom3d
    
    Set oGeom3dFactory = Nothing
    Set oGeom3d = Nothing
    Set oObjUtil = Nothing
    
    Exit Function

ErrorHandler:
    Err.Raise StrMfgLogError(Err, MODULE, METHOD, , "SMCustomWarningMessages", 5005, , "RULES")
End Function

