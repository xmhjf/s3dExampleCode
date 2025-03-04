Attribute VB_Name = "Utilities"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   Utilities.bas
'   Author: MU
'   Creation Date:
'   Description:
'       TODO - fill in header description information
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'  19.DEC.2007     PK            Undone the changes made in TR-132021
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Option Explicit
Public Const E_FAIL = -2147467259
Public Const TOLERANCE = 0.000001
Public Const PI = 3.14159265


''' This function creates persistent/Transient Arc based on centre point
''' NormalVector, start and end points of the curve
'''<{(PlaceArcByCen begin)}>
Public Function PlaceArcByCen(pResourceMgr As IUnknown, _
                            ByRef centerPoint As IJDPosition, _
                            ByRef startPoint As IJDPosition, _
                            ByRef endPoint As IJDPosition, _
                            ByRef normVector As IJDVector) _
                            As IngrGeom3D.Arc3d
    
    On Error GoTo ErrorHandler
        
    Dim oArc As IngrGeom3D.Arc3d
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set oArc = oGeomFactory.Arcs3d.CreateByCtrNormStartEnd(pResourceMgr, _
                              centerPoint.x, centerPoint.y, centerPoint.z, _
                             normVector.x, normVector.y, normVector.z, _
                             startPoint.x, startPoint.y, startPoint.z, _
                            endPoint.x, endPoint.y, endPoint.z)
   
    Set PlaceArcByCen = oArc
    Set oArc = Nothing
    Set oGeomFactory = Nothing

    Exit Function
    
ErrorHandler:
   Err.Raise E_FAIL
End Function
'''<{(PlaceArcByCen end)}>


''' This function creates persistent/Transient Line based on
''' start and end points of the line
'''<{(Line begin)}>
Public Function PlaceLine(pResourceMgr As IUnknown, ByRef startPoint As IJDPosition, _
                            ByRef endPoint As IJDPosition) _
                            As IngrGeom3D.Line3d
    
    On Error GoTo ErrorHandler
        
    Dim oLine As IngrGeom3D.Line3d
    Dim oGeomFactory As IngrGeom3D.GeometryFactory
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    
    'MsgBox "about to create the line"
    ' Create Line object
    Set oLine = oGeomFactory.Lines3d.CreateBy2Points(pResourceMgr, _
                                startPoint.x, startPoint.y, startPoint.z, _
                                endPoint.x, endPoint.y, endPoint.z)
    
    
    Set PlaceLine = oLine
    Set oLine = Nothing
    Set oGeomFactory = Nothing

    Exit Function
    
ErrorHandler:
   Err.Raise E_FAIL
End Function
'''<{(Line end)}>


''' This function returns a ComplexString3d created by
''' curves provided
'''<{(PlaceCString begin)}>
Public Function PlaceCString(ByVal startPosition As AutoMath.DPosition, _
                                        ByVal curves As Collection) _
                                        As IngrGeom3D.ComplexString3d

    
    On Error GoTo ErrorHandler
        
    Dim objCString     As IngrGeom3D.ComplexString3d
    Dim oCurPoint       As New AutoMath.DPosition
    Dim curve          As IngrGeom3D.IJCurve
    Dim trCurve        As IngrGeom3D.IJTransform
    Dim objCurve       As Object
    Dim oElems         As IJElements
    Dim oVectorMove    As IJDVector
    Dim oGeomFactory   As IngrGeom3D.GeometryFactory
    Dim x1             As Double
    Dim y1             As Double
    Dim z1             As Double
    Dim x2             As Double
    Dim y2             As Double
    Dim z2             As Double
    
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set oVectorMove = New AutoMath.DVector
    Set oElems = New JObjectCollection
    Set oCurPoint = startPosition
    Dim count As Integer
    count = 1
    For Each objCurve In curves
        Set curve = objCurve
        curve.EndPoints x1, y1, z1, x2, y2, z2
        count = count + 1
        oVectorMove.Set oCurPoint.x - x1, oCurPoint.y - y1, oCurPoint.z - z1
        Dim tForm   As New AutoMath.DT4x4
        tForm.Translate oVectorMove
        Set trCurve = objCurve
        trCurve.Transform tForm
        Set tForm = Nothing
        oElems.Add trCurve
        curve.EndPoints x1, y1, z1, x2, y2, z2
        oCurPoint.Set x2, y2, z2
    Next objCurve
    
    Set objCString = oGeomFactory.ComplexStrings3d.CreateByCurves( _
                                       Nothing, oElems)
                                       
    
    Set PlaceCString = objCString
    
    Set objCString = Nothing
    Set oGeomFactory = Nothing
    oElems.Clear
    Set oElems = Nothing

    Exit Function
    
ErrorHandler:
     Err.Raise E_FAIL
End Function
'''<{(PlaceCString end)}>


''' This function creates persistent revolution based on curve
''' axis of revolution and angle
'''<{(Revolution begin)}>
Public Function PlaceRevolution(ByVal oResourceManager As IUnknown, _
                                ByVal objCurve As Object, _
                                ByVal axisVector As IJDVector, _
                                ByVal oCenterPoint As IJDPosition, _
                                revAngle As Double, _
                                isCapped As Boolean)
   
' Construction of revolution
   On Error GoTo ErrorHandler
        
    Dim oGeomFactory     As IngrGeom3D.GeometryFactory
    Dim objRevolution   As IngrGeom3D.Revolution3d
    Set oGeomFactory = New IngrGeom3D.GeometryFactory
    Set objRevolution = oGeomFactory.Revolutions3d.CreateByCurve( _
                                                        oResourceManager, _
                                                        objCurve, _
                                                        axisVector.x, axisVector.y, axisVector.z, _
                                                        oCenterPoint.x, oCenterPoint.y, oCenterPoint.z, _
                                                        revAngle, isCapped)

    Set PlaceRevolution = objRevolution
    Set objRevolution = Nothing
    Set oGeomFactory = Nothing
    
    Exit Function

ErrorHandler:
   Err.Raise E_FAIL
End Function
'''<{(Revolution end)}>

' CMCache custom method to cache the input argument into a parameter contend and the reverse conversion
'It is up to you to find a way to convert your reference data object to a parameter content
Public Function CMCacheForPart(pInputCM As Object, bArgToCache As Boolean, pToConvert As Object, ByRef pOutput As Object)

 If bArgToCache Then
       
        'Need to convert the graphic input pToConvert into a Parameter ( pOutput)
        Dim oPC As IJDParameterContent
        Set oPC = New DParameterContent
        
        Dim oPart As IJDPart
        Set oPart = pToConvert
        'MsgBox "Partnumber :" & oPart.PartNumber
        
        ' Create a PC whose value is an identifier of the input
        ' Raju,
        ' the property Part_Number must be retrieved form from the pToConvert argument.
        ' I am hardcoding it
        '
        oPC.Type = igString
        oPC.String = oPart.PartNumber

        Set oPart = Nothing

        ' Always return this PC
        Set pOutput = oPC
        Set oPC = Nothing
 Else
        'Need to convert the cached Parameter pToConvert into your reference data object pOutput
        Dim oPCout As IJDParameterContent
        Dim oPCCache As IJDParameterContent
        Set oPCout = New DParameterContent
        
        ' Here there is three options
        ' o Return a parameter contents, that containts the part_number stored by the pToConvert argument
        '   In this case the edit command will have to retrieve the Part object when needed.
        '   Note : It is better to return a copy of the cached object instead the cached object itself.
        '          This allow to avoid the edition of the cached object.
        ' o Retrieve the Part_number and retrieve your part object with it.
        '   With this solution you can have assoc assertion while when this method is called you are in
        '   in a compute process.
        ' o Get the symbol (or equipment) from the pInputCM, query for the IJDReferencesArg interface,
        '   then get the argument at index 1 it is the SiteProxy.
        '   With this last method, the part has to be passed by reference to the symbol,
        '   but that is what you doing with your design.
        
        ' returning NULL means that the cache method is unable to resolve the cache.
        ' as of now anyway the part is passed as an input argument. this issue of caching
        ' has to be resolved in cycle2.
        
        Set pOutput = Nothing
        
        ' Here is the implementation of the option 1
'        Set oPCCache = pToConvert
'        oPCout.Type = oPCCache.Type
'        oPCout.String = oPCCache.String
'        Set pOutput = oPCout
'        Set oPCout = Nothing
        
        ' Here is the implementation of the option 2
'        Dim oRefDB As GSCADRefDataServices.IJDRefDBGlobal
'        Dim oPart As IJDPart

'        Set oRefDB = New GSCADRefDataServices.RefDBGlobal
'        Set oPart = oRefDB.GetCatalogPart("Storage Tanks", oPCCache.String)
'        Set pOutput = oRefDB
'        Set oRefDB = Nothing
'        Set oPart = Nothing
End If

End Function




