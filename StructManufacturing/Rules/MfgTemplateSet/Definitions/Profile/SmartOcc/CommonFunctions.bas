Attribute VB_Name = "CommonFunctions"
Public Function GetConcaveSide(ByRef pPlatePart As Object) As Integer
Const METHOD = "GetConcaveSide"
On Error GoTo ErrorHandler

    'get COG of the plate
    Dim oCOGPosition As IJDPosition
    Dim oMfgGeomHelper As IJMfgGeomHelper
    Set oMfgGeomHelper = New MfgGeomHelper
    Set oCOGPosition = oMfgGeomHelper.GetCenterOfGravityOfModelBody(pPlatePart)
  
    ''Get the distance of COG to Base surface
        'get the base surface
        'get the shortest distance from base surface
    Dim dMinDistToBase As Double
    Dim dMinDistToOffset As Double
    Dim oSurface As IJSurfaceBody
    Dim oModelBody As IJDModelBody
    Dim oPosOnSurface As IJDPosition
    
    'setup part support
    Dim oSDPartSupport As IJPartSupport
    Set oSDPartSupport = New PlatePartSupport
    Set oSDPartSupport.Part = pPlatePart
    
    Dim oSDPlatePartSupport As IJPlatePartSupport
    Set oSDPlatePartSupport = oSDPartSupport
    
    'get the base surface
    Call oSDPlatePartSupport.GetSurface(BaseSide, oSurface)
        
    'get the shortest distance from base surface
    Set oModelBody = oSurface
    Call oModelBody.GetMinimumDistanceFromPosition(oCOGPosition, oPosOnSurface, dMinDistToBase)
    
    'get the offset surface
    Call oSDPlatePartSupport.GetSurface(OffsetSide, oSurface)
    
    'get the shortest distance from base surface
    Set oModelBody = oSurface
    Call oModelBody.GetMinimumDistanceFromPosition(oCOGPosition, oPosOnSurface, dMinDistToOffset)
    
    If dMinDistToBase < dMinDistToOffset Then
        GetConcaveSide = 5130 'BaseSide
    Else
        GetConcaveSide = 5131 'OffsetSide
    End If

    
CleanUp:

    Exit Function
ErrorHandler:
    Err.Raise Err.Number, Err.Description
    GoTo CleanUp
End Function
' ********************************************************************************************
'         !!!!! End Private Code !!!!!
' ********************************************************************************************

