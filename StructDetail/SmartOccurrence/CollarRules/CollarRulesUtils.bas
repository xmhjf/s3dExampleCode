Attribute VB_Name = "CollarRulesUtils"
Option Explicit
Private Const MODULE = "CollarRules::CollarRulesUtils"
'

'  This method is private to this project and creates a corner snipe or a drain hole.(Essentially both are corner features).
'  Inputs:
'    oCollar                     ---- Collar object on which to place snipe or drain hole
'    eCtx                        ---- Base or Offset face (CTX_BASE,CTX_OFFSET)
'    eBottomXid              ---- Support1 port xid  ( JXSEC_BOTTOM, JXSEC_BOTTOM_LEFT,JXSEC_BOTTOM_RIGHT)
'    eWebXid                  ---- Support2 port xid (JXSEC_LEFT, JXSEC_RIGHT)
'    strSnipeOrDrainHole ---- Type of object to create ("SnipeOrCollar", "DrainHole")
'    oObjectCreated        ---- Snipe or drain created

Private Sub CreateSnipeOrDrainHole(ByVal oCollar As Object, _
                                                      ByVal oResourceManager As IUnknown, _
                                                      ByVal eCtx As eUSER_CTX_FLAGS, _
                                                      ByVal eBottomXid As JXSEC_CODE, _
                                                      ByVal eWebXid As JXSEC_CODE, _
                                                      ByVal strSnipeOrDrainHole As String, _
                                                      ByRef oObjectCreated As Object)
    On Error GoTo ErrorHandler
        
    Dim oFacePort As IJPort
    Dim oWebPort As IJPort
    Dim oBottomPort As IJPort

    Dim oCollarWrapper As New StructDetailObjects.Collar

    Set oCollarWrapper.object = oCollar
    oCollarWrapper.GetPortsForCornerFeature eCtx, eWebXid, eBottomXid, oFacePort, oWebPort, oBottomPort
    Set oCollarWrapper = Nothing
        
    Dim oCornerFeatureWrapper As New StructDetailObjects.CornerFeature
    
    ' Bottom is Support1, Web is Support2
    oCornerFeatureWrapper.Create oResourceManager, _
                                                  oFacePort, _
                                                  oBottomPort, _
                                                  oWebPort, _
                                                  strSnipeOrDrainHole, _
                                                  oCollar
    Set oFacePort = Nothing
    Set oWebPort = Nothing
    Set oBottomPort = Nothing
    
    Set oObjectCreated = oCornerFeatureWrapper.object
    Set oCornerFeatureWrapper = Nothing
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CreateSnipeOrDrainHole").Number
End Sub
                                                               
'  This method creates a snipe
'  Inputs:
'    oCollar                     ---- Collar object on which to place snipe or drain hole
'    eCtx                        ---- Base or Offset face (CTX_BASE,CTX_OFFSET)
'    eBottomXid              ---- Support1 port xid  ( JXSEC_BOTTOM, JXSEC_BOTTOM_LEFT,JXSEC_BOTTOM_RIGHT)
'    eWebXid                  ---- Support2 port xid (JXSEC_LEFT, JXSEC_RIGHT)
'    oObjectCreated        ---- Snipe created
Public Sub CreateCornerSnipe(ByVal oCollar As Object, _
                                             ByVal oResourceManager As IUnknown, _
                                             ByVal eCtx As eUSER_CTX_FLAGS, _
                                             ByVal eBottomXid As JXSEC_CODE, _
                                             ByVal eWebXid As JXSEC_CODE, _
                                             ByRef oObjectCreated As Object)

    On Error GoTo ErrorHandler
    
    CreateSnipeOrDrainHole oCollar, _
                                        oResourceManager, _
                                        eCtx, _
                                        eBottomXid, _
                                        eWebXid, _
                                        "SnipeOnCollar", _
                                        oObjectCreated
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CreateCornerSnipe").Number
End Sub

'  This method creates a drain hole
'  Inputs:
'    oCollar                     ---- Collar object on which to place snipe or drain hole
'    eCtx                        ---- Base or Offset face (CTX_BASE or CTX_OFFSET)
'    eBottomXid              ---- Support1 port xid  ( JXSEC_BOTTOM, JXSEC_BOTTOM_LEFT,JXSEC_BOTTOM_RIGHT)
'    eWebXid                  ---- Support2 port xid (JXSEC_LEFT, JXSEC_RIGHT)
'    oObjectCreated        ---- Drain hole created
 
Public Sub CreateDrainHole(ByVal oCollar As Object, _
                                         ByVal oResourceManager As IUnknown, _
                                         ByVal eCtx As eUSER_CTX_FLAGS, _
                                         ByVal eBottomXid As JXSEC_CODE, _
                                         ByVal eWebXid As JXSEC_CODE, _
                                         ByRef oObjectCreated As Object)
    On Error GoTo ErrorHandler
    
    CreateSnipeOrDrainHole oCollar, _
                                        oResourceManager, _
                                        eCtx, _
                                        eBottomXid, _
                                        eWebXid, _
                                        "DrainHole", _
                                        oObjectCreated
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "CreateDrainHole").Number
End Sub
 
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
'$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
Public Sub ConstructBaseEdgePC(ByVal eBaseEdgeId As GSCADSDCreateModifyUtilities.JXSEC_CODE, _
                               ByVal oMD As IJDMemberDescription, _
                               ByVal oResourceManager As IUnknown, _
                               ByRef oObject As Object)
    '
    ' Create Physical Connection between Collar Edge and Base Plate
    Dim oBasePlatePort As IJPort
    Dim oCollarBottomPort As IJPort
    Dim oLastBasePlatePort As IJPort
    
    Dim oSDO_Helper As StructDetailObjects.Helper
    Dim oSDO_Collar As StructDetailObjects.Collar
    Dim oSDO_PhysicalConn As StructDetailObjects.PhysicalConn
    
    On Error GoTo ErrorHandler
    
    Set oSDO_Collar = New StructDetailObjects.Collar
    Set oSDO_Collar.object = oMD.CAO
    Set oCollarBottomPort = oSDO_Collar.SubPort(eBaseEdgeId)
    Set oBasePlatePort = oSDO_Collar.BasePlatePort
    Set oSDO_Collar = Nothing
    
    
    Set oSDO_Helper = New StructDetailObjects.Helper
    Set oLastBasePlatePort = oSDO_Helper.GetEquivalentLastPort(oBasePlatePort)
    Set oSDO_Helper = Nothing
    
    ' Construct PC
    Set oSDO_PhysicalConn = New StructDetailObjects.PhysicalConn
    oSDO_PhysicalConn.Create oResourceManager, oCollarBottomPort, oLastBasePlatePort, _
                             "TeeWeld", oMD.CAO, ConnectionStandard
    Set oObject = oSDO_PhysicalConn.object
    Set oSDO_PhysicalConn = Nothing
    
    Exit Sub
  
ErrorHandler:
    Err.Raise LogError(Err, MODULE, "ConstructBaseEdgePC").Number
End Sub
 

