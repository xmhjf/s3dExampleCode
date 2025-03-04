Attribute VB_Name = "GetStablePort"
Option Explicit

'   tr 70168 - check for stable port if avaliable - needed for IntelliShip
'
'   While selecting an external surface for creating a bounded eqp fndn,
'   SmartSketch gives us the located object. We have to check whether a StablePort
'   can be obtained, if so we return the StablePort.
'   Else we return the same located object.
'
Private Const MODULE = "GetStablePort:: "

Public Function GetStablePortIfApplicable(oLocObject As Object) As Object
Const METHOD = "GetStablePortIfApplicable"
On Error GoTo ErrorHandler
    
    Dim oRetVal         As Object
    Dim oLocatedPort    As IJPort
    Dim oStablePort     As IJPort
    Dim oIJStablePort   As IJStablePort
    Dim oConnectable    As IJConnectable
    ' This is the value we return by default, if we don't find a stable port
    Set oRetVal = oLocObject

    If Not oLocObject Is Nothing Then
        If TypeOf oLocObject Is IJPlane _
                    And Not TypeOf oLocObject Is IJConnectable _
                    And TypeOf oLocObject Is IJPort Then
                    
            Set oLocatedPort = oLocObject
            Set oConnectable = oLocatedPort.Connectable
            
            If TypeOf oConnectable Is IJStablePort Then
                
                Set oIJStablePort = oConnectable
                
                ' as per GSCAD dev this could return NULL
                ' hence adding resume next and changing back to err handler for any other errors
                On Error Resume Next
                Set oStablePort = oIJStablePort.StablePort(oLocatedPort)
                On Error GoTo ErrorHandler
                If Not oStablePort Is Nothing Then
                    If TypeOf oStablePort Is IJPlane Then
                        Set oRetVal = oStablePort
                    End If
                End If
                
            End If
            
        End If
    End If
    
    Set GetStablePortIfApplicable = oRetVal

    Exit Function
    
ErrorHandler:
    HandleError MODULE, METHOD
End Function




