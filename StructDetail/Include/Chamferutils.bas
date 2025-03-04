Attribute VB_Name = "Chamferutils"
Option Explicit

Private Const MODULE = "S:\StructDetail\Data\SmartOccurrence\AssyConnRules\Chamferutils.bas"
'

'***********************************************************************
' METHOD:  ChamferOffsetNormalData
'
' DESCRIPTION:  Compare the normal of the molded surface of two
'               connected plate parts. If they point to the same
'               direction, then return true. If not, return false.
'***********************************************************************
Public Sub ChamferOffsetNormalData( _
                                oSDO_ChamferPart1 As StructDetailObjects.PlatePart, _
                                oSDO_ChamferPart2 As StructDetailObjects.PlatePart, _
                                dBaseDelta As Double, _
                                dOffsetDelta As Double, _
                                bNormalConsistent As Boolean)
Const METHOD = "::ChamferOffsetNormalData"
On Error GoTo ErrorHandler
    
    Dim dDelta As Double
    Dim dChamfer1OffsetBase As Double
    Dim dChamfer1OffsetOffset As Double
    
    Dim dChamfer2OffsetBase As Double
    Dim dChamfer2OffsetOffset As Double

    Dim oOtherConnectable As IJConnectable
    Dim oChamferConnectable As IJConnectable
    
    ' Determine if the Plates Molded Surface Normals are Consistent
    Set oChamferConnectable = oSDO_ChamferPart1.object
    Set oOtherConnectable = oSDO_ChamferPart2.object
    If oChamferConnectable Is oOtherConnectable Then
        bNormalConsistent = True
    Else
        bNormalConsistent = CompareNormalOfPlateParts(oChamferConnectable, _
                                                      oOtherConnectable)
    End If
    
    ' calculate the Chamfer Offsets distances
    If bNormalConsistent Then
        dChamfer1OffsetBase = oSDO_ChamferPart1.OffsetToBaseFace
        dChamfer1OffsetOffset = oSDO_ChamferPart1.OffsetToOffsetFace

        dChamfer2OffsetBase = oSDO_ChamferPart2.OffsetToBaseFace
        dChamfer2OffsetOffset = oSDO_ChamferPart2.OffsetToOffsetFace
    
        dBaseDelta = dChamfer1OffsetBase - dChamfer2OffsetBase
        dOffsetDelta = dChamfer1OffsetOffset - dChamfer2OffsetOffset
        
'#####################################################################################
'#####################################################################################
'zMsgBox "ChamferUtils::ChamferOffsetNormalData ... " & vbCrLf & _
        "... oSDO_ChamferPart1= " & oSDO_ChamferPart1.Name & vbCrLf & _
        "... oSDO_ChamferPart2= " & oSDO_ChamferPart2.Name & vbCrLf & _
        "... bNormalConsistent= " & bNormalConsistent & vbCrLf & _
        "... oSD_ChamferPart1.OffsetToBaseFace= " & dChamfer1OffsetBase & vbCrLf & _
        "... oSD_ChamferPart2.OffsetToBaseFace= " & dChamfer2OffsetBase & vbCrLf & _
        "...dBaseDelta  =" & dBaseDelta & vbCrLf & _
        "... oSD_ChamferPart1.OffsetToOffsetFace= " & dChamfer1OffsetOffset & vbCrLf & _
        "... oSD_ChamferPart2.OffsetToOffsetFace= " & dChamfer2OffsetOffset & vbCrLf & _
        "...dOffsetDelta=" & dOffsetDelta
'#####################################################################################
'#####################################################################################
    
    Else
        dChamfer1OffsetBase = oSDO_ChamferPart1.OffsetToBaseFace
        dChamfer1OffsetOffset = oSDO_ChamferPart1.OffsetToOffsetFace

        dChamfer2OffsetBase = oSDO_ChamferPart2.OffsetToBaseFace
        dChamfer2OffsetOffset = oSDO_ChamferPart2.OffsetToOffsetFace
    
        ' Know that ChamferPart1 "Base" Port is aling with ChamferPart2 "Offset" Port
        ' Determine if Part1/Part2 Base/Offset Ports require Chamfer
        dBaseDelta = 0#
        dOffsetDelta = 0#
        dDelta = Abs(dChamfer1OffsetBase) - Abs(dChamfer2OffsetOffset)
        If Abs(dDelta) > 0.0001 Then
            If Abs(dChamfer1OffsetBase) > Abs(dChamfer2OffsetOffset) Then
                dBaseDelta = Abs(dChamfer2OffsetOffset) - Abs(dChamfer1OffsetBase)
            Else
                dOffsetDelta = Abs(dChamfer1OffsetBase) - Abs(dChamfer2OffsetOffset)
            End If
        End If
        
        ' Know that ChamferPart1 "Offset" Port is aling with ChamferPart2 "Base" Port
        ' Determine if Part1/Part2 Offset/Base Ports require Chamfer
        dDelta = Abs(dChamfer1OffsetOffset) - Abs(dChamfer2OffsetBase)
        If Abs(dDelta) > 0.0001 Then
            If Abs(dChamfer1OffsetOffset) > Abs(dChamfer2OffsetBase) Then
                dDelta = Abs(dChamfer1OffsetOffset) - Abs(dChamfer2OffsetBase)
                If Abs(dDelta) > Abs(dOffsetDelta) Then
                    dOffsetDelta = dDelta
                End If
                
            Else
                dDelta = Abs(dChamfer2OffsetBase) - Abs(dChamfer1OffsetOffset)
                If Abs(dDelta) > Abs(dBaseDelta) Then
                    dBaseDelta = dDelta
                End If
            End If
        End If

'#####################################################################################
'#####################################################################################
'zMsgBox "ChamferUtils::ChamferOffsetNormalData ... " & vbCrLf & _
        "... oSDO_ChamferPart1= " & oSDO_ChamferPart1.Name & vbCrLf & _
        "... oSDO_ChamferPart2= " & oSDO_ChamferPart2.Name & vbCrLf & _
        "... bNormalConsistent= " & bNormalConsistent & vbCrLf & _
        "... oSD_ChamferPart1.OffsetToBaseFace  = " & dChamfer1OffsetBase & vbCrLf & _
        "... oSD_ChamferPart2.OffsetToOffsetFace= " & dChamfer2OffsetOffset & vbCrLf & _
        "...dBaseDelta  =" & dBaseDelta & vbCrLf & _
        "... oSD_ChamferPart1.OffsetToOffsetFace= " & dChamfer1OffsetOffset & vbCrLf & _
        "... oSD_ChamferPart2.OffsetToBaseFace  = " & dChamfer2OffsetBase & vbCrLf & _
        "...dOffsetDelta=" & dOffsetDelta
'#####################################################################################
'#####################################################################################
    End If
        
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Sub

'***********************************************************************
' METHOD:  CompareNormalOfPlateParts
'
' DESCRIPTION:  Compare the normal of the molded surface of two
'               connected plate parts. If they point to the same
'               direction, then return true. If not, return false.
'***********************************************************************
Public Function CompareNormalOfPlateParts(oPlatePart1 As Object, _
                                          oPlatePart2 As Object) As Boolean
                                                              
Const METHOD = "::CompareNormalOfPlateParts"
On Error GoTo ErrorHandler

    Dim oSDO_PlatePart As StructDetailObjects.PlatePart
    
    Set oSDO_PlatePart = New StructDetailObjects.PlatePart
    Set oSDO_PlatePart.object = oPlatePart1
    CompareNormalOfPlateParts = oSDO_PlatePart.CompareNormalOfMoldedSurfaces(oPlatePart2)
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Function

'***********************************************************************
' METHOD:  CheckMultipleAssemblyConnections
'
' DESCRIPTION: This function is called from the Conditionals for creating a Chamfer
'              for the Plate Edge to Plate Edge cases
'               see AssyConnRules.PlateEdgeToPlateEdgeDef class
'               and methods CMCreateChamfer1 and CMCreateChamfer1
'
' Check if there are Muiltple Assembly Connections between the two parts
' if there are Muiltple Assembly Connections this indicates that is
' is a special case of Edge to Edge that does NOT require a Chamfer
'
'                           |---------------------
' Connect2(offset)          |
'   ======================= |
' Connect2 (base)        || |  Connect1 Edge
'   ======================= |
'   ------------------------|
'
' In this case (Built up bounded by Built up)
' There are Assembly Connections:
'       Between Connect1 Edge and Connect2 Edge
'       Between Connect1 Edge and Connect2 (base)
' Therefore,
' we only want a Physical Connection Between Connect1 Edge and Connect2 Edge
'***********************************************************************
Public Function CheckMultipleAssemblyConnections(oConnectedObject1 As Object, _
                                                 oConnectedPort1 As Object, _
                                                 oConnectedObject2 As Object, _
                                                 oConnectedPort2 As Object) _
                                                 As Boolean
Const METHOD = "::CheckMultipleAssemblyConnections"
On Error GoTo ErrorHandler
   
    Dim nCount As Long
    Dim nOtherConnections As Long
    Dim bConnected As Boolean
    
    Dim oObject As Object
    Dim oConnections As IJElements
    
    Dim oPort As IJPort
    Dim oConnectable1 As IJConnectable
    Dim oConnectable2 As IJConnectable
    Dim oAssemblyPort1 As Object
    Dim oAssemblyPort2 As Object
    Dim oAssemblyPart1 As Object
    Dim oAssemblyPart2 As Object
    
    Dim oSDO_AssemblyConn As StructDetailObjects.AssemblyConn
    
    ' Initialize the return value.
    CheckMultipleAssemblyConnections = False
    
    ' Check/verify the given onput Objects (IJConnectable)
    If oConnectedObject1 Is Nothing Then
        Exit Function
    ElseIf oConnectedObject2 Is Nothing Then
        Exit Function
    End If
    
    If Not TypeOf oConnectedObject1 Is IJConnectable Then
        Exit Function
    ElseIf Not TypeOf oConnectedObject2 Is IJConnectable Then
        Exit Function
    End If
    
    ' Do Not check for Muiltiple Connections if Ports are the same
    If oConnectedPort1 Is Nothing Then
    ElseIf oConnectedPort2 Is Nothing Then
    ElseIf oConnectedPort1 Is oConnectedPort2 Then
        Exit Function
    End If
    
    ' Get all Connections to between oConnectable1 and oConnectable2
    Set oConnectable1 = oConnectedObject1
    Set oConnectable2 = oConnectedObject2
    oConnectable1.isConnectedTo oConnectable2, bConnected, oConnections

    ' If there is less than 2 connections between two parts, exit the function.
    nCount = oConnections.Count
    If nCount < 2 Then
        Exit Function
    End If
    
    ' Loop through all the connection to find out how many assembly connections.
    nOtherConnections = 0
    For Each oObject In oConnections
        'Check/verify that Connection is not the given (oAssemblyConnection)
        If oObject Is Nothing Then
        Else
            If TypeOf oObject Is IJAssemblyConnection Then
                If nOtherConnections = 0 Then
                    Set oSDO_AssemblyConn = New StructDetailObjects.AssemblyConn
                End If
                
                ' Get the Ports and Connectables from current Assembly Connection
                Set oAssemblyPart1 = Nothing
                Set oAssemblyPart2 = Nothing
                Set oSDO_AssemblyConn.object = oObject
                Set oAssemblyPort1 = oSDO_AssemblyConn.Port1
                If TypeOf oAssemblyPort1 Is IJPort Then
                    Set oPort = oAssemblyPort1
                    Set oAssemblyPart1 = oPort.Connectable
                    Set oPort = Nothing
                End If
                
                Set oAssemblyPort2 = oSDO_AssemblyConn.Port2
                If TypeOf oAssemblyPort2 Is IJPort Then
                    Set oPort = oAssemblyPort2
                    Set oAssemblyPart2 = oPort.Connectable
                    Set oPort = Nothing
                End If
                
                ' Check if this Assembly Connection Ports match either of the Given
                If oConnectedPort1 Is Nothing And oConnectedPort2 Is Nothing Then
                    nOtherConnections = nOtherConnections + 1
                
                ElseIf oAssemblyPort1 Is Nothing Then
                ElseIf oAssemblyPort2 Is Nothing Then
                
                ElseIf oAssemblyPort1 Is oConnectedPort1 Then
                    If oAssemblyPart2 Is oConnectable2 Then
                        nOtherConnections = nOtherConnections + 1
                    End If
                ElseIf oAssemblyPort1 Is oConnectedPort2 Then
                    If oAssemblyPart2 Is oConnectable1 Then
                        nOtherConnections = nOtherConnections + 1
                    End If
                    
                ElseIf oAssemblyPort2 Is oConnectedPort1 Then
                    If oAssemblyPart1 Is oConnectable2 Then
                        nOtherConnections = nOtherConnections + 1
                    End If
                
                ElseIf oAssemblyPort2 Is oConnectedPort2 Then
                    If oAssemblyPart2 Is oConnectable1 Then
                        nOtherConnections = nOtherConnections + 1
                    End If
                End If
            End If
        End If
    
        Set oObject = Nothing
    Next
    
    If nOtherConnections > 1 Then
        CheckMultipleAssemblyConnections = True
    End If
    
    Exit Function
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD).Number
End Function

Public Sub Create_ChamferEdge(ByVal pMemberDescription As IJDMemberDescription, _
                              ByVal pResourceManager As IUnknown, _
                              sRootClass As String, _
                              ByRef pChamferObject As Object)
Const METHOD = "::Create_ChamferEdge"
    On Error GoTo ErrorHandler
    
    Dim sRoot As String
    Dim sERROR As String
    
    Dim oPart1 As Object
    Dim oPart2 As Object
    Dim oPort1 As Object
    Dim oPort2 As Object
    Dim oParent As Object
    
    Dim oSDO_Chamfer As StructDetailObjects.Chamfer
    Dim oSDO_AssyConn As StructDetailObjects.AssemblyConn
    
    ' Initialize wrapper class and get the 2 ports
    sERROR = "Setting Assembly Connection Inputs"
    Set oSDO_AssyConn = New StructDetailObjects.AssemblyConn
    Set oSDO_AssyConn.object = pMemberDescription.CAO
    
    sERROR = "Getting Assembly Connection Parts associated to ports"
    Set oPort1 = oSDO_AssyConn.Port1
    Set oPort2 = oSDO_AssyConn.Port2
    
    Set oPart1 = oSDO_AssyConn.ConnectedObject1
    Set oPart2 = oSDO_AssyConn.ConnectedObject2
    
    ' Set the Assembly connection as parent
    sERROR = "Setting system parent to Member Description Custom Assembly"
    Set oParent = pMemberDescription.CAO
       
    ' Create Chamfer between Plate Edge and Edge Reinforcement
    sERROR = "Creating Chamfer"
    Set oSDO_Chamfer = New StructDetailObjects.Chamfer
    If TypeOf oPart1 Is IJPlate Then
        oSDO_Chamfer.Create pResourceManager, oPort1, oPort2, sRootClass, oParent
    Else
        oSDO_Chamfer.Create pResourceManager, oPort2, oPort1, sRootClass, oParent
    End If
                               
    sERROR = "Setting Chamfer to private variable"
    Set pChamferObject = oSDO_Chamfer.object
    
    Exit Sub
    
ErrorHandler:
    Err.Raise LogError(Err, MODULE, METHOD, sERROR).Number
End Sub

