VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSimplePhysical"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SP3DEqpUSSClassType" ,"OTHER"
Attribute VB_Ext_KEY = "SP3DV6UpgradeSO" ,"Upgraded by Eqp SO Upgrade Wizard at 11/29/2004-5:18:22 AM"
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'
'   Copyright (c) 2003, Intergraph Corporation. All rights reserved.
'
'   CSimplePhysical.cls
'   Author:          KV
'   Creation Date:  Monday, Dec 9 2002
'   Description:
' This class module is the place for user to implement graphical part of VBSymbol for this aspect
'
'   Change History:
'   dd.mmm.yyyy     who     change description
'   -----------         -----        ------------------
'   09.Jul.2003     SymbolTeam(India)       Copyright Information, Header  is added.
'   29.Nov.2004     V6UpgradeSO        Made compatible with Smart Occurrence based Equipments
'  08.SEP.2006     KKC  DI-95670  Replace names with initials in all revision history sheets and symbols
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

Option Explicit
Private Const MODULE = "Physical:" 'Used for error messages

Private Sub Class_Initialize()
'''
End Sub

Public Sub run(ByVal m_outputColl As Object, ByRef arrayOfInputs(), arrayOfOutputs() As String)
    
    Const METHOD = "run"
    On Error GoTo ErrorLabel
          
    Dim objTrailerBodyCyli          As Object
    Dim objTrailerBodyBox           As Object
    Dim objTrailerBodyRtor          As Object
    Dim objTrailerBodyPyra          As Object
    Dim objTrailerTireCyli          As Object
    Dim objTrailerBodyDish          As Object
    Dim objTrailerBodyCtor          As Object
    
    Dim iOutput                     As Integer
    Dim iCount                      As Integer
    
    Dim myOri                       As Orientation
    Set myOri = New Orientation
          
    Dim geomFactory                 As New IngrGeom3D.GeometryFactory

    Dim CenterPos                   As New AutoMath.DPosition
      CenterPos.Set 0.254, -5.0038, 1.14618
    Dim startPoint                  As New AutoMath.DPosition
      startPoint.Set 0.254, 1.0922, 1.14618
    Dim endPoint                    As New AutoMath.DPosition
      endPoint.Set 0.254, -11.0998, 1.14618
   
    iOutput = 0

' Insert your code for output 1(Trailer Body Cyli) Main Body Cylinder of Trailer
      iCount = 1
      Set objTrailerBodyCyli = CreateCylinder(m_outputColl, "E 0.0 S 5.2578 U 0.89218", 1.8288, 12.192, "N 0 U 0")
' Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyCyli
      Set objTrailerBodyCyli = Nothing
          
' Insert your code for output 2(Trailer Body Box) bottom support for cylinder
      iCount = iCount + 1
      myOri.ResetDefaultAxis
      Set objTrailerBodyPyra = CreatePyramid(m_outputColl, "W 0.00159 S 5.2578 U 0.2159", myOri, 1.272, 12.192, 1.7272, 12.192, 0, 0, 0.762)

' Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyPyra
      Set objTrailerBodyPyra = Nothing
      
' Insert your code for output 3(Trailer Body Cyli)  Tire Driver's Side Front
      iCount = iCount + 1
      Set objTrailerTireCyli = CreateCylinder(m_outputColl, "W 0.9271 S 9.6647 D 0.80645", 1.0287, 0.635, "E 90 U 0")

' Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerTireCyli
      Set objTrailerTireCyli = Nothing
      
' Insert your code for output 4(Trailer Body Cyli)  Tire Driver's Side Rear
      iCount = iCount + 1
      Set objTrailerTireCyli = CreateCylinder(m_outputColl, "W 0.9271 S 10.9474 D 0.80645", 1.0287, 0.635, "E 90 U 0")

' Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerTireCyli
      Set objTrailerTireCyli = Nothing
      
' Insert your code for output 5(Trailer Body Cyli)  Tire Passenger's Side Front
      iCount = iCount + 1
      Set objTrailerTireCyli = CreateCylinder(m_outputColl, "E 0.9271 S 9.6647 D 0.80645", 1.0287, 0.635, "E 90 U 0")
' Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerTireCyli
      Set objTrailerTireCyli = Nothing
      
' Insert your code for output 6(Trailer Body Cyli)  Tire Passenger's Side Rear
      iCount = iCount + 1
      Set objTrailerTireCyli = CreateCylinder(m_outputColl, "E 0.9271 S 10.9474 D 0.80645", 1.0287, 0.635, "E 90 U 0")

' Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerTireCyli
      Set objTrailerTireCyli = Nothing

' Insert your code for output 7(Trailer Body Box)  Front Trailer Support Box
      iCount = iCount + 1
      Set objTrailerBodyBox = CreateBox(m_outputColl, "W 0.00159 S 1.33509 D 0.28099", myOri, 0.9652, 4.2672002, 0.2286, True)

' Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyBox
      Set objTrailerBodyBox = Nothing

' Insert your code for output 8(Trailer Body Box)  Upper Rear Trailer Support Box
      iCount = iCount + 1
      Set objTrailerBodyBox = CreateBox(m_outputColl, "W 0.44609 S 9.18845 D 0.36989", myOri, 0.0762, 5.08, 0.381, True)

' Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyBox
      Set objTrailerBodyBox = Nothing

' Insert your code for output 9(Trailer Body Box)  Upper Rear Trailer Support Box
      iCount = iCount + 1
      Set objTrailerBodyBox = CreateBox(m_outputColl, "E 0.44609 S 9.18845 D 0.36989", myOri, 0.0762, 5.08, 0.381, True)

' Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyBox
      Set objTrailerBodyBox = Nothing

' Insert your code for output 10(Trailer Body Box)  Lower Rear Trailer Bumper Support Box1
      iCount = iCount + 1
      Set objTrailerBodyBox = CreateBox(m_outputColl, "E 0.33496 S 10.82993 D 0.635", myOri, 0.1524, 2.032, 0.1524, True)
       
' Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyBox
      Set objTrailerBodyBox = Nothing

' Insert your code for output 11(Trailer Body Box)  Lower Rear Trailer Bumper Support Box2
      iCount = iCount + 1
      Set objTrailerBodyBox = CreateBox(m_outputColl, "W 0.32067 S 10.82993 D 0.635", myOri, 0.1524, 2.032, 0.1524, True)
       
' Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyBox
      Set objTrailerBodyBox = Nothing

' Insert your code for output 12(Trailer Body Box)  Rear Trailer Bumper
      iCount = iCount + 1
      Set objTrailerBodyBox = CreateBox(m_outputColl, "W 0.01588 S 11.8618 D 0.65881", myOri, 1.6256, 0.05715, 0.2286, True)
       
' Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyBox
      Set objTrailerBodyBox = Nothing
      
' Insert your code for output 13(Trailer Body Box)  Rear Trailer Tail-light Support Driver's side
      iCount = iCount + 1
      Set objTrailerBodyBox = CreateBox(m_outputColl, "W 0.82868 S 11.6713 D 0.2921", myOri, 0.6858, 0.1016, 0.2032, True)
       
' Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyBox
      Set objTrailerBodyBox = Nothing

' Insert your code for output 14(Trailer Body Box)  Rear Trailer Tail-light Support Passenger's side
      iCount = iCount + 1
      Set objTrailerBodyBox = CreateBox(m_outputColl, "E 0.8255 S 11.6713 D 0.2921", myOri, 0.6858, 0.1016, 0.2032, True)
       
' Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyBox
      Set objTrailerBodyBox = Nothing

' Insert your code for output 15(Trailer Body Box)  Wheel Cover Rear Trailer Wheel Cover Driver's side
      iCount = iCount + 1
      Set objTrailerBodyBox = CreateBox(m_outputColl, "W 0.87154 S 10.30288 D 0.2159", myOri, 0.8016875, 1.3716, 0.0508, True)
       
' Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyBox
      Set objTrailerBodyBox = Nothing
      
' Insert your code for output 16(Trailer Body Box)  Wheel Cover Rear Trailer Wheel Cover Passenger's side
      iCount = iCount + 1
      Set objTrailerBodyBox = CreateBox(m_outputColl, "E 0.87948 S 10.30288 D 0.2159", myOri, 0.8016875, 1.3716, 0.0508, True)
       
' Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyBox
      Set objTrailerBodyBox = Nothing

' Insert your code for output 17(trailer body Cyli)tail light outside driver side
      iCount = 1
      Set objTrailerBodyCyli = CreateCylinder(m_outputColl, "W 1.00013 S 11.71734 D 0.2921", 0.127, 0.0254, "N 0 U 0")
      
' Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyCyli
      Set objTrailerBodyCyli = Nothing
      
' Insert your code for output 18(trailer body Cyli) tail light inside driver side
      iCount = 1
      Set objTrailerBodyCyli = CreateCylinder(m_outputColl, "W 0.73978 S 11.71734 D 0.2921", 0.127, 0.0254, "N 0 U 0")
      
' Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyCyli
      Set objTrailerBodyCyli = Nothing
      
' Insert your code for output 19(trailer body Cyli) tail light inside passenger side
      iCount = 1
      Set objTrailerBodyCyli = CreateCylinder(m_outputColl, "E 0.65405 S 11.71734 D 0.2921", 0.127, 0.0254, "N 0 U 0")
      
' Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyCyli
      Set objTrailerBodyCyli = Nothing
      
' Insert your code for output 20(trailer body Cyli) tail light outside passenger side
      iCount = 1
      Set objTrailerBodyCyli = CreateCylinder(m_outputColl, "E 0.91281 S 11.71734 D 0.2921", 0.127, 0.0254, "N 0 U 0")
      
' Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyCyli
      Set objTrailerBodyCyli = Nothing

' Insert your code for output 21(trailer body Rtor) wheel cover front driver side rtor
 
      iCount = 1
      Set objTrailerBodyRtor = CreateRectangularTorus(m_outputColl, "W 0.889 S 9.61866 D 0.77311", 0.5334, 0.5842, 90, 0, 0.762, "N 0 U 90", "S 180 U 0")

' Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyRtor
      Set objTrailerBodyRtor = Nothing

' Insert your code for output 22(trailer body Rtor) wheel cover rear driver side rtor
      iCount = 1
      Set objTrailerBodyRtor = CreateRectangularTorus(m_outputColl, "W 0.889 S 10.98233 D 0.76994", 0.5334, 0.5842, 90, 0, 0.762, "S 180 U 0", "N 0 D -90")
' Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyRtor
      Set objTrailerBodyRtor = Nothing

' Insert your code for output 23(trailer body Rtor) wheel cover front passenger side rtor
      iCount = 1
      Set objTrailerBodyRtor = CreateRectangularTorus(m_outputColl, "E 0.889 S 9.61866 D 0.77311", 0.5334, 0.5842, 90, 0, 0.762, "N 0 U 90", "S 180 U 0")
      
' Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyRtor
      Set objTrailerBodyRtor = Nothing

' Insert your code for output 24(trailer body Rtor) wheel cover rear passenger side rtor
      iCount = 1
      Set objTrailerBodyRtor = CreateRectangularTorus(m_outputColl, "E 0.889 S 10.98233 D 0.76994", 0.5334, 0.5842, 90, 0, 0.762, "S 180 U 0", "N 0 D -90")
      
' Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyRtor
      Set objTrailerBodyRtor = Nothing

' Insert your code for output 25(Trailer body box)  support box for front trailer supports
      iCount = iCount + 1
      Set objTrailerBodyBox = CreateBox(m_outputColl, "E 0.0 S 2.87338 D 0.57309", myOri, 0.9652, 1.2192, 0.3556, True)
       
' Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyBox
      Set objTrailerBodyBox = Nothing

' Insert your code for output 26(Trailer body box)  support leg driver side
      iCount = iCount + 1
      Set objTrailerBodyBox = CreateBox(m_outputColl, "W 0.54927 S 2.85433 D 0.83979", myOri, 0.127, 0.127, 0.762, True)
       
' Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyBox
      Set objTrailerBodyBox = Nothing

' Insert your code for output 27(Trailer body box)  support leg passenger side
      iCount = iCount + 1
      Set objTrailerBodyBox = CreateBox(m_outputColl, "E 0.54769 S 2.85433 D 0.83979", myOri, 0.127, 0.127, 0.762, True)
       
' Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyBox
      Set objTrailerBodyBox = Nothing

' Insert your code for output 28(Trailer body box)  support foot driver side
      iCount = iCount + 1
      Set objTrailerBodyBox = CreateBox(m_outputColl, "W 0.54927 S 2.85433 D 1.25119", myOri, 0.3048, 0.3048, 0.0254, True)
       
' Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyBox
      Set objTrailerBodyBox = Nothing

' Insert your code for output 29(Trailer body box)  support foot passenger side
      iCount = iCount + 1
      Set objTrailerBodyBox = CreateBox(m_outputColl, "E 0.54769 S 2.85433 D 1.25119", myOri, 0.3048, 0.3048, 0.0254, True)
       
' Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyBox
      Set objTrailerBodyBox = Nothing

' Insert your code for output 30(Trailer body box)  box for trailer label on passenger side
      iCount = iCount + 1
      myOri.ResetDefaultAxis
      myOri.RotationAboutY = 45
      myOri.RotationAboutZ = 90
      myOri.ApplyRotations
      Set objTrailerBodyBox = CreateBox(m_outputColl, "E 0.9144 S 0.65088 U 0.9144", myOri, 0.3048, 0.0047625, 0.3048, True)
       
' Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyBox
      Set objTrailerBodyBox = Nothing

' Insert your code for output 31(Trailer Body Cyli) Passenger side Rear leg support
      iCount = 1
      myOri.ResetDefaultAxis
      myOri.RotationAboutY = 45
      myOri.RotationAboutZ = 90
      myOri.ApplyRotations
      Set objTrailerBodyCyli = CreateCylinder(m_outputColl, "E 0.54769 S 3.19723 D 0.87154", 0.0381, 0.762, , myOri)
      
' Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyCyli
      Set objTrailerBodyCyli = Nothing
            
' Insert your code for output 32(Trailer Body Cyli) Passenger side Front leg support
      iCount = 1
      myOri.ResetDefaultAxis
      myOri.RotationAboutY = -45
      myOri.RotationAboutZ = 90
      myOri.ApplyRotations
      Set objTrailerBodyCyli = CreateCylinder(m_outputColl, "E 0.54769 S 2.51143 D 0.87154", 0.0381, 0.762, , myOri)

' Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyCyli
      Set objTrailerBodyCyli = Nothing

' Insert your code for output 33(Trailer Body Cyli) Driver side Rear leg support
      iCount = 1
      myOri.ResetDefaultAxis
      myOri.RotationAboutY = 45
      myOri.RotationAboutZ = 90
      myOri.ApplyRotations
      Set objTrailerBodyCyli = CreateCylinder(m_outputColl, "W 0.54769 S 3.19723 D 0.87154", 0.0381, 0.762, , myOri)

' Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyCyli
      Set objTrailerBodyCyli = Nothing

' Insert your code for output 34(Trailer Body Cyli) Driver side Front leg support
      iCount = 1
      myOri.ResetDefaultAxis
      myOri.RotationAboutY = -45
      myOri.RotationAboutZ = 90
      myOri.ApplyRotations
      Set objTrailerBodyCyli = CreateCylinder(m_outputColl, "W 0.54769 S 2.51143 D 0.87154", 0.0381, 0.762, , myOri)

 'Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyCyli
      Set objTrailerBodyCyli = Nothing
      
' Insert your code for output 35(Trailer Body Cyli) Top of Trailer access hole inside cyli
      iCount = 1
      Set objTrailerBodyCyli = CreateCylinder(m_outputColl, "E 0.00159 S 5.2578 U 1.7907", 0.6096, 0.1524, "N 0 U 90")

 'Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyCyli
      Set objTrailerBodyCyli = Nothing

' Insert your code for output 36(Trailer Body Cyli) Top of Trailer access hole outside cyli
      iCount = 1
      Set objTrailerBodyCyli = CreateCylinder(m_outputColl, "E 0.00159 S 5.2578 U 1.905", 0.762, 0.0254, "N 0 U 90")

 'Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyCyli
      Set objTrailerBodyCyli = Nothing
      
' Insert your code for output 37(Trailer Body Dish) Top of Trailer access hole
      iCount = 1
      Set objTrailerBodyDish = CreateDish(m_outputColl, "E 0.00159 S 5.2578 U 1.9177", "N 0 U 90", 0.6096, 0.0508, 0, False)
      
 'Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyDish
      Set objTrailerBodyDish = Nothing
      
' Insert your code for output 38(Trailer body box)  box surrounding center access hole in top of trailer
      iCount = iCount + 1
      myOri.ResetDefaultAxis
      Set objTrailerBodyBox = CreateBox(m_outputColl, "E 0.00159 S 5.2578 U 1.79705", myOri, 0.9144, 1.4732, 0.0381, True)
       
' Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyBox
      Set objTrailerBodyBox = Nothing
      
' Insert your code for output 39(Trailer Body Cyli) Top of Trailer access hole rail suport uprights
      iCount = 1
      Set objTrailerBodyCyli = CreateCylinder(m_outputColl, "E 0.42386 S 4.5466 U 1.94151", 0.0254, 0.254, "N 0 U 90")

 'Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyCyli
      Set objTrailerBodyCyli = Nothing
      
' Insert your code for output 40(Trailer Body Cyli) Top of Trailer access hole rail suport uprights
      iCount = 1
      Set objTrailerBodyCyli = CreateCylinder(m_outputColl, "W 0.43974 S 4.5466 U 1.94151", 0.0254, 0.254, "N 0 U 90")

 'Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyCyli
      Set objTrailerBodyCyli = Nothing
      
' Insert your code for output 41(Trailer Body Cyli) Top of Trailer access hole rail suport uprights
      iCount = 1
      Set objTrailerBodyCyli = CreateCylinder(m_outputColl, "W 0.43974 S 5.2578 U 1.94151", 0.0254, 0.254, "N 0 U 90")

 'Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyCyli
      Set objTrailerBodyCyli = Nothing
      
' Insert your code for output 42(Trailer Body Cyli) Top of Trailer access hole rail suport uprights
      iCount = 1
      Set objTrailerBodyCyli = CreateCylinder(m_outputColl, "W 0.43974 S 5.969 U 1.94151", 0.0254, 0.254, "N 0 U 90")

 'Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyCyli
      Set objTrailerBodyCyli = Nothing
      
' Insert your code for output 43(Trailer Body Cyli) Top of Trailer access hole rail suport uprights
      iCount = 1
      Set objTrailerBodyCyli = CreateCylinder(m_outputColl, "E 0.42386 S 5.969 U 1.94151", 0.0254, 0.254, "N 0 U 90")

 'Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyCyli
      Set objTrailerBodyCyli = Nothing
      
' Insert your code for output 44(Trailer Body Cyli) Top of Trailer access hole rail suport uprights
      iCount = 1
      Set objTrailerBodyCyli = CreateCylinder(m_outputColl, "E 0.42228 S 5.52768 U 1.91611", 0.0254, 0.2032, "N 0 U 90")

 'Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyCyli
      Set objTrailerBodyCyli = Nothing
      
' Insert your code for output 45(Trailer Body Cyli) Top of Trailer access hole rail suport uprights
      iCount = 1
      Set objTrailerBodyCyli = CreateCylinder(m_outputColl, "E 0.42228 S 4.9784 U 1.9177", 0.0254, 0.206375, "N 0 U 90")

 'Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyCyli
      Set objTrailerBodyCyli = Nothing

' Insert your code for output 46(Trailer Body Cyli) Top of Trailer access hole rails
      iCount = 1
      Set objTrailerBodyCyli = CreateCylinder(m_outputColl, "E 0.42228 S 4.7371 U 2.06851", 0.0254, 0.381, "S 180 U 0")

 'Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyCyli
      Set objTrailerBodyCyli = Nothing
      
' Insert your code for output 47(Trailer Body Cyli) Top of Trailer access hole rails
      iCount = 1
      Set objTrailerBodyCyli = CreateCylinder(m_outputColl, "W 0.00794 S 4.5466 U 2.06851", 0.0254, 0.8636, "E 90 U 0")

 'Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyCyli
      Set objTrailerBodyCyli = Nothing

' Insert your code for output 48(Trailer Body Cyli) Top of Trailer access hole rails
      iCount = 1
      Set objTrailerBodyCyli = CreateCylinder(m_outputColl, "W 0.43815 S 5.26415 U 2.06851", 0.0254, 1.4351, "S 180 U 0")

 'Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyCyli
      Set objTrailerBodyCyli = Nothing

' Insert your code for output 49(Trailer Body Cyli) Top of Trailer access hole rails
      iCount = 1
      Set objTrailerBodyCyli = CreateCylinder(m_outputColl, "W 0.00794 S 5.969 U 2.06851", 0.0254, 0.8636, "E 90 U 0")

 'Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyCyli
      Set objTrailerBodyCyli = Nothing

' Insert your code for output 50(Trailer Body Cyli) Top of Trailer access hole rails
      iCount = 1
      Set objTrailerBodyCyli = CreateCylinder(m_outputColl, "E 0.42228 S 5.7785 U 2.06851", 0.0254, 0.4064, "S 180 U 0")

 'Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyCyli
      Set objTrailerBodyCyli = Nothing

' Insert your code for output 51(Trailer Body Cyli) Trailer ladder side rail suport uprights rear
      iCount = 1
      Set objTrailerBodyCyli = CreateCylinder(m_outputColl, "E 1.0541 S 5.47211 U 0.28892", 0.0254, 1.1572875, "N 0 U 90")

 'Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyCyli
      Set objTrailerBodyCyli = Nothing
      
' Insert your code for output 52(Trailer Body Cyli) Trailer ladder side rail suport uprights
      iCount = 1
      Set objTrailerBodyCyli = CreateCylinder(m_outputColl, "E 1.0541 S 5.01491 U 0.28892", 0.0254, 1.1572875, "N 0 U 90")

 'Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyCyli
      Set objTrailerBodyCyli = Nothing
      
' Insert your code for output 53(Trailer Body Cyli) Trailer ladder steps starting at bottom
      iCount = 1
      Set objTrailerBodyCyli = CreateCylinder(m_outputColl, "E 1.05251 S 5.23557 D 0.27781", 0.0254, 0.4572, "N 0 U 0")

 'Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyCyli
      Set objTrailerBodyCyli = Nothing

' Insert your code for output 54(Trailer Body Cyli) Trailer ladder steps
      iCount = 1
      Set objTrailerBodyCyli = CreateCylinder(m_outputColl, "E 1.05251 S 5.23557 U 0.02699", 0.0254, 0.4572, "N 0 U 0")

 'Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyCyli
      Set objTrailerBodyCyli = Nothing
      
' Insert your code for output 55(Trailer Body Cyli) Trailer ladder steps
      iCount = 1
      Set objTrailerBodyCyli = CreateCylinder(m_outputColl, "E 1.05251 S 5.23557 U 0.33179", 0.0254, 0.4572, "N 0 U 0")

 'Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyCyli
      Set objTrailerBodyCyli = Nothing
      
' Insert your code for output 56(Trailer Body Cyli) Trailer ladder steps
      iCount = 1
      Set objTrailerBodyCyli = CreateCylinder(m_outputColl, "E 1.05251 S 5.23557 U 0.63659", 0.0254, 0.4572, "N 0 U 0")

 'Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyCyli
      Set objTrailerBodyCyli = Nothing
      
' Insert your code for output 57(Trailer Body Cyli) Trailer ladder steps
      iCount = 1
      Set objTrailerBodyCyli = CreateCylinder(m_outputColl, "E 1.04775 S 5.23557 U 0.94139", 0.0254, 0.4572, "N 0 U 0")

 'Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyCyli
      Set objTrailerBodyCyli = Nothing
      
' Insert your code for output 58(Trailer Body Cyli) Trailer ladder steps
      iCount = 1
      Set objTrailerBodyCyli = CreateCylinder(m_outputColl, "E 0.97631 S 5.23557 U 1.24301", 0.0254, 0.4572, "N 0 U 0")

 'Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyCyli
      Set objTrailerBodyCyli = Nothing
      
' Insert your code for output 59(Trailer Body Cyli) Trailer ladder steps
      iCount = 1
      Set objTrailerBodyCyli = CreateCylinder(m_outputColl, "E 0.83503 S 5.23557 U 1.46526", 0.0254, 0.4572, "N 0 U 0")

 'Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyCyli
      Set objTrailerBodyCyli = Nothing
      
' Insert your code for output 60(Trailer Body Cyli) Trailer ladder steps
      iCount = 1
      Set objTrailerBodyCyli = CreateCylinder(m_outputColl, "E 0.61119 S 5.23557 U 1.67005", 0.0254, 0.4572, "N 0 U 0")

 'Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyCyli
      Set objTrailerBodyCyli = Nothing
      
' Insert your code for output 61(Trailer Body Cyli) Trailer ladder steps finishing at top
      iCount = 1
      Set objTrailerBodyCyli = CreateCylinder(m_outputColl, "E 0.48101 S 5.24034 U 1.7399", 0.0254, 0.4572, "N 0 U 0")

 'Set the output
      iOutput = iOutput + 1
      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyCyli
      Set objTrailerBodyCyli = Nothing
      
'' Insert your code for output 62(Trailer Body Ctor) Trailer ladder Ctor outside rear support
'      iCount = 1
'      Set objTrailerBodyCtor = CreateCircularTorus(m_outputColl, "E 0.1016 S 5.47211 U 0.86678", "D", "W", 0.9398, 0.9652, 90)
'
' 'Set the output
'      iOutput = iOutput + 1
'      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyCtor
'      Set objTrailerBodyCtor = Nothing
'
'' Insert your code for output 63(Trailer Body Ctor) Trailer ladder Ctor outside front support
'      iCount = 1
'      Set objTrailerBodyCtor = CreateCircularTorus(m_outputColl, "E 0.1016 S 5.01491 U 0.86678", "D", "W", 0.9398, 0.9652, 90)
'
' 'Set the output
'      iOutput = iOutput + 1
'      m_outputColl.AddOutput arrayOfOutputs(iOutput), objTrailerBodyCtor
'      Set objTrailerBodyCtor = Nothing
      
    Exit Sub
    
ErrorLabel:
    ReportUnanticipatedError MODULE, METHOD
    Resume Next
    
End Sub

