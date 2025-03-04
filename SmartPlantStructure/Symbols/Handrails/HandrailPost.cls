VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HandrailPost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'******************************************************************
' Copyright (C) 2010, Intergraph Corporation. All rights reserved.
'
'File
'    HandrailPost.cls
'
'Author
'       GG
'
'Description
' As the data holder for the information of post or treatment

'Notes
'
'History:
'*******************************************************************

'local variable(s) to hold property value(s)
Private mvarBasePos As DPosition 'local copy
Private mvarDirectionVec As DVector 'local copy
Private mvarPostType As Integer 'local copy '2: Reguler Post; 5: Circular Treatment; 10: Rectangular Treatment
'local variable(s) to hold property value(s)
Private mvarIndex As Integer 'local copy
'local variable(s) to hold property value(s)
Private mvarSectionAngle As Double 'local copy
'local variable(s) to hold property value(s)
Private mvarPathVec As DVector 'local copy
Public Property Set PathVec(ByVal vData As DVector)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.PathVec = Form1
    Set mvarPathVec = vData
End Property


Public Property Get PathVec() As DVector
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PathVec
    Set PathVec = mvarPathVec
End Property



Public Property Let SectionAngle(ByVal vData As Double)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SectionAngle = 5
    mvarSectionAngle = vData
End Property


Public Property Get SectionAngle() As Double
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SectionAngle
    SectionAngle = mvarSectionAngle
End Property



Public Property Let Index(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Index = 5
    mvarIndex = vData
End Property


Public Property Get Index() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Index
    Index = mvarIndex
End Property



Public Property Let PostType(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.PostType = 5
    mvarPostType = vData
End Property


Public Property Get PostType() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.PostType
    PostType = mvarPostType
End Property



Public Property Set DirectionVec(ByVal vData As DVector)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.TopPos = Form1
    Set mvarDirectionVec = vData
End Property


Public Property Get DirectionVec() As DVector
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TopPos
    Set DirectionVec = mvarDirectionVec
End Property



Public Property Set BasePos(ByVal vData As DPosition)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.BasePos = Form1
    Set mvarBasePos = vData
End Property


Public Property Get BasePos() As DPosition
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.BasePos
    Set BasePos = mvarBasePos
End Property




Private Sub Class_Initialize()
'Set mvarBasePos = New DPosition 'local copy
'Set mvarDirectionVec = New DVector 'local copy

End Sub

Private Sub Class_Terminate()
'Set mvarBasePos = Nothing 'local copy
'Set mvarDirectionVec = Nothing 'local copy

End Sub
