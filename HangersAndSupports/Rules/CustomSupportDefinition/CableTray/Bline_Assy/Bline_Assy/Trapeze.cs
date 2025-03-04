//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Trapeze.cs
//   Bline_Assy,Ingr.SP3D.Content.Support.Rules.Trapeze
//   Author       :  Vijaya
//   Creation Date:  22.Nov.2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   22.Nov.2012    Vijaya   CR-CP-219114-Initial Creation
//   22-Jan-2015    PVK      TR-CP-264951  Resolve coverity issues found in November 2014 report
//   07.Sep.2015    PR   TR 277225	B-Line Hangers do not place correctly 
//   17/12/2015     Ramya   TR 284319	Multiple Record exception dumps are created on copy pasting supports
//   04/07/2016     Siva    DM-CP-296666 	B-Line Trapeze Assembly Support doesn't place on cableway turn feature. 
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.Generic;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Structure.Middle;
using Ingr.SP3D.Content.Support.Symbols;

namespace Ingr.SP3D.Content.Support.Rules
{
    //-----------------------------------------------------------------------------------
    //Namespace of this class is Ingr.SP3D.Content.Support.Rules
    //It is recommended that customers specify namespace of their Rules to be
    //<CompanyName>.SP3D.Content.<Specialization>.
    //It is also recommended that if customers want to change this Rule to suit their
    //requirements, they should change namespace/Rule name so the identity of the modified
    //Rule will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------
    public class Trapeze : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        //Constants
        private const string BeamClamp1 = "BeamClamp1";
        private const string BeamClamp2 = "BeamClamp2";
        private const string LeftHangerRod = "LeftHangerRod";
        private const string RightHangerRod = "RightHangerRod";
        private const string TrapezePart = "Trapeze";
        private string[] Nutpartkeys = new string[15];
        private string[] Trapezepartkeys = new string[15];
        private string partkey;
        private const double Inch = 0.0254;
        string trapezeType;
        int IntialParts;
        int trapeze_Begin;
        int trapeze_End;
        int numOfRoutes;
        private string refPortName;
        private double[] OriginX_Route = new double[5];
        private double[] OriginY_Route = new double[5];
        private double[] OriginZ_Route = new double[5];
        private double[] width = new double[5];
        private double[] height = new double[5];
        private double[] clampOffset = new double[5];
        private double temp;
        private double leftHgrRodOffset;
        private double rightHgrRodOffset;
        int leftFaceNumber;
        int rightFaceNumber;
        int StructFaceNumber;

        //-----------------------------------------------------------------------------------
        //Get Assembly Catalog Parts
        //-----------------------------------------------------------------------------------
        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();

                    numOfRoutes = SupportHelper.SupportedObjects.Count;

                    double[] width = new double[numOfRoutes];
                    double[] height = new double[numOfRoutes];

                    //Get width and height of the Cabletray(s)  
                    for (int index = 0; index < numOfRoutes; index++)
                    {
                        CableTrayObjectInfo cabletrayInfo = (CableTrayObjectInfo)SupportedHelper.SupportedObjectInfo(index + 1);
                        width[index] = (cabletrayInfo.Width * 1000 / 25.4);
                        height[index] = (cabletrayInfo.Depth * 1000 / 25.4);
                    }
                    //configuration definition of the assembly
                    int numberOfParts = 4;
                    IntialParts = numberOfParts;

                    //For Trapeze
                    trapeze_Begin = numberOfParts + 1;
                    numberOfParts = numberOfParts + numOfRoutes;
                    trapeze_End = trapeze_Begin + numOfRoutes - 1;

                    //For Nuts
                    int nut_Begin = trapeze_End + 1;
                    numberOfParts = numberOfParts + numOfRoutes * 4;
                    int nut_End = numberOfParts;

                    //Get the occurrence attributes values
                    PropertyValueCodelist availRodLengthCodelist = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrAssyRodLength", "RodLength");
                    int rodLengthCodelist = (int)availRodLengthCodelist.PropValue;

                    PropertyValueCodelist beamClampCodelist = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrAssyTrapeze", "BeamClamp");
                    long beamClampCodelistValue = (long)beamClampCodelist.PropValue;

                    PropertyValueCodelist MaterialCodelist = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrAssyTrapeze", "Material");
                    int MaterialCodelistValue = (int)MaterialCodelist.PropValue;

                    PropertyValueCodelist typeCodelist = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrAssyTrapeze", "Type");
                    long typeCodelistValue = (long)typeCodelist.PropValue;

                    PropertyValueCodelist finishCodeListValue = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrAssyTrapeze", "Finish");
                    if (finishCodeListValue.PropValue == -1)
                    {
                        int defaultvalue = 5;
                        support.SetPropertyValue(defaultvalue, "IJOAHgrAssyTrapeze", "Finish");
                    }

                    string rodLength = availRodLengthCodelist.PropertyInfo.CodeListInfo.GetCodelistItem((int)rodLengthCodelist).DisplayName;
                    string beamClampType = beamClampCodelist.PropertyInfo.CodeListInfo.GetCodelistItem((int)beamClampCodelistValue).DisplayName;
                    string Material = MaterialCodelist.PropertyInfo.CodeListInfo.GetCodelistItem((int)MaterialCodelistValue).DisplayName;
                    trapezeType = typeCodelist.PropertyInfo.CodeListInfo.GetCodelistItem((int)typeCodelistValue).ShortDisplayName;

                    parts.Add(new PartInfo(BeamClamp1, beamClampType));
                    parts.Add(new PartInfo(BeamClamp2, beamClampType));
                    parts.Add(new PartInfo(LeftHangerRod, "ATR_0.5_" + rodLength));
                    parts.Add(new PartInfo(RightHangerRod, "ATR_0.5_" + rodLength));

                    //Finish exists only when the channel is steel
                    CommonBline_Assembly cmnBline_Assembly = new CommonBline_Assembly();

                    for (int index = 1; index <= numOfRoutes; index++)
                    {
                        Dictionary<String, String> parameters = new Dictionary<string, string>();

                        parameters.Add("TrapezeFinish", support.GetPropertyValue("IJOAHgrAssyTrapeze", "Finish").ToString());
                        parameters.Add("Material", support.GetPropertyValue("IJOAHgrAssyTrapeze", "Material").ToString());
                        parameters.Add("TrapezeType", support.GetPropertyValue("IJOAHgrAssyTrapeze", "Type").ToString());

                        string trapeze = cmnBline_Assembly.GetPartFromTable("BLineAssySteelTrAUX", "IJUAHgrAssySteelTrAUX", "TrapezePartNo", parameters);
                        Trapezepartkeys[index - 1] = "Trapeze" + index.ToString();
                        parts.Add(new PartInfo(Trapezepartkeys[index - 1], trapeze));
                    }
                    //Adding Nuts
                    for (int index = nut_Begin; index <= nut_End; index++)
                    {
                        int index_Nut = index - trapeze_End;
                        Nutpartkeys[index_Nut - 1] = "Nut" + index_Nut.ToString();
                        parts.Add(new PartInfo(Nutpartkeys[index_Nut - 1], "Anvil_HEX_NUT_3"));
                    }

                    return parts;
                }
                catch (Exception e)
                {
                    Type myType = this.GetType();
                    CmnException e1 = new CmnException("Error in Get Assembly Catalog Parts." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                    throw e1;
                }

            }
        }
        //-----------------------------------------------------------------------------------
        //Get Max Route Connection Value
        //-----------------------------------------------------------------------------------
        public override int ConfigurationCount
        {
            get
            {
                return 2;
            }
        }
        //-----------------------------------------------------------------------------------
        //Get Assembly Joints
        //-----------------------------------------------------------------------------------
        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;

                BoundingBoxHelper.CreateStandardBoundingBoxes(false);
                BoundingBox BBX;

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    BBX = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.SupportedAndSupporting);
                }
                else
                {
                    BBX = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.Supported);
                }

                double BBXWidth = BBX.Width;
                double BBXHeight = BBX.Height;
                //====== ======
                //get structure count if more than one Structure
                //====== ======

                Boolean Isturn = false;

                BusinessObject turnRefPort = RefPortHelper.ReferencePort("TurnRef");

                if (turnRefPort != null)
                    Isturn = true;
                else
                    Isturn = false;


                if (SupportHelper.PlacementType != PlacementType.PlaceByReference)
                {
                    int structCount = SupportHelper.SupportingObjects.Count;

                    if (SupportHelper.PlacementType != PlacementType.PlaceByStruct)
                    {
                        if (Isturn == false)
                        {
                            if (structCount < 2)
                            {
                                MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "At least two supporting members to be selected to place this support", "", "Trapeze.cs", 1);
                            }
                        }
                    }
                }

                bool[] isOffsetApplied = new bool[2];
                isOffsetApplied = GetIsLugEndOffsetApplied();

                string[] structPort = new string[2];
                structPort = GetIndexedStructPortName(isOffsetApplied);
                string leftStructPort = structPort[0];
                string rightStructPort = structPort[1];

                int numOfRoutes = SupportHelper.SupportedObjects.Count;

                //Get width and height of the Cabletray(s)  
                for (int index = 1; index <= numOfRoutes; index++)
                {
                    Array.Resize(ref width, numOfRoutes);
                    Array.Resize(ref height, numOfRoutes);

                    CableTrayObjectInfo cabletrayInfo = (CableTrayObjectInfo)SupportedHelper.SupportedObjectInfo(index);
                    width[index - 1] = cabletrayInfo.Width;
                    height[index - 1] = cabletrayInfo.Depth;
                }

                //Sort Width (Get the maximum width)
                for (int i = 0; i < numOfRoutes; i++)
                {
                    if (numOfRoutes > 1)
                    {
                        for (int j = i + 1; j < numOfRoutes; j++)
                        {
                            if (width[j] > width[i])
                            {
                                temp = width[i];
                                width[i] = width[j];
                                width[j] = temp;
                            }
                        }
                    }
                }

                double sortCTWidth = width[0];

                //Get orientation from ports
                for (int index = 1; index <= numOfRoutes; index++)
                {
                    if (index == 1)
                        refPortName = "Route";
                    else
                        refPortName = "Route_" + index;

                    Matrix4X4 portOrientation = RefPortHelper.PortLCS(refPortName);

                    Array.Resize(ref OriginX_Route, numOfRoutes);
                    Array.Resize(ref OriginY_Route, numOfRoutes);
                    Array.Resize(ref OriginZ_Route, numOfRoutes);

                    OriginX_Route[index - 1] = portOrientation.Origin.X;
                    OriginY_Route[index - 1] = portOrientation.Origin.Y;
                    OriginZ_Route[index - 1] = portOrientation.Origin.Z;
                }

                //Get the "Width" property from the beam cross section
                double flangethickness = 0;

                if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                {
                    flangethickness = 0.02;
                }
                else
                {
                    if (SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member || SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.HangerBeam)
                    {
                        flangethickness = SupportingHelper.SupportingObjectInfo(1).FlangeThickness;
                    }
                }

                if (SupportHelper.PlacementType != PlacementType.PlaceByReference)
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct || Isturn == true)
                    {
                        StructFaceNumber = SupportingHelper.SupportingObjectInfo(1).FaceNumber;
                    }
                    else
                    {
                        // Get the current face number
                        if (SupportHelper.SupportingObjects == null)
                        {
                            leftFaceNumber = 1025;
                            rightFaceNumber = 1025;
                            goto CREATE_JOINTS;

                        }
                        if (SupportHelper.SupportingObjects.Count == 0)
                        {
                            leftFaceNumber = 1025;
                            rightFaceNumber = 1025;
                            goto CREATE_JOINTS;
                        }

                        if (SupportHelper.PlacementType == PlacementType.PlaceByPoint)
                        {
                            if (support.SupportingObjects.Count > 1)
                            {
                                leftFaceNumber = SupportingHelper.SupportingObjectInfo(1).FaceNumber;
                                rightFaceNumber = SupportingHelper.SupportingObjectInfo(2).FaceNumber;
                            }
                            else
                            {
                                if (Isturn == false)
                                {
                                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "At least two supporting members to be selected to place this support", "", "Trapeze.cs", 1);
                                }
                            }
                        }
                    }
                else
                {
                        leftFaceNumber = 1025;
                        rightFaceNumber = 1025;
                }
                

            CREATE_JOINTS:
                //====== ======
                // Set Values of Part Occurance Attributes
                //====== ======
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
                SupportComponent BC1 = componentDictionary[BeamClamp1];
                SupportComponent BC2 = componentDictionary[BeamClamp2];
                SupportComponent[] trapezePart = new SupportComponent[numOfRoutes];
                SupportComponent nutPart = componentDictionary[Nutpartkeys[0]];




                BC1.SetPropertyValue(flangethickness, "IJUAHgrBlineBmClampB300", "HO");
                BC2.SetPropertyValue(flangethickness, "IJUAHgrBlineBmClampB300", "HO");

                Plane ConfRouteplaneA, ConfRouteplaneB  ;
                ConfRouteplaneA = ConfRouteplaneB  =  Plane.XY;
                Axis rodBotConfigAxisA, rodBotConfigAxisB, ConfRightStructAxisA, ConfRightStructAxisB, ConfLeftStructAxisA, ConfLeftStructAxisB, ConfRouteAxisA, ConfRouteAxisB;
                rodBotConfigAxisA = rodBotConfigAxisB = ConfRightStructAxisA = ConfRightStructAxisB = ConfLeftStructAxisA = ConfLeftStructAxisB = ConfRouteAxisA = ConfRouteAxisB = Axis.X;


                //Get the distance between Route and the structure ports
                leftHgrRodOffset = RefPortHelper.DistanceBetweenPorts("Route", "BBSR_Low", PortDistanceType.Horizontal);
                rightHgrRodOffset = RefPortHelper.DistanceBetweenPorts("Route", "BBSR_High", PortDistanceType.Horizontal);


                for (int index = 1; index <= numOfRoutes; index++)
                {
                    Array.Resize(ref clampOffset, numOfRoutes);
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct || Isturn == true)
                    {
                        if (index == 1)
                               clampOffset[index - 1] = RefPortHelper.DistanceBetweenPorts("BBSR_Low", "Route", PortDistanceType.Horizontal);
                        else
                            clampOffset[index - 1] = RefPortHelper.DistanceBetweenPorts("BBSR_Low", "Route_" + index, PortDistanceType.Horizontal);
                    }
                    else
                    {
                        if (index == 1)
                            clampOffset[index - 1] = RefPortHelper.DistanceBetweenPorts(leftStructPort, "Route", PortDistanceType.Horizontal);
                        else
                            clampOffset[index - 1] = RefPortHelper.DistanceBetweenPorts(leftStructPort, "Route_" + index, PortDistanceType.Horizontal);
                    }
                }
                Double CCL = 0;
                Double WH = 0;
                Double distanceStruct12Ports = 0;
                Double E = 0;
                E = (double)((PropertyValueDouble)BC2.GetPropertyValue("IJUAHgrBlineBmClampB300", "E")).PropValue;
                BusinessObject[] trapeze = new BusinessObject[numOfRoutes];
                for (int index = 0; index < numOfRoutes; index++)
                {
                    trapezePart[index] = componentDictionary[Trapezepartkeys[index]];
                    trapeze[index] = trapezePart[index].GetRelationship("madeFrom", "part").TargetObjects[0];

                }
                BusinessObject nut = nutPart.GetRelationship("madeFrom", "part").TargetObjects[0];

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct || Isturn == true || SupportHelper.PlacementType == PlacementType.PlaceByReference)
                {
                    if (numOfRoutes == 1)
                        CCL = (double)((PropertyValueDouble)trapeze[0].GetPropertyValue("IJOAHgrBlineTrapChCutL", "TCCL")).PropValue;
                    else
                        //Add largest Tray Width to accommodate Clamps when its postion is Outside.
                        CCL = BBXWidth + sortCTWidth + Inch;
                }
                else
                {
                    distanceStruct12Ports = RefPortHelper.DistanceBetweenPorts(rightStructPort, leftStructPort, PortDistanceType.Horizontal);

                }

                int routeIndex = 0;
                for (int index = trapeze_Begin; index <= trapeze_End; index++)
                {
                    if (index == trapeze_Begin)
                        routeIndex = 1;
                    else
                        routeIndex = routeIndex + 1;
                    double routeStructConfigAngle = GetRouteStructConfigAngle(this, "Route", "Structure", PortAxisType.Y);

                    if (SupportHelper.PlacementType == PlacementType.PlaceByPoint && Isturn == false)
                    {
                        CCL = distanceStruct12Ports +  2*Inch;
                        WH = distanceStruct12Ports - clampOffset[routeIndex - 1] + Inch / 2;
                    }

                    //Set the Channel Cut Length for all Trapezes                      
                    trapezePart[index - trapeze_Begin].SetPropertyValue(CCL, "IJOAHgrBlineTrapChCutL", "TCCL");

                    //Set the WH (distance from Hanger Rod to Cabletray center) dimension
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct || Isturn == true || SupportHelper.PlacementType == PlacementType.PlaceByReference)
                    {
                        if (numOfRoutes == 1)
                            WH = (CCL - Inch) / 2;
                        else
                            WH = CCL - clampOffset[routeIndex - 1] - 3 * Inch / 2 - sortCTWidth / 2;

                    }
                    PropertyValueCodelist insideOutsideCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrAssyInsideOutside", "InsideOutside");
                    int insideOutsideValue = (int)insideOutsideCodeList.PropValue;

                    PropertyValueCodelist clampCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJOAHgrAssyClampGuide", "ClampGuide");
                    int clampValue = (int)clampCodeList.PropValue;

                    trapezePart[index - trapeze_Begin].SetPropertyValue(insideOutsideValue, "IJOAHgrBlineInsideOutside", "Inside_Outside");
                    trapezePart[index - trapeze_Begin].SetPropertyValue(clampValue, "IJOAHgrBlineClampGuide", "Clamp_Guide");
                    trapezePart[index - trapeze_Begin].SetPropertyValue(WH, "IJOAHgrBlineTrDim", "WH");

                }
                //Place by Point
                Double nutThickness;
                Double planeOffset1;
                Double planeOffset2;
                Double connPlaneOffset;

                Double RouteGlobalZ;

                RouteGlobalZ = Math.Round(((RefPortHelper.AngleBetweenPorts("Route", PortAxisType.Z, OrientationAlong.Global_Z))*180)/Math.PI, 0);

                if (SupportHelper.PlacementType != PlacementType.PlaceByStruct)
                {
                    Double byPointAngle1;
                    Double byPointAngle2;
                    Double byPointAngle3;
                 
                    byPointAngle1 = GetRouteStructConfigAngle(this, "Route", "Structure", PortAxisType.Y);
                    byPointAngle2 = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, "Structure", PortAxisType.X, OrientationAlong.Direct);
                    byPointAngle3 = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, "Struct_2", PortAxisType.X, OrientationAlong.Direct);

                    if (Math.Abs(byPointAngle1) < Math.PI / 2)
                    {

                        //ConfRouteIndex = 3436;
                        ConfRouteplaneA = Plane.XY;
                        ConfRouteplaneB = Plane.ZX;
                        ConfRouteAxisA = Axis.Y;
                        ConfRouteAxisB = Axis.NegativeZ;
                    }
                    else
                    {
                        //ConfRouteIndex = 3372;
                        ConfRouteplaneA = Plane.XY;
                        ConfRouteplaneB = Plane.NegativeZX;
                        ConfRouteAxisA = Axis.Y;
                        ConfRouteAxisB = Axis.NegativeZ;
                    }

                }
                //Create Joints 
                double ChThickness = (double)((PropertyValueDouble)trapeze[0].GetPropertyValue("IJUAHgrBlineChannelDim", "CT")).PropValue;
                double ChHeight = (double)((PropertyValueDouble)trapeze[0].GetPropertyValue("IJUAHgrBlineChannelDim", "CH")).PropValue;

                double RefXYorient = Math.Round(((RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, "Structure", PortAxisType.Y, OrientationAlong.Direct)) * 180) / Math.PI, 0);

                Axis CfgRouteAxisA = Axis.X, CfgRouteAxisB = Axis.X;

                if (HgrCompareDoubleService.cmpdbl(RefXYorient, 180) == true)
                {
                    CfgRouteAxisA = Axis.X;
                    CfgRouteAxisB = Axis.X;
                }
                else
                {
                    CfgRouteAxisA = Axis.X;
                    CfgRouteAxisB = Axis.NegativeX;
                }

                int bottom_TrapezeIndex = 0;
                double startZ;
                int botomRouteIndex = 0;
                //To find the bottom Cable Tray
                int routeobjects = SupportHelper.SupportedObjects.Count;
                double minDistance = 10000000;
                for (int index = 1; index <= routeobjects; index++)
                {
                    startZ = SupportedHelper.SupportedObjectInfo(index).StartLocation.Z;
                    if (startZ < minDistance)
                    {
                        minDistance = startZ;
                        botomRouteIndex = index;
                    }
                }

                bottom_TrapezeIndex = trapeze_Begin + botomRouteIndex - 1;
                int nut_Index = 0;
                int nut_Begin = 0;
                int trapeze_Index = 0;
                for (routeIndex = 1; routeIndex <= numOfRoutes; routeIndex++)
                {
                    if (routeIndex == 1)
                    {
                        refPortName = "Route";
                        trapeze_Begin = IntialParts + 1;
                        trapeze_Index = trapeze_Begin;
                        nut_Index = trapeze_End + 1;
                        nut_Begin = nut_Index;
                    }
                    else
                    {
                        refPortName = "Route_" + routeIndex;
                        trapeze_Index = trapeze_Index + 1;
                        nut_Index = nut_Index + 1;
                    }

                    connPlaneOffset = OriginZ_Route[0] - OriginZ_Route[routeIndex - 1];

                    //Add joints between Trapeze and Route
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct || SupportHelper.PlacementType == PlacementType.PlaceByReference)
                    {
                        if (routeIndex == 1)
                            JointHelper.CreateRigidJoint(Trapezepartkeys[trapeze_Index - trapeze_Begin], "TrayPort", "-1", refPortName, Plane.XY, Plane.XY, CfgRouteAxisA, CfgRouteAxisB, 1.2 * (height[trapeze_Index - IntialParts - 1] / 2), 0, 0);
                        else
                            JointHelper.CreateRigidJoint(Trapezepartkeys[trapeze_Begin - trapeze_Begin], "InThrdRH1", Trapezepartkeys[trapeze_Index - trapeze_Begin], "InThrdRH1", Plane.XY, Plane.XY, Axis.X, Axis.X, -connPlaneOffset, 0, 0);
                    }
                    else
                    {
                        if (routeIndex == 1)
                        {
                            if (Isturn == true)
                            {
                                if (HgrCompareDoubleService.cmpdbl(RouteGlobalZ, 90) == true)
                                {
                                    JointHelper.CreateRigidJoint(Trapezepartkeys[trapeze_Index - trapeze_Begin], "TrayPort", "-1", refPortName, ConfRouteplaneA, ConfRouteplaneB, ConfRouteAxisA, ConfRouteAxisB, 1.2 * (height[trapeze_Index - IntialParts - 1] / 2), 0, 0);
                                }
                                else
                                {
                                    JointHelper.CreateRigidJoint(Trapezepartkeys[trapeze_Index - trapeze_Begin], "TrayPort", "-1", refPortName, Plane.XY, Plane.XY, Axis.X, Axis.X, 1.2 * (height[trapeze_Index - IntialParts - 1] / 2), 0, 0);
                                }
                            }
                            else
                            {
                                JointHelper.CreateRigidJoint(Trapezepartkeys[trapeze_Index - trapeze_Begin], "TrayPort", "-1", refPortName, ConfRouteplaneA, ConfRouteplaneB, ConfRouteAxisA, ConfRouteAxisB, 1.2 * (height[trapeze_Index - IntialParts - 1] / 2), 0, 0);
                            }
                        }
                        else
                        {
                            JointHelper.CreateRigidJoint(Trapezepartkeys[trapeze_Begin - trapeze_Begin], "InThrdRH1", Trapezepartkeys[trapeze_Index - trapeze_Begin], "InThrdRH1", Plane.XY, Plane.XY, Axis.X, Axis.X, -connPlaneOffset, 0, 0);
                        }
                    }


                    nutThickness = (double)((PropertyValueDouble)nut.GetPropertyValue("IJUAHgrAnvil_hex_nut", "T")).PropValue;

                    if (trapezeType.ToLower() == "standard")
                        planeOffset1 = nutThickness - ChThickness;
                    else
                        planeOffset1 = -ChHeight + nutThickness - ChThickness;

                    //Add Joints between the bottom nuts and the rods
                    JointHelper.CreateRigidJoint(Trapezepartkeys[trapeze_Index - trapeze_Begin], "InThrdRH1", Nutpartkeys[nut_Index - nut_Begin], "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.NegativeX, planeOffset1, 0, 0);

                    nut_Index = nut_Index + 1;
                    JointHelper.CreateRigidJoint(Trapezepartkeys[trapeze_Index - trapeze_Begin], "InThrdRH2", Nutpartkeys[nut_Index - nut_Begin], "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.NegativeX, planeOffset1, 0, 0);


                    planeOffset2 = ChHeight + 2 * nutThickness - ChThickness;


                    nut_Index = nut_Index + 1;
                    //Add Joints between the Top nuts and the rods

                    JointHelper.CreateRigidJoint(Trapezepartkeys[trapeze_Index - trapeze_Begin], "InThrdRH1", Nutpartkeys[nut_Index - nut_Begin], "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.NegativeX, planeOffset2, 0, 0);

                    nut_Index = nut_Index + 1;
                    JointHelper.CreateRigidJoint(Trapezepartkeys[trapeze_Index - trapeze_Begin], "InThrdRH2", Nutpartkeys[nut_Index - nut_Begin], "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.NegativeX, planeOffset2, 0, 0);
                }

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    // rodBotConfigIndex = 82;
                    rodBotConfigAxisA = Axis.Y;
                    rodBotConfigAxisB = Axis.Y;
                }
                else
                {
                    //rodBotConfigIndex = 91;
                    rodBotConfigAxisA = Axis.Z;
                    rodBotConfigAxisB = Axis.Z;
                }

                //Add joints between bottom most Trapeze and Left rod  
                JointHelper.CreateRigidJoint(Trapezepartkeys[bottom_TrapezeIndex - trapeze_Begin], "InThrdRH1", LeftHangerRod, "BotExThdRH",Plane.XY, Plane.NegativeXY, Axis.X, Axis.X , 0, 0, 0);


                //Add Joints between bottom most Trapeze and Right rod
                JointHelper.CreateRigidJoint(Trapezepartkeys[bottom_TrapezeIndex - trapeze_Begin], "InThrdRH2", RightHangerRod, "BotExThdRH", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, 0, 0);

  
                
                ////Add joints between Rods and Beam Clamp
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    if (StructFaceNumber == 1025 || StructFaceNumber == 1026)
                    {
                        JointHelper.CreateRigidJoint(LeftHangerRod, "TopExThdRH", BeamClamp1, "InThrdRH1", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, 0);
                        JointHelper.CreateRigidJoint(RightHangerRod, "TopExThdRH", BeamClamp2, "InThrdRH1", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, 0);
                    }
                    else
                    {
                        JointHelper.CreateRigidJoint(LeftHangerRod, "TopExThdRH", BeamClamp1, "InThrdRH1", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, 0, 0);
                        JointHelper.CreateRigidJoint(RightHangerRod, "TopExThdRH", BeamClamp2, "InThrdRH1", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, 0, 0);
                    }
                }
                else
                {
                    JointHelper.CreateRigidJoint(LeftHangerRod, "TopExThdRH", BeamClamp1, "InThrdRH1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    JointHelper.CreateRigidJoint(RightHangerRod, "TopExThdRH", BeamClamp2, "InThrdRH1", Plane.XY, Plane.XY, Axis.X, Axis.NegativeX, 0, 0, 0);
                }

                //Add Joints between top and bottom ports of the left hanger rod
                JointHelper.CreatePrismaticJoint(LeftHangerRod, "BotExThdRH", LeftHangerRod, "TopExThdRH",Plane.ZX,Plane.NegativeZX, Axis.Z, Axis.NegativeZ,0,0);

                //Add Joints between top and bottom ports of the right hanger rod
                JointHelper.CreatePrismaticJoint(RightHangerRod, "BotExThdRH", RightHangerRod, "TopExThdRH", Plane.ZX, Plane.NegativeZX, Axis.Z, Axis.NegativeZ, 0, 0);

                Plane PlanarCfgIdx = Plane.XY;

                if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                {
                    PlanarCfgIdx = Plane.XY;
                }
                else
                {
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    {
                        if (Isturn == true)
                        {
                            PlanarCfgIdx = Plane.ZX;
                        }
                        else
                        {
                            PlanarCfgIdx = Plane.XY;
                        }
                    }
                    else
                    {
                        PlanarCfgIdx = Plane.ZX;
                    }
                }

                //Add joints between C Clamp and the Structure
                JointHelper.CreatePointOnPlaneJoint(BeamClamp1, "Steel", "-1", "Structure", PlanarCfgIdx);
                JointHelper.CreatePointOnPlaneJoint(BeamClamp2, "Steel", "-1", "Structure", PlanarCfgIdx);

            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in Get Assembly Joints." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }

        }
        //-----------------------------------------------------------------------------------
        //Get Route Connections
        //-----------------------------------------------------------------------------------
        public override Collection<ConnectionInfo> SupportedConnections
        {
            get
            {
                Collection<ConnectionInfo> routeConns = new Collection<ConnectionInfo>();
                int NumOfRoutes = SupportHelper.SupportedObjects.Count;

                for (int index = 1; index <= NumOfRoutes; index++)
                {
                    partkey = Trapezepartkeys[index - 1];
                    int connecttoroute = index;
                    routeConns.Add(new ConnectionInfo(partkey, connecttoroute));
                }

                return routeConns;
            }
        }
        //-----------------------------------------------------------------------------------
        //Get Struct Connections
        //-----------------------------------------------------------------------------------
        public override Collection<ConnectionInfo> SupportingConnections
        {
            get
            {
                Collection<ConnectionInfo> structConns = new Collection<ConnectionInfo>();
                structConns.Add(new ConnectionInfo(BeamClamp1, 1));
                return structConns;
            }
        }
        //-----------------------------------------------------------------------------------
        //BOM Description
        //-----------------------------------------------------------------------------------
        #region ICustomHgrBOMDescription Members
        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string sBOMString = "";
            try
            {
                string hardware;
                string clampType;
                long insideOuside;
                long clampGuide;
                long beamClamp;

                PropertyValueCodelist beamClampCodelist;
                beamClampCodelist = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrAssyTrapeze", "BeamClamp");
                beamClamp = beamClampCodelist.PropValue;

                PropertyValueCodelist insideOutsideCodelist;
                insideOutsideCodelist = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrAssyInsideOutside", "InsideOutside");
                insideOuside = insideOutsideCodelist.PropValue;
                hardware = oSupportOrComponent.GetPropertyValue("IJUAHgrAssyHardware", "Hardware").ToString();

                PropertyValueCodelist clampGuideCodelist;
                clampGuideCodelist = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrAssyClampGuide", "ClampGuide");
                clampGuide = clampGuideCodelist.PropValue;


                if (insideOuside == 0)
                    insideOuside = 1;
                string insideOusideLongStringValue = insideOutsideCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(insideOutsideCodelist.PropValue).DisplayName;

                if (beamClamp == 0)
                    beamClamp = 1;
                string beamClampLongStringValue = beamClampCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(beamClampCodelist.PropValue).DisplayName;

                if (clampGuide == 0)
                    clampGuide = 1;
                string clampGuideLongStringValue = clampGuideCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(clampGuideCodelist.PropValue).DisplayName;
                clampType = oSupportOrComponent.GetPropertyValue("IJUAHgrAssyClampType", "ClampType").ToString();

                sBOMString = "B-Line Assy Trapeze Support with CT " + clampGuideLongStringValue + "type: " + clampType + " mounted, " + insideOusideLongStringValue + " and with Beam Clamp: " + beamClampLongStringValue;
                return sBOMString;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in BOM of Assembly - Trapeze BOM" + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        #endregion
        /// <summary>
        /// This method returns the direct angle between Route and structur ports in Radians
        /// </summary>
        /// <param name="routePortName">string-route Port Name.</param>
        /// <param name="structPortName">string-structure Port Name.</param>
        /// <param name="axisType">PortAxisType-axis Type.</param>
        /// <returns>double</returns>        
        /// <code>
        ///   byPointAngle1=GetRouteStructConfigAngle("Route", "Structure", PortAxisType.Y);
        ///</code>
        public double GetRouteStructConfigAngle(CustomSupportDefinition customSupportDefinition, String routePortName, String structPortName, PortAxisType axisType)
        {
            try
            {

                //get the appropriate axis
                Vector[] vecAxis = new Vector[2];

                Matrix4X4 routeMatrix = customSupportDefinition.RefPortHelper.PortLCS(routePortName);
                Position routepoint = routeMatrix.Origin;

                switch (axisType)
                {
                    case PortAxisType.X:
                        {
                            vecAxis[0] = routeMatrix.XAxis;
                            break;
                        }
                    case PortAxisType.Y:
                        {
                            vecAxis[0] = routeMatrix.ZAxis.Cross(routeMatrix.XAxis);
                            break;
                        }
                    case PortAxisType.Z:
                        {
                            vecAxis[0] = routeMatrix.ZAxis;
                            break;
                        }
                }
                Matrix4X4 structMatrix = customSupportDefinition.RefPortHelper.PortLCS(structPortName);
                Position structPoint = structMatrix.Origin;
                vecAxis[1] = structPoint.Subtract(routepoint);

                return GetAngleBetweenVectors(vecAxis[0], vecAxis[1]);

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        /// <summary>
        /// This method returns the direct angle between two vectors in Radians
        /// </summary>
        /// <param name="vector1">The vector1(Vector).</param>
        /// <param name="vector2">The vector2(Vector).</param>
        /// <returns>double</returns>        
        /// <code>
        ///       ContentHelper contentHelper = new ContentHelper();
        ///       double value;
        ///       value = contentHelper. GetAngleBetweenVectors(vector1, vector2 );
        ///</code>

        public double GetAngleBetweenVectors(Vector vector1, Vector vector2)
        {
            try
            {
                double dblDotProd = (vector1.Dot(vector2) / (vector1.Length * vector2.Length));
                double dblArcCos = 0.0;
                if (HgrCompareDoubleService.cmpdbl(Math.Abs(dblDotProd), 1) == false)
                {
                    dblArcCos = Math.PI / 2 - Math.Atan(dblDotProd / Math.Sqrt(1 - dblDotProd * dblDotProd));
                }
                else if (HgrCompareDoubleService.cmpdbl(dblDotProd, -1) == true)
                {
                    dblArcCos = Math.PI;
                }
                else if (HgrCompareDoubleService.cmpdbl(dblDotProd, 1) == true)
                {
                    dblArcCos = 0;
                }
                return dblArcCos;
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }
        /// <summary>
        /// This method differentiate multiple structure input object based on relative position.
        /// </summary>
        /// <param name="IsOffsetApplied">IsOffsetApplied-The offset applied.(Boolean[])</param>
        /// <returns>String array</returns>
        /// <code>
        ///     string[] idxStructPort = GetIndexedStructPortName(bIsOffsetApplied);
        /// </code>
        public String[] GetIndexedStructPortName(Boolean[] IsOffsetApplied)
        {
            String[] structurePort = new String[2];
            int structureCount = SupportHelper.SupportingObjects.Count;
            int i;

            if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
            {
                structurePort[0] = "Structure";
                structurePort[1] = "Structure";
            }
            else
            {
                structurePort[0] = "Structure";
                structurePort[1] = "Struct_2";

                if (structureCount > 1)
                {
                    if (SupportHelper.PlacementType != PlacementType.PlaceByStruct)
                    {
                        for (i = 0; i <= 1; i++)
                        {
                            double angle = 0;
                            if ((SupportHelper.SupportingObjects.Count != 0))
                            {
                                if ((SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member || SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.HangerBeam) && IsOffsetApplied[i] == false)
                                {
                                    angle = RefPortHelper.PortConfigurationAngle("Route", structurePort[i], PortAxisType.Y);
                                }
                            }
                            //the port is the right structure port
                            if (Math.Abs(angle) < Math.PI / 2.0)
                            {
                                if (i == 0)
                                {
                                    structurePort[0] = "Struct_2";
                                    structurePort[1] = "Structure";
                                }
                            }
                            //the port is the left structure port
                            else
                            {
                                if (i == 1)
                                {
                                    structurePort[0] = "Struct_2";
                                    structurePort[1] = "Structure";
                                }
                            }
                        }
                    }
                }
                else
                    structurePort[1] = "Structure";
            }
            //switch the OffsetApplied flag
            if (structurePort[0] == "Struct_2")
            {
                Boolean flag = IsOffsetApplied[0];
                IsOffsetApplied[0] = IsOffsetApplied[1];
                IsOffsetApplied[1] = flag;
            }

            return structurePort;
        }
        /// <summary>
        /// This method checks whether a offset value needs to be explicitly specified on both end of the leg.
        /// </summary>
        /// <returns>Boolean array</returns>
        /// <code>
        ///     Boolean[] bIsOffsetApplied = GetIsLugEndOffsetApplied();  
        ///</code>
        public Boolean[] GetIsLugEndOffsetApplied()
        {
            try
            {
                Collection<BusinessObject> StructureObjects;
                Boolean[] isOffsetApplied = new Boolean[2];

                //first route object is set as primary route object
                StructureObjects = SupportHelper.SupportingObjects;

                isOffsetApplied[0] = true;
                isOffsetApplied[1] = true;

                if (StructureObjects != null)
                {
                    if (StructureObjects.Count >= 2)
                    {
                        for (int index = 0; index <= 1; index++)
                        {
                            //check if offset is to be applied
                            const double RouteStructAngle = Math.PI / 180.0;
                            Boolean varRuleApplied = true;
                            if ((SupportHelper.SupportingObjects.Count != 0))
                            {
                                if (SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member)
                                {
                                    //if angle is within 1 degree, regard as parallel case
                                    //Also check for Sloped structure                                
                                    MemberPart memberPart = (MemberPart)SupportHelper.SupportingObjects[index];
                                    ICurve memberCurve = memberPart.Axis;

                                    Vector supportedVector = new Vector();
                                    Vector supportingVector = new Vector();

                                    if (SupportedHelper.SupportedObjectInfo(1).SupportedObjectType == SupportedObjectType.Pipe)
                                    {
                                        Position startLocation = new Position(SupportedHelper.SupportedObjectInfo(1).StartLocation);
                                        Position endLocation = new Position(SupportedHelper.SupportedObjectInfo(1).EndLocation);
                                        supportedVector = new Vector(endLocation - startLocation);
                                    }
                                    if (memberCurve is ILine)
                                    {
                                        ILine line = (ILine)memberCurve;
                                        supportingVector = line.Direction;
                                    }

                                    double angle = GetAngleBetweenVectors(supportingVector, supportedVector);
                                    double refAngle1 = RefPortHelper.AngleBetweenPorts("Struct_2", PortAxisType.Y, OrientationAlong.Global_X) - Math.PI / 2;
                                    double refAngle2 = RefPortHelper.AngleBetweenPorts("Struct_2", PortAxisType.X, OrientationAlong.Global_X);

                                    if (angle < (refAngle1 + 0.001) & angle > (refAngle1 - 0.001))
                                        angle = angle - Math.Abs(refAngle1);
                                    else if (angle < (refAngle2 + 0.001) & angle > (refAngle2 - 0.001))
                                        angle = angle - Math.Abs(refAngle2);
                                    else
                                        angle = angle + Math.Abs(refAngle1) + Math.Abs(refAngle2);
                                    angle = angle + Math.Abs(refAngle1) + Math.Abs(refAngle2);
                                    if (Math.Abs(angle) < RouteStructAngle || Math.Abs(angle - Math.PI) < RouteStructAngle)
                                        varRuleApplied = false;
                                }
                            }
                            isOffsetApplied[index] = varRuleApplied;
                        }
                    }
                }

                return isOffsetApplied;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in GetIsLugEndOffsetApplied Method of Bline_Assy." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }


    }
}




