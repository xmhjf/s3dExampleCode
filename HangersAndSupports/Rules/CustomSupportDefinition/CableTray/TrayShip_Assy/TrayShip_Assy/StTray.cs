//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   StTray.cs
//   TrayShip_Assy,Ingr.SP3D.Content.Support.Rules.StTray
//   Author       :  Vijay
//   Creation Date:  12/07/2013   
//   Description:    

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   12/07/2013     Vijay    CR-CP-224487  Convert HS_TrayShip_Assy to C# .Net  
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;
using Ingr.SP3D.Support.Middle;

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
    public class StTray : CustomSupportDefinition
    {
        private int shape, connection, plateType, noOfLegs, legBegin, legEnd, partCount, padBegin, padEnd, numClamps, clampBegin, numBolts, boltsBegin, trayBegin, trayEnd;
        private string sectionSize, uBoltPart;
        private bool includePad, overhangRule;
        private double length, plateHeight, leftOffset, rightOffset, gap, underLength, width, height, padDiameter, padThickness;
        static int numOfRoutes = 0, boltsEnd = 0, clampEnd = 0;
        string[] part, clamp, bolt;

        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();

                    //Get the attributes from assembly
                    shape = (int)((PropertyValueCodelist)support.GetPropertyValue("IJUAhsTrayAsmTrayShape", "Shape")).PropValue;
                    connection = (int)((PropertyValueCodelist)support.GetPropertyValue("IJUAhsTrayAsmConn", "Connection")).PropValue;
                    PropertyValueCodelist sectionSizeCodelist = (PropertyValueCodelist)support.GetPropertyValue("IJUAhsTrayAsmLSecSize", "SectionSize");
                    sectionSize = sectionSizeCodelist.PropertyInfo.CodeListInfo.DisplayName;
                    includePad = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJUAhsTrayAsmPad", "IncludePad")).PropValue;
                    length = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsTrayAsmTrayL", "L")).PropValue;
                    plateHeight = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsTrayAsmTrayH", "H")).PropValue;
                    leftOffset = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsTrayAsmLeftOH", "L1")).PropValue;
                    rightOffset = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsTrayAsmRightOH", "L2")).PropValue;
                    gap = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsTrayAsmGap", "Gap")).PropValue;
                    underLength = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsTrayAsmUnderL", "UnderLength")).PropValue;
                    overhangRule = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJUAhsTrayAsmOverhang", "Overhang")).PropValue;
                    plateType = (int)((PropertyValueCodelist)support.GetPropertyValue("IJUAhsTrayAsmTrayType", "PlateType")).PropValue;

                    CableTrayObjectInfo cableTrayInfo = (CableTrayObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                    numOfRoutes = SupportHelper.SupportedObjects.Count;
                    width = cableTrayInfo.Width;
                    height = cableTrayInfo.Depth;

                    //Create the list of part classes required by the type
                    trayBegin = 1;
                    trayEnd = numOfRoutes;
                    part = new string[numOfRoutes + 1];
                    for (int index = 1; index <= trayEnd; index++)
                    {
                        part[index] = "idxPart" + index;
                        if (shape == 5)    //Rectangle
                            parts.Add(new PartInfo(part[index], "StraightTray_1"));  //"TrayShip_RectStPlate_1"
                        else if (shape == 1)  //Round
                            parts.Add(new PartInfo(part[index], "StraightTrayCirRail_1"));   //"TrayShip_RndStPlate_1"
                    }

                    double positionLength = Math.Abs(length);
                    CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                    IEnumerable<BusinessObject> trayshipPartclass = null;
                    PartClass trayShipServiceStNoOfLegsPartClass = (PartClass)catalogBaseHelper.GetPartClass("TrayShipSrv_StNoOfLegs");

                    if (trayShipServiceStNoOfLegsPartClass.PartClassType.Equals("HgrServiceClass"))
                        trayshipPartclass = trayShipServiceStNoOfLegsPartClass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                    else
                        trayshipPartclass = trayShipServiceStNoOfLegsPartClass.Parts;

                    trayshipPartclass = trayshipPartclass.Where(part1 => (string)((PropertyValueString)part1.GetPropertyValue("IJUAhsTraySrvStTray", "UnitType")).PropValue == "mm" && (((double)((PropertyValueDouble)part1.GetPropertyValue("IJUAhsTraySrvStTray", "LengthMin")).PropValue < (length * 1000)) && ((double)((PropertyValueDouble)part1.GetPropertyValue("IJUAhsTraySrvStTray", "LengthMax")).PropValue > (length * 1000))));
                    if (trayshipPartclass.Count() > 0)
                    {
                        noOfLegs = Convert.ToInt32((string)((PropertyValueString)trayshipPartclass.ElementAt(0).GetPropertyValue("IJUAhsTraySrvStTray", "NoOfLegs")).PropValue);
                        if (overhangRule == true)
                        {
                            leftOffset = (double)((PropertyValueDouble)trayshipPartclass.ElementAt(0).GetPropertyValue("IJUAhsTraySrvStTray", "LeftOverhang")).PropValue / 1000;
                            rightOffset = (double)((PropertyValueDouble)trayshipPartclass.ElementAt(0).GetPropertyValue("IJUAhsTraySrvStTray", "RightOverhang")).PropValue / 1000;
                        }
                    }

                    //For St Plate
                    legBegin = numOfRoutes + 1;
                    legEnd = legBegin + noOfLegs - 1;

                    PartClass trayShipServicePadDimPartClass = (PartClass)catalogBaseHelper.GetPartClass("TrayShipSrv_PadDim");

                    if (trayShipServicePadDimPartClass.PartClassType.Equals("HgrServiceClass"))
                        trayshipPartclass = trayShipServicePadDimPartClass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                    else
                        trayshipPartclass = trayShipServicePadDimPartClass.Parts;

                    trayshipPartclass = trayshipPartclass.Where(part12s => (int)((PropertyValueCodelist)part12s.GetPropertyValue("IJUAhsTraySrvPadDim", "SectionSize")).PropValue == (int)(sectionSizeCodelist.PropValue));
                    if (trayshipPartclass.Count() > 0)
                    {
                        padDiameter = (double)((PropertyValueDouble)trayshipPartclass.ElementAt(0).GetPropertyValue("IJUAhsTraySrvPadDim", "PadDia")).PropValue / 1000;
                        padThickness = (double)((PropertyValueDouble)trayshipPartclass.ElementAt(0).GetPropertyValue("IJUAhsTraySrvPadDim", "PadThick")).PropValue / 1000;
                    }


                    PartClass traySrvUBoltSelPartClass = (PartClass)catalogBaseHelper.GetPartClass("TrayShipSrv_UBoltSel");

                    if (traySrvUBoltSelPartClass.PartClassType.Equals("HgrServiceClass"))
                        trayshipPartclass = traySrvUBoltSelPartClass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                    else
                        trayshipPartclass = traySrvUBoltSelPartClass.Parts;

                    foreach (BusinessObject Bo in trayshipPartclass)
                    {
                        double minValue = (double)((PropertyValueDouble)Bo.GetPropertyValue("IJUAhsTraySrvUBoltSel", "PlHeightMin")).PropValue;
                        double maxValue = (double)((PropertyValueDouble)Bo.GetPropertyValue("IJUAhsTraySrvUBoltSel", "PlHeightMax")).PropValue;
                        string units = (string)((PropertyValueString)Bo.GetPropertyValue("IJUAhsTraySrvUBoltSel", "UnitType")).PropValue;
                        double pv1 = plateHeight * 1000;
                        bool value = (minValue <= pv1) && (maxValue >= pv1);
                        if (units == "mm")
                        {
                            if (value)
                            {
                                uBoltPart = (string)((PropertyValueString)Bo.GetPropertyValue("IJUAhsTraySrvUBoltSel", "UBoltPart")).PropValue;
                                break;
                            }
                        }
                    }
                    for (int index = legBegin; index <= legEnd; index++)
                    {
                        Array.Resize(ref part, numOfRoutes + 1 + index + 1);
                        part[index] = "idxPart" + index;
                        parts.Add(new PartInfo(part[index], sectionSizeCodelist + " " + "Japan-2005"));
                    }

                    if (includePad)
                    {
                        partCount = parts.Count;
                        padBegin = partCount + 1;
                        padEnd = padBegin + noOfLegs - 1;
                        for (int index = padBegin; index <= padEnd; index++)
                        {
                            Array.Resize(ref part, numOfRoutes + 1 + index + 1);
                            part[index] = "idxPart" + index;
                            parts.Add(new PartInfo(part[index], "Util_Plate_Metric_1"));
                        }
                    }

                    if (connection == 1)       //Clamped
                    {
                        partCount = parts.Count;
                        numClamps = noOfLegs * numOfRoutes;
                        clampBegin = (partCount + 1);
                        clampEnd = clampBegin + numClamps - 1;
                        clamp = new string[clampEnd + 1];
                        for (int index = clampBegin; index <= clampEnd; index++)
                        {
                            clamp[index] = "iClamp" + index;
                            parts.Add(new PartInfo(clamp[index], uBoltPart));
                        }
                    }
                    else if (connection == 2)      //Bolts
                    {
                        partCount = parts.Count;
                        numBolts = noOfLegs * numOfRoutes;
                        boltsBegin = (partCount + 1);
                        boltsEnd = boltsBegin + numBolts - 1;
                        bolt = new string[boltsEnd + 1];
                        for (int index = boltsBegin; index <= boltsEnd; index++)
                        {
                            bolt[index] = "iBolt" + index;
                            parts.Add(new PartInfo(bolt[index], "Util_Fixed_Cyl_Metric_1"));
                        }
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
        //Get Assembly Catalog Parts
        //-----------------------------------------------------------------------------------

        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;

                BusinessObject horizontalSectionPart = componentDictionary[part[legBegin]].GetRelationship("madeFrom", "part").TargetObjects[0];
                CrossSection crossSection = (CrossSection)horizontalSectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];

                double sectionWidth = crossSection.Width;
                double sectionDepth = crossSection.Depth;
                double sectionThick = (double)((PropertyValueDouble)crossSection.GetPropertyValue("IStructFlangedSectionDimensions", "tf")).PropValue;
                string hgrPort = string.Empty;
                Matrix4X4 hangerPort = new Matrix4X4();
                double[] hangerPortOriginZ_Route = new double[SupportHelper.SupportedObjects.Count + 1];
                for (int idxRoute = 1; idxRoute <= numOfRoutes; idxRoute++)
                {
                    if (idxRoute == 1)
                        hgrPort = "Route";
                    else
                        hgrPort = "Route_" + idxRoute;

                    hangerPort = RefPortHelper.PortLCS(hgrPort);

                    hangerPortOriginZ_Route[idxRoute] = hangerPort.Origin.Z;
                }

                //=======================================
                //Do Something if more than one Structure
                //=======================================
                //get structure count

                Boolean[] isOffsetApplied = TrayShipAseemblyServices.GetIsLugEndOffsetApplied(this);
                string[] structPort = TrayShipAseemblyServices.GetIndexedStructPortName(this, isOffsetApplied);
                string leftStructPort = structPort[0];
                string rightStructPort = structPort[1];

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    leftStructPort = "Structure";
                    rightStructPort = "Structure";
                }

                double stPlateHeight;
                if (shape == 5)   //Rectangle
                    stPlateHeight = height;
                else if (shape == 1)  //Round
                    stPlateHeight = height / 2;

                double boltRad = sectionThick * 2;
                double boltLen = sectionThick * 6;
                double rodDiameter = plateHeight;
                double clampOD = rodDiameter;

                for (int index = trayBegin; index <= trayEnd; index++)
                {
                    if (shape == 1)   //Round Plate
                    {
                        componentDictionary[part[index]].SetPropertyValue(plateHeight / 2, "IJOAHgrTrayThickness", "Thickness");
                        componentDictionary[part[index]].SetPropertyValue(rodDiameter, "IJOAHgrTrayUBoltDia", "D");
                        componentDictionary[part[index]].SetPropertyValue(width - 2 * rodDiameter, "IJOAHgrTrayWidth", "Width");
                        componentDictionary[part[index]].SetPropertyValue(plateType, "IJOAHgrTrayPlTyp", "PlateType");
                    }
                    else
                    {
                        componentDictionary[part[index]].SetPropertyValue(width, "IJOAHgrTrayWidth", "Width");
                        componentDictionary[part[index]].SetPropertyValue(plateHeight, "IJOAHgrTrayThickness", "Thickness");
                    }
                    componentDictionary[part[index]].SetPropertyValue(length, "IJOAHgrTrayLength", "Length");
                }

                if (connection == 1)      //Clamped
                {
                    for (int index = clampBegin; index <= clampEnd; index++)
                    {
                        componentDictionary[clamp[index]].SetPropertyValue(sectionThick, "IJOAHgrGenUBoltFlgThick", "FlgThick");
                        componentDictionary[clamp[index]].SetPropertyValue(clampOD / 2, "IJOAHgrGenericUBolt", "PipeRadius");
                    }
                }
                else if (connection == 2)      //Bolted
                {
                    for (int index = boltsBegin; index <= boltsEnd; index++)
                    {
                        componentDictionary[bolt[index]].SetPropertyValue(boltLen, "IJOAHgrUtilMetricL", "L");
                        componentDictionary[bolt[index]].SetPropertyValue(boltRad, "IJOAHgrUtilMetricRadius", "Radius");
                    }
                }

                for (int index = legBegin; index <= legEnd; index++)
                {
                    PropertyValueCodelist braceCodelist = (PropertyValueCodelist)componentDictionary[part[index]].GetPropertyValue("IJOAhsSteelCP", "CP1");
                    braceCodelist.PropValue = 1;
                    componentDictionary[part[index]].SetPropertyValue(braceCodelist.PropValue, "IJOAhsSteelCP", "CP1");
                    componentDictionary[part[index]].SetPropertyValue(0.0, "IJOAhsBeginCap", "BeginCapRotZ");
                    componentDictionary[part[index]].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                    componentDictionary[part[index]].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                    if (!includePad)
                        padThickness = 0;
                    if (noOfLegs == 2)
                    {
                        if (index == legBegin)
                        {
                            componentDictionary[part[index]].SetPropertyValue(underLength, "IJUAHgrOccOverLength", "BeginOverLength");
                            componentDictionary[part[index]].SetPropertyValue(-padThickness, "IJUAHgrOccOverLength", "EndOverLength");
                        }
                        else
                        {
                            componentDictionary[part[index]].SetPropertyValue(-padThickness, "IJUAHgrOccOverLength", "BeginOverLength");
                            componentDictionary[part[index]].SetPropertyValue(underLength, "IJUAHgrOccOverLength", "EndOverLength");
                        }
                    }
                    else
                    {
                        if (index == legBegin || index == legBegin + 3 || index == legBegin + 5)
                        {
                            componentDictionary[part[index]].SetPropertyValue(-padThickness, "IJUAHgrOccOverLength", "BeginOverLength");
                            componentDictionary[part[index]].SetPropertyValue(underLength, "IJUAHgrOccOverLength", "EndOverLength");
                        }
                        else
                        {
                            componentDictionary[part[index]].SetPropertyValue(underLength, "IJUAHgrOccOverLength", "BeginOverLength");
                            componentDictionary[part[index]].SetPropertyValue(-padThickness, "IJUAHgrOccOverLength", "EndOverLength");
                        }
                    }
                }

                if (includePad)
                {
                    for (int index = padBegin; index <= padEnd; index++)
                    {
                        componentDictionary[part[index]].SetPropertyValue(padDiameter, "IJOAHgrUtilMetricWidth", "Width");
                        componentDictionary[part[index]].SetPropertyValue(padDiameter, "IJOAHgrUtilMetricDepth", "Depth");
                        componentDictionary[part[index]].SetPropertyValue(padThickness, "IJOAHgrUtilMetricT", "T");
                    }
                }

                //Set attributes on assembly object
                support.SetPropertyValue(connection, "IJUAhsTrayAsmConn", "Connection");
                support.SetPropertyValue(shape, "IJUAhsTrayAsmTrayShape", "Shape");
                support.SetPropertyValue(noOfLegs.ToString(), "IJUAhsTrayAsmNoOfLegs", "NoOfLegs");
                support.SetPropertyValue(leftOffset, "IJUAhsTrayAsmLeftOH", "L1");
                support.SetPropertyValue(rightOffset, "IJUAhsTrayAsmRightOH", "L2");

                double dRouteStPltPlaneOff = 0, dRouteStPltAxisOff = 0, dLegStPltPlaneOff = 0, dLegStOriginOffset = 0, dBoltPlaneOffset = 0;
                string strLegPort1 = string.Empty;
                string strLegPort2 = string.Empty;

                if (shape == 5)    //Rectangle
                {
                    dRouteStPltPlaneOff = plateHeight + height / 2;
                    dLegStPltPlaneOff = 0;
                    dLegStOriginOffset = width / 2;
                }
                else if (shape == 1)
                {
                    dLegStPltPlaneOff = -plateHeight / 4;
                    dLegStOriginOffset = width / 2;
                    dRouteStPltPlaneOff = plateHeight * 3 / 4 + height / 2;
                }

                dBoltPlaneOffset = plateHeight / 2;

                strLegPort1 = "BeginCap";
                strLegPort2 = "EndCap";

                if (Configuration == 1)
                    dRouteStPltAxisOff = length / 2;
                else
                    dRouteStPltAxisOff = -length / 2;

                int[] indexRoute = new int[SupportHelper.SupportedObjects.Count + 1];
                double dTestDist = TrayShipAseemblyServices.GetMaxRouteStructDistance(this, SupportHelper.SupportedObjects.Count, ref indexRoute);

                //=============
                //Create Joints
                //=============
                //Create a Joint Factory
                //Create an object to hold the Joint
                int middleLegs = 0, lowest;
                string routePort = string.Empty;
                double extraLength = 0, trayPlaneOffset = 0;

                lowest = indexRoute[1];

                if (lowest == 1)
                    routePort = "Route";
                else
                    routePort = "Route_" + lowest;

                JointHelper.CreateRigidJoint(part[1], "Route", "-1", routePort, Plane.XY, Plane.XY, Axis.X, Axis.Y, dRouteStPltPlaneOff, dRouteStPltAxisOff, 0);

                for (int index = 2; index <= numOfRoutes; index++)
                {
                    trayPlaneOffset = -(hangerPortOriginZ_Route[lowest] - hangerPortOriginZ_Route[indexRoute[index]]);

                    if (index % 2 == 0)
                        extraLength = gap;
                    JointHelper.CreateRigidJoint(part[1], "Route", part[index], "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, trayPlaneOffset, extraLength, 0);//idxPart[index]
                }

                for (int index = legBegin; index <= legEnd; index++)
                {
                    //Add Joint Between the Ports of Vertical Beam 1
                    JointHelper.CreatePrismaticJoint(part[index], "BeginCap", part[index], "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                }

                if (noOfLegs == 4 || noOfLegs == 6)
                {
                    //Add Joint Between the Horizontal and Vertical Beams 9380
                    JointHelper.CreateRigidJoint(part[1], "CableTray1", part[legBegin], strLegPort2, Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, dLegStPltPlaneOff, -leftOffset + sectionWidth / 2, dLegStOriginOffset);

                    //Add Joint Between the Ports of Vertical Beam 1
                    JointHelper.CreatePointOnPlaneJoint(part[legBegin], strLegPort1, "-1", leftStructPort, Plane.XY);

                    //Add Joint Between the Horizontal and Vertical Beams 2212
                    JointHelper.CreateRigidJoint(part[1], "CableTray1", part[legBegin + 1], strLegPort1, Plane.XY, Plane.XY, Axis.X, Axis.NegativeX, dLegStPltPlaneOff, -leftOffset + sectionWidth / 2, -dLegStOriginOffset);

                    //Add Joint Between the Ports of Vertical Beam 2
                    JointHelper.CreatePointOnPlaneJoint(part[legBegin + 1], strLegPort2, "-1", rightStructPort, Plane.XY);

                    //Add Joint Between the Horizontal and Vertical Beams 1188
                    JointHelper.CreateRigidJoint(part[1], "CableTray2", part[legBegin + 2], strLegPort1, Plane.XY, Plane.XY, Axis.X, Axis.X, dLegStPltPlaneOff, rightOffset - sectionWidth / 2, dLegStOriginOffset);

                    //Add Joint Between the Ports of Vertical Beam 1
                    JointHelper.CreatePointOnPlaneJoint(part[legBegin + 2], strLegPort2, "-1", leftStructPort, Plane.XY);

                    //Add Joint Between the Horizontal and Vertical Beams 2276
                    JointHelper.CreateRigidJoint(part[1], "CableTray2", part[legBegin + 3], strLegPort2, Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeX, dLegStPltPlaneOff, rightOffset - sectionWidth / 2, -dLegStOriginOffset);

                    //Add Joint Between the Ports of Vertical Beam 2
                    JointHelper.CreatePointOnPlaneJoint(part[legBegin + 3], strLegPort1, "-1", rightStructPort, Plane.XY);
                }

                if (noOfLegs == 2 || noOfLegs == 6)
                {
                    if (noOfLegs == 2)
                        middleLegs = legBegin;
                    else
                        middleLegs = legBegin + 4;

                    //Add Joint Between the Horizontal and Vertical Beams 9444
                    JointHelper.CreateRigidJoint(part[1], "Route", part[middleLegs], strLegPort1, Plane.XY, Plane.XY, Axis.X, Axis.X, dLegStPltPlaneOff, -sectionWidth / 2, dLegStOriginOffset);

                    //Add Joint Between the Ports of Vertical Beam 2
                    JointHelper.CreatePointOnPlaneJoint(part[middleLegs], strLegPort2, "-1", rightStructPort, Plane.XY);

                    //Add Joint Between the Horizontal and Vertical Beams 2276
                    JointHelper.CreateRigidJoint(part[1], "Route", part[middleLegs + 1], strLegPort2, Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeX, dLegStPltPlaneOff, -sectionWidth / 2, -dLegStOriginOffset);

                    //Add Joint Between the Ports of Vertical Beam 2
                    JointHelper.CreatePointOnPlaneJoint(part[middleLegs + 1], strLegPort1, "-1", leftStructPort, Plane.XY);
                }

                double dBoltPlaneOffset2 = 0;
                for (int index = 1; index <= numOfRoutes; index++)
                {
                    dBoltPlaneOffset2 = -(hangerPortOriginZ_Route[lowest] - hangerPortOriginZ_Route[indexRoute[index]]);

                    if (connection == 1)
                    {
                        if (noOfLegs == 4 || noOfLegs == 6)
                        {
                            //Add Joint Between the Horizontal and Vertical Beams
                            JointHelper.CreateRigidJoint(clamp[clampBegin + (noOfLegs * (index - 1))], "Route", part[legBegin], strLegPort2, Plane.ZX, Plane.XY, Axis.Z, Axis.NegativeX, dBoltPlaneOffset + dBoltPlaneOffset2, sectionDepth / 2, -clampOD / 2);

                            //Add Joint Between the Horizontal and Vertical Beams
                            JointHelper.CreateRigidJoint(clamp[clampBegin + (noOfLegs * (index - 1)) + 1], "Route", part[legBegin + 1], strLegPort1, Plane.ZX, Plane.XY, Axis.X, Axis.NegativeY, -(dBoltPlaneOffset + dBoltPlaneOffset2), -clampOD / 2, sectionDepth / 2);

                            //Add Joint Between the Horizontal and Vertical Beams 2277
                            JointHelper.CreateRigidJoint(clamp[clampBegin + (noOfLegs * (index - 1)) + 2], "Route", part[legBegin + 2], strLegPort1, Plane.ZX, Plane.NegativeXY, Axis.X, Axis.Y, dBoltPlaneOffset + dBoltPlaneOffset2, -clampOD / 2, -sectionDepth / 2);

                            //Add Joint Between the Horizontal and Vertical Beams
                            JointHelper.CreateRigidJoint(clamp[clampBegin + (noOfLegs * (index - 1)) + 3], "Route", part[legBegin + 3], strLegPort2, Plane.ZX, Plane.NegativeXY, Axis.X, Axis.Y, -(dBoltPlaneOffset + dBoltPlaneOffset2), -clampOD / 2, -sectionDepth / 2);
                        }

                        int dxClamp;
                        if (noOfLegs == 2 || noOfLegs == 6)
                        {
                            if (noOfLegs == 2)
                                dxClamp = clampBegin + noOfLegs * (index - 1);
                            else
                                dxClamp = clampBegin + (noOfLegs * (index - 1)) + 4;

                            //Add Joint Between the Horizontal and Vertical Beams
                            JointHelper.CreateRigidJoint(clamp[dxClamp], "Route", part[middleLegs], strLegPort1, Plane.ZX, Plane.NegativeXY, Axis.X, Axis.Y, dBoltPlaneOffset + dBoltPlaneOffset2, -clampOD / 2, -sectionDepth / 2);

                            //Add Joint Between the Horizontal and Vertical Beams
                            JointHelper.CreateRigidJoint(clamp[dxClamp + 1], "Route", part[middleLegs + 1], strLegPort2, Plane.ZX, Plane.NegativeXY, Axis.X, Axis.Y, -(dBoltPlaneOffset + dBoltPlaneOffset2), -clampOD / 2, -sectionDepth / 2);
                        }
                    }
                    else if (connection == 2)
                    {
                        if (noOfLegs == 4 || noOfLegs == 6)
                        {
                            //Add Joint Between the Horizontal and Vertical Beams
                            JointHelper.CreateRigidJoint(bolt[boltsBegin + (noOfLegs * (index - 1))], "StartOther", part[legBegin], strLegPort2, Plane.ZX, Plane.XY, Axis.Z, Axis.X, dBoltPlaneOffset + dBoltPlaneOffset2, -sectionWidth / 2, boltLen / 2);

                            //Add Joint Between the Horizontal and Vertical Beams
                            JointHelper.CreateRigidJoint(bolt[boltsBegin + (noOfLegs * (index - 1)) + 1], "EndOther", part[legBegin + 1], strLegPort1, Plane.ZX, Plane.XY, Axis.Z, Axis.X, -(dBoltPlaneOffset + dBoltPlaneOffset2), -sectionDepth / 2, -boltLen / 2);

                            //Add Joint Between the Horizontal and Vertical Beams 2277
                            JointHelper.CreateRigidJoint(bolt[boltsBegin + (noOfLegs * (index - 1)) + 2], "EndOther", part[legBegin + 2], strLegPort1, Plane.ZX, Plane.NegativeXY, Axis.X, Axis.Y, dBoltPlaneOffset + dBoltPlaneOffset2, -boltLen / 2, -sectionWidth / 2);

                            //Add Joint Between the Horizontal and Vertical Beams
                            JointHelper.CreateRigidJoint(bolt[boltsBegin + (noOfLegs * (index - 1)) + 3], "StartOther", part[legBegin + 3], strLegPort2, Plane.ZX, Plane.NegativeXY, Axis.X, Axis.Y, -(dBoltPlaneOffset + dBoltPlaneOffset2), boltLen / 2, -sectionDepth / 2);
                        }

                        int dxBolts;
                        if (noOfLegs == 2 || noOfLegs == 6)
                        {
                            if (noOfLegs == 2)
                                dxBolts = boltsBegin + noOfLegs * (index - 1);
                            else
                                dxBolts = boltsBegin + (noOfLegs * (index - 1)) + 4;

                            //Add Joint Between the Horizontal and Vertical Beams
                            JointHelper.CreateRigidJoint(bolt[dxBolts], "EndOther", part[middleLegs], strLegPort1, Plane.ZX, Plane.NegativeXY, Axis.X, Axis.Y, dBoltPlaneOffset + dBoltPlaneOffset2, -boltLen / 2, -sectionWidth / 2);

                            //Add Joint Between the Horizontal and Vertical Beams
                            JointHelper.CreateRigidJoint(bolt[dxBolts + 1], "StartOther", part[middleLegs + 1], strLegPort2, Plane.ZX, Plane.NegativeXY, Axis.X, Axis.Y, -(dBoltPlaneOffset + dBoltPlaneOffset2), boltLen / 2, -sectionDepth / 2);
                        }
                    }
                }

                if (includePad)
                {
                    if (noOfLegs == 4 || noOfLegs == 6)
                    {
                        JointHelper.CreateRigidJoint(part[padBegin], "TopStructure", part[legBegin], strLegPort1, Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, -sectionWidth / 3, sectionWidth / 3);

                        JointHelper.CreateRigidJoint(part[padBegin + 1], "BotStructure", part[legBegin + 1], strLegPort2, Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, -sectionWidth / 3, sectionWidth / 3);

                        JointHelper.CreateRigidJoint(part[padBegin + 2], "BotStructure", part[legBegin + 2], strLegPort2, Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, -sectionWidth / 3, sectionWidth / 3);

                        JointHelper.CreateRigidJoint(part[padBegin + 3], "TopStructure", part[legBegin + 3], strLegPort1, Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, -sectionWidth / 3, sectionWidth / 3);
                    }

                    int dxPartTmp;
                    if (noOfLegs == 2 || noOfLegs == 6)
                    {
                        if (noOfLegs == 2)
                            dxPartTmp = legEnd;
                        else
                            dxPartTmp = legEnd + 4;

                        JointHelper.CreateRigidJoint(part[dxPartTmp], "BotStructure", part[middleLegs], strLegPort2, Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, -sectionWidth / 3, sectionWidth / 3);

                        JointHelper.CreateRigidJoint(part[dxPartTmp + 1], "TopStructure", part[middleLegs + 1], strLegPort1, Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, -sectionWidth / 3, sectionWidth / 3);
                    }
                }
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
                try
                {
                    //Create a collection to hold the ALL Route connection information
                    Collection<ConnectionInfo> routeConnections = new Collection<ConnectionInfo>();

                    routeConnections.Add(new ConnectionInfo(part[1], 1));       //partindex, routeindex

                    //Return the collection of Route connection information.
                    return routeConnections;
                }
                catch (Exception e)
                {
                    Type myType = this.GetType();
                    CmnException e1 = new CmnException("Error in Get Route Connections." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                    throw e1;
                }
            }
        }

        //-----------------------------------------------------------------------------------
        //Get Struct Connections
        //-----------------------------------------------------------------------------------
        public override Collection<ConnectionInfo> SupportingConnections
        {
            get
            {
                try
                {
                    //Create a collection to hold ALL the Structure Connection information
                    Collection<ConnectionInfo> structConnections = new Collection<ConnectionInfo>();

                    structConnections.Add(new ConnectionInfo(part[1], 1));      //partindex, routeindex

                    //Return the collection of Structure connection information.
                    return structConnections;
                }
                catch (Exception e)
                {
                    Type myType = this.GetType();
                    CmnException e1 = new CmnException("Error in Get Struct Connections." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                    throw e1;
                }
            }
        }
    }
}
