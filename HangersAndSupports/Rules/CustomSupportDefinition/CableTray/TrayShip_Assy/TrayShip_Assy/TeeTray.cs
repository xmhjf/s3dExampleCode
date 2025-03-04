//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   TeeTray.cs
//   TrayShip_Assy,Ingr.SP3D.Content.Support.Rules.TeeTray
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
    public class TeeTray : CustomSupportDefinition
    {
        private int shape, connection, plateType, noOfLegs, legBegin, legEnd, partCount, padBegin, padEnd, numClamps, clampBegin, numBolts, boltsBegin;
        private string  uBoltPart;
        private bool includePad, overhangRule;
        private double length, plateHeight, leftOffset, rightOffset, dUnderLength, beginAngle, endAngle, bendRadius, width, height, padDiameter, padThickness;
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
                    includePad = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJUAhsTrayAsmPad", "IncludePad")).PropValue;
                    length = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsTrayAsmTrayL", "L")).PropValue;
                    plateHeight = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsTrayAsmTrayH", "H")).PropValue;
                    leftOffset = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsTrayAsmLeftOH", "L1")).PropValue;
                    rightOffset = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsTrayAsmRightOH", "L2")).PropValue;
                    dUnderLength = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsTrayAsmUnderL", "UnderLength")).PropValue;
                    overhangRule = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJUAhsTrayAsmOverhang", "Overhang")).PropValue;
                    beginAngle = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsTrayAsmBegAngle", "BeginAngle")).PropValue;
                    endAngle = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsTrayAsmEndAngle", "EndAngle")).PropValue;
                    plateType = (int)((PropertyValueCodelist)support.GetPropertyValue("IJUAhsTrayAsmTrayType", "PlateType")).PropValue;

                    CableTrayObjectInfo cabletrayInfo = (CableTrayObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                    width = cabletrayInfo.Width;
                    height = cabletrayInfo.Depth;
                    bendRadius = cabletrayInfo.BendRadius;
                    numOfRoutes = SupportHelper.SupportedObjects.Count;
                    
                    //Create the list of part classes required by the type
                    part = new string[numOfRoutes + 1];
                    for (int index = 1; index <= numOfRoutes; index++)
                    {
                        if (shape == 5)    //Rectangle
                        {
                            part[index] = "idxPart" + index;
                            parts.Add(new PartInfo(part[index], "TeeTray_1"));  //"TrayShip_RectStPlate_1"
                        }
                        else if (shape == 1)  //Round
                        {
                            part[index] = "idxPart" + index;
                            parts.Add(new PartInfo(part[index], "TeeTrayCirRail_1"));   //"TrayShip_RndStPlate_1"
                        }
                    }

                    double positionLength = Math.Abs(length);
                    CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                    IEnumerable<BusinessObject> trayshipPartclass = null;
                    PartClass traySrvStTrayPartClass = (PartClass)catalogBaseHelper.GetPartClass("TrayShipSrv_StNoOfLegs");

                    if (traySrvStTrayPartClass.PartClassType.Equals("HgrServiceClass"))
                        trayshipPartclass = traySrvStTrayPartClass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                    else
                        trayshipPartclass = traySrvStTrayPartClass.Parts;

                    trayshipPartclass = trayshipPartclass.Where(part1 => (string)((PropertyValueString)part1.GetPropertyValue("IJUAhsTraySrvStTray", "UnitType")).PropValue == "mm" && (((double)((PropertyValueDouble)part1.GetPropertyValue("IJUAhsTraySrvStTray", "LengthMin")).PropValue < (plateHeight * 1000)) && ((double)((PropertyValueDouble)part1.GetPropertyValue("IJUAhsTraySrvStTray", "LengthMax")).PropValue > (plateHeight * 1000))));
                    if (trayshipPartclass.Count() > 0)
                    {
                        noOfLegs = 6;
                        if (overhangRule == true)
                        {
                            leftOffset = (double)((PropertyValueDouble)trayshipPartclass.ElementAt(0).GetPropertyValue("IJUAhsTraySrvStTray", "LeftOverhang")).PropValue / 1000;
                            rightOffset = (double)((PropertyValueDouble)trayshipPartclass.ElementAt(0).GetPropertyValue("IJUAhsTraySrvStTray", "RightOverhang")).PropValue / 1000;
                        }
                    }

                    //For Pad
                    legBegin = numOfRoutes + 1;
                    legEnd = legBegin + noOfLegs - 1;

                   
                    PartClass traySrvPadDimPartClass = (PartClass)catalogBaseHelper.GetPartClass("TrayShipSrv_PadDim");

                    if (traySrvPadDimPartClass.PartClassType.Equals("HgrServiceClass"))
                        trayshipPartclass = traySrvPadDimPartClass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                    else
                        trayshipPartclass = traySrvPadDimPartClass.Parts;

                    trayshipPartclass = trayshipPartclass.Where(part12s => (int)((PropertyValueCodelist)part12s.GetPropertyValue("IJUAhsTraySrvPadDim", "SectionSize")).PropValue == (int)(sectionSizeCodelist.PropValue));
                    if (trayshipPartclass.Count() > 0)
                    {
                        padDiameter = (double)((PropertyValueDouble)trayshipPartclass.ElementAt(0).GetPropertyValue("IJUAhsTraySrvPadDim", "PadDia")).PropValue / 1000;
                        padThickness = (double)((PropertyValueDouble)trayshipPartclass.ElementAt(0).GetPropertyValue("IJUAhsTraySrvPadDim", "PadThick")).PropValue / 1000;
                    }

                    //For UBolt
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
                    //For Legs
                    for (int index = legBegin; index <= legEnd; index++)
                    {
                        Array.Resize(ref part, numOfRoutes + 1 + index + 1);
                        part[index] = "idxPart" + index;
                        parts.Add(new PartInfo(part[index], sectionSizeCodelist + " " + "Japan-2005"));
                    }

                    //For Pad
                    if (includePad == true)
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
                            Array.Resize(ref part, numOfRoutes + 1 + index + 1);
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
                CrossSection crosssection = (CrossSection)horizontalSectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];

                double sectionWidth = crosssection.Width;
                double sectionDepth = crosssection.Depth;
                double sectionThickness = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructFlangedSectionDimensions", "tf")).PropValue;
                string hgrPort = string.Empty;
                Matrix4X4 hangerPort = new Matrix4X4();
                double[] hangerPortOriginZ_Route = new double[5];
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
                double boltRad = sectionThickness * 2;
                double boltLen = sectionThickness * 6;
                double rodDiameter = plateHeight;
                double clampOD = rodDiameter;

                for (int index = 1; index <= numOfRoutes; index++)
                {
                    if (shape == 1)   //Round Plate
                    {
                        componentDictionary[part[index]].SetPropertyValue(plateType, "IJOAHgrTrayPlTyp", "PlateType");
                        componentDictionary[part[index]].SetPropertyValue(plateHeight / 2, "IJOAHgrTrayThickness", "Thickness");
                        componentDictionary[part[index]].SetPropertyValue(rodDiameter, "IJOAHgrTrayUBoltDia", "D");
                        componentDictionary[part[index]].SetPropertyValue(bendRadius + rodDiameter, "IJOAHgrTrayRadius", "Radius");
                        componentDictionary[part[index]].SetPropertyValue(width - 2 * rodDiameter, "IJOAHgrTrayWidth", "Width");
                        componentDictionary[part[index]].SetPropertyValue(width - 2 * rodDiameter, "IJOAHgrTrayWidth2", "Width2");
                        PropertyValueCodelist braceCodelist = (PropertyValueCodelist)componentDictionary[part[index]].GetPropertyValue("IJOAHgrTrayPlShape", "PlateShape");
                        braceCodelist.PropValue = 2;
                        componentDictionary[part[index]].SetPropertyValue(braceCodelist.PropValue, "IJOAHgrTrayPlShape", "PlateShape");
                    }
                    else if (shape == 5)
                    {
                        componentDictionary[part[index]].SetPropertyValue(bendRadius, "IJOAHgrTrayRadius", "Radius");
                        componentDictionary[part[index]].SetPropertyValue(width, "IJOAHgrTrayWidth", "Width");
                        componentDictionary[part[index]].SetPropertyValue(width, "IJOAHgrTrayWidth2", "Width2");
                        componentDictionary[part[index]].SetPropertyValue(plateHeight, "IJOAHgrTrayThickness", "Thickness");
                    }
                    componentDictionary[part[index]].SetPropertyValue(length, "IJOAHgrTrayLength", "Length");
                    componentDictionary[part[index]].SetPropertyValue(beginAngle, "IJOAHgrTrayBegAngle", "BeginAngle");
                    componentDictionary[part[index]].SetPropertyValue(endAngle, "IJOAHgrTrayEndAngle", "EndAngle");
                }

                if (connection == 1)      //Clamped
                {
                    for (int index = clampBegin; index <= clampEnd; index++)
                    {
                        componentDictionary[clamp[index]].SetPropertyValue(sectionThickness, "IJOAHgrGenUBoltFlgThick", "FlgThick");
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
                    if (includePad == false)
                        padThickness = 0;
                    if (Configuration == 1)
                    {
                        if (index == legBegin || index == legBegin + 3 || index == legBegin + 4)
                        {
                            componentDictionary[part[index]].SetPropertyValue(dUnderLength, "IJUAHgrOccOverLength", "BeginOverLength");
                            componentDictionary[part[index]].SetPropertyValue(-padThickness, "IJUAHgrOccOverLength", "EndOverLength");
                        }
                        else
                        {
                            componentDictionary[part[index]].SetPropertyValue(-padThickness, "IJUAHgrOccOverLength", "BeginOverLength");
                            componentDictionary[part[index]].SetPropertyValue(dUnderLength, "IJUAHgrOccOverLength", "EndOverLength");
                        }
                    }
                    else if (Configuration == 2)
                    {
                        if (index == legBegin || index == legBegin + 3 || index == legBegin + 4)
                        {
                            componentDictionary[part[index]].SetPropertyValue(-padThickness, "IJUAHgrOccOverLength", "BeginOverLength");
                            componentDictionary[part[index]].SetPropertyValue(dUnderLength, "IJUAHgrOccOverLength", "EndOverLength");
                        }
                        else
                        {
                            componentDictionary[part[index]].SetPropertyValue(dUnderLength, "IJUAHgrOccOverLength", "BeginOverLength");
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

                double routeStPltPlaneOff = 0, legStPltPlaneOff = 0, boltPlaneOffset = 0;
                string legPort1 = string.Empty, legPort2 = string.Empty, legFacePort1 = string.Empty, legFacePort2 = string.Empty, platePort1 = string.Empty, platePort2 = string.Empty;
                TrayShipAseemblyServices.ConfigIndex routeSt = new TrayShipAseemblyServices.ConfigIndex();
                
                if (Configuration == 1)
                {
                    routeSt = new TrayShipAseemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.X, Axis.NegativeX);
                    legPort1 = "BeginCap";
                    legPort2 = "EndCap";
                    legFacePort1 = "BeginFace";
                    legFacePort2 = "EndFace";
                    platePort1 = "TopStructure";
                    platePort2 = "BotStructure";

                    if (shape == 5)    //Rectangle
                    {
                        routeStPltPlaneOff = plateHeight + height / 2;
                        legStPltPlaneOff = 0;
                    }
                    else if (shape == 1)
                    {
                        legStPltPlaneOff = plateHeight / 4;
                        routeStPltPlaneOff = plateHeight * 3 / 4 + height / 2;
                    }
                    boltPlaneOffset = plateHeight / 2 + dUnderLength;
                }
                else if (Configuration == 2)
                {
                    routeSt = new TrayShipAseemblyServices.ConfigIndex(Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeX);

                    legPort1 = "EndCap";
                    legPort2 = "BeginCap";

                    legFacePort1 = "EndFace";
                    legFacePort2 = "BeginFace";

                    platePort1 = "BotStructure";
                    platePort2 = "TopStructure";

                    if (shape == 5)    //Rectangle
                    {
                        routeStPltPlaneOff = -height / 2;
                        legStPltPlaneOff = -plateHeight;
                    }
                    else if (shape == 1)
                    {
                        routeStPltPlaneOff = -plateHeight / 4 - height / 2;
                        legStPltPlaneOff = -plateHeight * 3 / 4;
                    }
                    boltPlaneOffset = -plateHeight / 2 - dUnderLength;
                }

                //=============
                //Create Joints
                //=============
                //Create a Joint Factory
                //Create an object to hold the Joint

                double leg3StPltAxisOff = width / 2 + bendRadius - bendRadius * Math.Cos(beginAngle);
                double leg3StPltOrgOff = -(width / 2 + bendRadius) + bendRadius * Math.Sin(beginAngle);
                double leg4StPltAxisOff = width / 2 + bendRadius - bendRadius * Math.Sin(endAngle);
                double leg4StPltOrgOff = -(width / 2 + bendRadius) + bendRadius * Math.Cos(endAngle);
                double leg5StPltAxisOff = width / 2 + bendRadius - bendRadius * Math.Cos(beginAngle);
                double leg5StPltOrgOff = (width / 2 + bendRadius) - bendRadius * Math.Sin(beginAngle);
                double leg6StPltAxisOff = (width / 2 + bendRadius) - bendRadius * Math.Cos(endAngle);
                double leg6StPltOrgOff = width / 2 + bendRadius - bendRadius * Math.Sin(endAngle);

                componentDictionary[part[legBegin + 2]].SetPropertyValue(beginAngle, "IJOAhsBeginCap", "BeginCapRotZ");
                componentDictionary[part[legBegin + 3]].SetPropertyValue(endAngle, "IJOAhsBeginCap", "BeginCapRotZ");
                componentDictionary[part[legBegin + 4]].SetPropertyValue((Math.PI / 2) - endAngle, "IJOAhsBeginCap", "BeginCapRotZ");
                componentDictionary[part[legBegin + 5]].SetPropertyValue((Math.PI / 2) - beginAngle, "IJOAhsBeginCap", "BeginCapRotZ");

                componentDictionary[part[legBegin + 2]].SetPropertyValue(beginAngle, "IJOAhsEndCap", "EndCapRotZ");
                componentDictionary[part[legBegin + 3]].SetPropertyValue(endAngle, "IJOAhsEndCap", "EndCapRotZ");
                componentDictionary[part[legBegin + 4]].SetPropertyValue(-endAngle, "IJOAhsEndCap", "EndCapRotZ");
                componentDictionary[part[legBegin + 5]].SetPropertyValue((Math.PI / 2) - beginAngle, "IJOAhsEndCap", "EndCapRotZ");

                int lowestIndex;
                string routePort = string.Empty;
                double trayPlaneOffset = 0;
                int[] routeIndex = new int[5];

                double dTestDist = TrayShipAseemblyServices.GetMaxRouteStructDistance(this, SupportHelper.SupportedObjects.Count, ref routeIndex);
                lowestIndex = routeIndex[1];

                if (lowestIndex == 1)
                    routePort = "Route";
                else
                    routePort = "Route_" + lowestIndex;

                JointHelper.CreateRigidJoint(part[1], "Route", "-1", routePort, routeSt.A, routeSt.B, routeSt.C, routeSt.D, routeStPltPlaneOff, 0, 0);



                for (int index = legBegin; index <= legEnd; index++)
                {
                    //Add Joint Between the Ports of Vertical Beam 1
                    JointHelper.CreatePrismaticJoint(part[index], "BeginCap", part[index], "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                }

                JointHelper.CreateRigidJoint(part[1], "CableTray1", part[legBegin], legPort1, Plane.XY, Plane.XY, Axis.X, Axis.Y, -legStPltPlaneOff, -width / 2, leftOffset);

                JointHelper.CreatePointOnPlaneJoint(part[legBegin], legPort2, "-1", leftStructPort, Plane.XY);

                JointHelper.CreateRigidJoint(part[1], "CableTray2", part[legBegin + 1], legPort2, Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeY, -legStPltPlaneOff, -width / 2, -rightOffset);

                JointHelper.CreatePointOnPlaneJoint(part[legBegin + 1], legPort1, "-1", rightStructPort, Plane.XY);

                JointHelper.CreateRigidJoint(part[1], "Route", part[legBegin + 2], legPort2, Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, -legStPltPlaneOff, leg3StPltAxisOff, leg3StPltOrgOff);

                JointHelper.CreatePointOnPlaneJoint(part[legBegin + 2], legPort1, "-1", leftStructPort, Plane.XY);

                JointHelper.CreateRigidJoint(part[1], "Route", part[legBegin + 3], legPort1, Plane.XY, Plane.XY, Axis.X, Axis.NegativeX, -legStPltPlaneOff, leg4StPltAxisOff, leg4StPltOrgOff);

                JointHelper.CreatePointOnPlaneJoint(part[legBegin + 3], legPort2, "-1", rightStructPort, Plane.XY);

                JointHelper.CreateRigidJoint(part[1], "Route", part[legBegin + 4], legPort1, Plane.XY, Plane.XY, Axis.Y, Axis.X, -legStPltPlaneOff, leg5StPltAxisOff, leg5StPltOrgOff);

                JointHelper.CreatePointOnPlaneJoint(part[legBegin + 4], legPort2, "-1", rightStructPort, Plane.XY);

                JointHelper.CreateRigidJoint(part[1], "Route", part[legBegin + 5], legPort2, Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, -legStPltPlaneOff, leg6StPltAxisOff, leg6StPltOrgOff);

                JointHelper.CreatePointOnPlaneJoint(part[legBegin + 5], legPort1, "-1", leftStructPort, Plane.XY);


                string pipePortName = string.Empty;

                for (int idxTray = 2; idxTray <= numOfRoutes; idxTray++)
                {
                    pipePortName = "Route_" + idxTray;
                    if (Configuration == 1)
                        trayPlaneOffset = -(hangerPortOriginZ_Route[lowestIndex] - hangerPortOriginZ_Route[routeIndex[idxTray]]);
                    else if (Configuration == 2)
                        trayPlaneOffset = (hangerPortOriginZ_Route[lowestIndex] - hangerPortOriginZ_Route[routeIndex[idxTray]]);

                    JointHelper.CreateRigidJoint(part[1], "Route", part[idxTray], "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, trayPlaneOffset, 0, 0);
                }

                //Add the Joints between the Clamp or Bolts and the Legs

                double boltPlaneOffset2 = 0;
                for (int idxTray = 1; idxTray <= numOfRoutes; idxTray++)
                {
                    if (Configuration == 1)
                        boltPlaneOffset2 = -(hangerPortOriginZ_Route[lowestIndex] - hangerPortOriginZ_Route[routeIndex[idxTray]]);
                    else if (Configuration == 2)
                        boltPlaneOffset2 = (hangerPortOriginZ_Route[lowestIndex] - hangerPortOriginZ_Route[routeIndex[idxTray]]);

                    if (connection == 1)
                    {
                        JointHelper.CreateRigidJoint(clamp[clampBegin + (noOfLegs * (idxTray - 1))], "Route", part[legBegin], legFacePort1, Plane.ZX, Plane.NegativeXY, Axis.Z, Axis.NegativeX, boltPlaneOffset + boltPlaneOffset2, 0, -clampOD / 2 - sectionWidth / 2);

                        JointHelper.CreateRigidJoint(clamp[clampBegin + (noOfLegs * (idxTray - 1)) + 1], "Route", part[legBegin + 1], legFacePort2, Plane.ZX, Plane.XY, Axis.X, Axis.NegativeY, boltPlaneOffset + boltPlaneOffset2, -clampOD / 2 - sectionWidth / 2, 0);

                        JointHelper.CreateRigidJoint(clamp[clampBegin + (noOfLegs * (idxTray - 1)) + 2], "Route", part[legBegin + 2], legFacePort2, Plane.ZX, Plane.XY, Axis.X, Axis.NegativeY, boltPlaneOffset + boltPlaneOffset2, -clampOD / 2 - sectionWidth / 2, 0);

                        JointHelper.CreateRigidJoint(clamp[clampBegin + (noOfLegs * (idxTray - 1)) + 3], "Route", part[legBegin + 3], legFacePort1, Plane.ZX, Plane.NegativeXY, Axis.Z, Axis.NegativeX, boltPlaneOffset + boltPlaneOffset2, 0, -clampOD / 2 - sectionWidth / 2);

                        JointHelper.CreateRigidJoint(clamp[clampBegin + (noOfLegs * (idxTray - 1)) + 4], "Route", part[legBegin + 4], legFacePort1, Plane.ZX, Plane.NegativeXY, Axis.Z, Axis.NegativeX, boltPlaneOffset + boltPlaneOffset2, 0, -clampOD / 2 - sectionWidth / 2);

                        JointHelper.CreateRigidJoint(clamp[clampBegin + (noOfLegs * (idxTray - 1)) + 5], "Route", part[legBegin + 5], legFacePort2, Plane.ZX, Plane.XY, Axis.X, Axis.NegativeY, boltPlaneOffset + boltPlaneOffset2, -clampOD / 2 - sectionWidth / 2, 0);
                    }
                    else if (connection == 2)
                    {
                        JointHelper.CreateRigidJoint(bolt[boltsBegin + (noOfLegs * (idxTray - 1))], "StartOther", part[legBegin], legFacePort1, Plane.YZ, Plane.XY, Axis.Z, Axis.NegativeX, -(boltPlaneOffset + boltPlaneOffset2), 0, -sectionDepth / 2 + boltLen / 2);

                        JointHelper.CreateRigidJoint(bolt[boltsBegin + (noOfLegs * (idxTray - 1)) + 1], "EndOther", part[legBegin + 1], legFacePort2, Plane.YZ, Plane.NegativeXY, Axis.Z, Axis.NegativeX, -(boltPlaneOffset + boltPlaneOffset2), 0, -(sectionDepth / 2 + boltLen / 2));

                        JointHelper.CreateRigidJoint(bolt[boltsBegin + (noOfLegs * (idxTray - 1)) + 2], "EndOther", part[legBegin + 2], legFacePort2, Plane.YZ, Plane.NegativeXY, Axis.Z, Axis.NegativeX, -(boltPlaneOffset + boltPlaneOffset2), 0, -(sectionDepth / 2 + boltLen / 2));

                        JointHelper.CreateRigidJoint(bolt[boltsBegin + (noOfLegs * (idxTray - 1)) + 3], "StartOther", part[legBegin + 3], legFacePort1, Plane.YZ, Plane.XY, Axis.Z, Axis.NegativeX, -(boltPlaneOffset + boltPlaneOffset2), 0, -sectionDepth / 2 + boltLen / 2);

                        JointHelper.CreateRigidJoint(bolt[boltsBegin + (noOfLegs * (idxTray - 1)) + 4], "StartOther", part[legBegin + 4], legFacePort1, Plane.YZ, Plane.XY, Axis.Z, Axis.NegativeX, -(boltPlaneOffset + boltPlaneOffset2), 0, -sectionDepth / 2 + boltLen / 2);

                        JointHelper.CreateRigidJoint(bolt[boltsBegin + (noOfLegs * (idxTray - 1)) + 5], "EndOther", part[legBegin + 5], legFacePort2, Plane.YZ, Plane.NegativeXY, Axis.Z, Axis.NegativeX, -(boltPlaneOffset + boltPlaneOffset2), 0, -(sectionDepth / 2 + boltLen / 2));
                    }
                }

                //Add Joint between the pads and Legs
                if (includePad == true)
                {
                    JointHelper.CreateRigidJoint(part[padBegin], platePort1, part[legBegin], legFacePort2, Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, 0, 0);

                    JointHelper.CreateRigidJoint(part[padBegin + 1], platePort2, part[legBegin + 1], legFacePort1, Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, 0, 0);

                    JointHelper.CreateRigidJoint(part[padBegin + 2], platePort2, part[legBegin + 2], legFacePort1, Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, 0, 0);

                    JointHelper.CreateRigidJoint(part[padBegin + 3], platePort1, part[legBegin + 3], legFacePort2, Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, 0, 0);

                    JointHelper.CreateRigidJoint(part[padBegin + 4], platePort1, part[legBegin + 4], legFacePort2, Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, 0, 0);

                    JointHelper.CreateRigidJoint(part[padBegin + 5], platePort2, part[legBegin + 5], legFacePort1, Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, 0, 0);
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
