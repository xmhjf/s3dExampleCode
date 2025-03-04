//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   TurnTray.cs
//   TrayShip_Assy,Ingr.SP3D.Content.Support.Rules.TurnTray
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
using Ingr.SP3D.Route.Middle;
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
    public class TurnTray : CustomSupportDefinition
    {
        private int shape, connection, plateType, noOfLegs, LegBegin, legEnd, partCount, padBegin, padEnd, numClamps, clampBegin, numBolts, boltsBegin;
        private string uBoltPart;
        private bool includePad;
        private double plateHeight, underLength, beginAngle, endAngle, bendAngle, throatRadius, width, height, padDiameter, padThickness;
        static int numOfRoutes = 0, clampEnd = 0, boltsEnd = 0;
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
                    beginAngle = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsTrayAsmBegAngle", "BeginAngle")).PropValue;
                    endAngle = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsTrayAsmEndAngle", "EndAngle")).PropValue;
                    PropertyValueCodelist sectionSizeCodelist = (PropertyValueCodelist)support.GetPropertyValue("IJUAhsTrayAsmLSecSize", "SectionSize");
                    includePad = (bool)((PropertyValueBoolean)support.GetPropertyValue("IJUAhsTrayAsmPad", "IncludePad")).PropValue;
                    plateHeight = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsTrayAsmTrayH", "H")).PropValue;
                    underLength = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAhsTrayAsmUnderL", "UnderLength")).PropValue;
                    plateType = (int)((PropertyValueCodelist)support.GetPropertyValue("IJUAhsTrayAsmTrayType", "PlateType")).PropValue;
                    CableTrayObjectInfo ctInfo = (CableTrayObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                    numOfRoutes = SupportHelper.SupportedObjects.Count;
                    width = ctInfo.Width;
                    height = ctInfo.Depth;
                    RouteFeature routeFeature = (RouteFeature)SupportHelper.SupportedObjects[1];
                    Route.Middle.TurnFeature turnFeauture = (TurnFeature)routeFeature;
                    throatRadius = turnFeauture.ThroatRadius;
                    bendAngle = ctInfo.BendAngle;

                    string unitType = "m";
                    double Width = 0;
                    if (unitType == "m")
                    {
                        Width = width * 1000;
                        Width = Math.Round(Width, 0);
                        unitType = "mm";
                    }
                    //Create the list of part classes required by the type
                    part = new string[numOfRoutes + 1];
                    for (int index = 1; index <= numOfRoutes; index++)
                    {
                        if (shape == 5)    //Rectangle
                        {
                            part[index] = "idxPart" + index;
                            parts.Add(new PartInfo(part[index], "TurnTray_1"));  //"TrayShip_RectStPlate_1"
                        }
                        else if (shape == 1)  //Round
                        {
                            part[index] = "idxPart" + index;
                            parts.Add(new PartInfo(part[index], "TurnTrayCRail_1"));   //"TrayShip_RndStPlate_1"
                        }
                    }

                    CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                    IEnumerable<BusinessObject> trayshipPartclass = null;
                    PartClass traySrvStTrayPartClass = (PartClass)catalogBaseHelper.GetPartClass("TrayShipSrv_NoOfLegs");

                    if (traySrvStTrayPartClass.PartClassType.Equals("HgrServiceClass"))
                        trayshipPartclass = traySrvStTrayPartClass.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                    else
                        trayshipPartclass = traySrvStTrayPartClass.Parts;

                    trayshipPartclass = trayshipPartclass.Where(part1 => (string)((PropertyValueString)part1.GetPropertyValue("IJUAhsTraySrvTurnTray", "UnitType")).PropValue == unitType && (((double)((PropertyValueDouble)part1.GetPropertyValue("IJUAhsTraySrvTurnTray", "WidthMin")).PropValue < (Width)) && ((double)((PropertyValueDouble)part1.GetPropertyValue("IJUAhsTraySrvTurnTray", "WidthMax")).PropValue > (Width))));
                    if (trayshipPartclass.Count() > 0)
                        noOfLegs = Convert.ToInt32((string)((PropertyValueString)trayshipPartclass.ElementAt(0).GetPropertyValue("IJUAhsTraySrvTurnTray", "NoOfLegs")).PropValue);


                    
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
                    //For Turn Plate
                    LegBegin = numOfRoutes + 1;
                    legEnd = LegBegin + noOfLegs - 1;

                    //For Legs
                    for (int index = LegBegin; index <= legEnd; index++)
                    {
                        Array.Resize(ref part, numOfRoutes + 1 + index + 1);
                        part[index] = "idxPart" + index;
                        parts.Add(new PartInfo(part[index], sectionSizeCodelist + " " + "Japan-2005"));
                    }

                    //For Pad
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
                return 1;
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

                BoundingBoxHelper.CreateStandardBoundingBoxes(true);
                BoundingBox boundingBox;

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    boundingBox = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.SupportedAndSupporting);
                else
                    boundingBox = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.Supported);

                BusinessObject horizontalSectionPart = componentDictionary[part[LegBegin]].GetRelationship("madeFrom", "part").TargetObjects[0];
                CrossSection crosssection = (CrossSection)horizontalSectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];

                double sectionWidth = crosssection.Width;
                double sectionDepth = crosssection.Depth;
                double sectionThickness = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructFlangedSectionDimensions", "tf")).PropValue;
                string hgrPort = string.Empty;
                Matrix4X4 hangerPort = new Matrix4X4();
                double[] hangerPortOriginZ_Route = new double[5];
                for (int indexRoute = 1; indexRoute <= numOfRoutes; indexRoute++)
                {
                    if (indexRoute == 1)
                        hgrPort = "Route";
                    else
                        hgrPort = "Route_" + indexRoute;

                    hangerPort = RefPortHelper.PortLCS(hgrPort);
                    hangerPortOriginZ_Route[indexRoute] = hangerPort.Origin.Z;
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

                for (int index = 1; index <= numOfRoutes; index++)
                {
                    if (shape == 1)   //Round Plate
                    {
                        componentDictionary[part[index]].SetPropertyValue(plateType, "IJOAHgrTrayPlTyp", "PlateType");
                        componentDictionary[part[index]].SetPropertyValue(plateHeight, "IJOAHgrTrayUBoltDia", "D");
                        componentDictionary[part[index]].SetPropertyValue(plateHeight / 2, "IJOAHgrTrayThickness", "Thickness");
                        componentDictionary[part[index]].SetPropertyValue(width, "IJOAHgrTrayWidth", "Width");
                    }
                    else if (shape == 5)   //Rectangle Plate
                    {
                        componentDictionary[part[index]].SetPropertyValue(width, "IJOAHgrTrayWidth", "Width");
                        componentDictionary[part[index]].SetPropertyValue(plateHeight, "IJOAHgrTrayThickness", "Thickness");
                    }
                    componentDictionary[part[index]].SetPropertyValue(throatRadius, "IJOAHgrTrayInRadius", "InsideRadius");
                    componentDictionary[part[index]].SetPropertyValue(bendAngle, "IJOAHgrTrayAngle", "Angle");
                    componentDictionary[part[index]].SetPropertyValue(beginAngle, "IJOAHgrTrayBegAngle", "BeginAngle");
                    componentDictionary[part[index]].SetPropertyValue(endAngle, "IJOAHgrTrayEndAngle", "EndAngle");
                }

                double boltRad = sectionThickness * 2;
                double boltLen = sectionThickness * 6;
                double clampOD = plateHeight;

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

                for (int index = LegBegin; index <= legEnd; index++)
                {
                    PropertyValueCodelist braceCodelist = (PropertyValueCodelist)componentDictionary[part[index]].GetPropertyValue("IJOAhsSteelCP", "CP1");
                    braceCodelist.PropValue = 1;
                    componentDictionary[part[index]].SetPropertyValue(braceCodelist.PropValue, "IJOAhsSteelCP", "CP1");
                    componentDictionary[part[index]].SetPropertyValue(0.0, "IJOAhsBeginCap", "BeginCapRotZ");
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

                if (noOfLegs == 5)
                {
                    PropertyValueCodelist braceCodelist = (PropertyValueCodelist)componentDictionary[part[LegBegin + 4]].GetPropertyValue("IJOAhsSteelCP", "CP1");
                    braceCodelist.PropValue = 4;
                    componentDictionary[part[LegBegin + 4]].SetPropertyValue(braceCodelist.PropValue, "IJOAhsSteelCP", "CP1");
                }
                else if (noOfLegs == 6)
                {
                    PropertyValueCodelist braceCodelist = (PropertyValueCodelist)componentDictionary[part[LegBegin + 4]].GetPropertyValue("IJOAhsSteelCP", "CP1");
                    braceCodelist.PropValue = 4;
                    componentDictionary[part[LegBegin + 4]].SetPropertyValue(braceCodelist.PropValue, "IJOAhsSteelCP", "CP1");
                    braceCodelist.PropValue = 2;
                    componentDictionary[part[LegBegin + 5]].SetPropertyValue(braceCodelist.PropValue, "IJOAhsSteelCP", "CP1");
                    componentDictionary[part[LegBegin + 5]].SetPropertyValue(3 * Math.PI / 2, "IJOAhsBeginCap", "BeginCapRotZ");
                }
                //Set attributes on assembly object
                support.SetPropertyValue(connection, "IJUAhsTrayAsmConn", "Connection");
                support.SetPropertyValue(shape, "IJUAhsTrayAsmTrayShape", "Shape");
                support.SetPropertyValue(noOfLegs.ToString(), "IJUAhsTrayAsmNoOfLegs", "NoOfLegs");

                double boltPlaneOffset = 0, trayPlaneOffset = 0;
                string legPort1 = string.Empty, legPort2 = string.Empty, legFacePort1 = string.Empty, legFacePort2 = string.Empty, overLength1 = string.Empty, overLength2 = string.Empty, padPort1 = string.Empty, padPort2 = string.Empty;
                string platePort1 = string.Empty, platePort2 = string.Empty;
                TrayShipAseemblyServices.ConfigIndex routeTurn = new TrayShipAseemblyServices.ConfigIndex();
                int connectionValue = 0;

                double legTurnPltPlaneOff = 0, routeTurnPltPlaneOff = 0;
                double routeTurnPltAxisOff = (throatRadius + width / 2) * Math.Tan(bendAngle / 2);
                double routeTurnPltOriginOff = throatRadius + width / 2;
                double angle1 = Math.Round(RefPortHelper.AngleBetweenPorts("TurnRef", PortAxisType.Y, "Route", PortAxisType.X, OrientationAlong.Direct), 4);
                double angle2 = Math.Round(RefPortHelper.AngleBetweenPorts("TurnRef", PortAxisType.X, "Route", PortAxisType.X, OrientationAlong.Direct), 4);

                if ((bendAngle - 0.001) < Math.PI / 2 && (bendAngle + 0.001) > Math.PI / 2)
                {
                    if (angle1 > Math.PI / 2 && angle1 < Math.PI)
                    {
                        if (angle2 > 0 && angle2 < Math.PI / 2)
                            connectionValue = 1;
                        else if (angle2 > Math.PI / 2 && angle2 < Math.PI)
                            connectionValue = 2;
                    }
                    else if (angle1 > 0 && angle1 < Math.PI / 2)
                    {
                        if (angle2 > 0 && angle2 < Math.PI / 2)
                            connectionValue = 7;
                        else if (angle2 > Math.PI / 2 && angle2 < Math.PI)
                            connectionValue = 6;
                    }
                }
                else
                {
                    if (angle1 > 0 && angle1 < (Math.PI / 2 - bendAngle))
                    {
                        if (angle2 > 0 && angle2 < (Math.PI / 2 - beginAngle))
                            connectionValue = 8;
                        else if (angle2 > (Math.PI / 2 - beginAngle) && angle2 < Math.PI / 2)
                            connectionValue = 4;
                        else if (angle2 > Math.PI / 2 && angle2 < (Math.PI / 2 + beginAngle))
                            connectionValue = 3;
                        else if (angle2 > (Math.PI / 2 + beginAngle) && angle2 < Math.PI)
                            connectionValue = 5;
                    }
                    else if (angle1 > (Math.PI / 2 - bendAngle) && angle1 < Math.PI / 2)
                    {
                        if (angle2 > 0 && angle2 < (Math.PI / 2 - beginAngle))
                            connectionValue = 7;
                        else if (angle2 > (Math.PI / 2 - beginAngle) && angle2 < Math.PI / 2)
                            connectionValue = 1;
                        else if (angle2 > Math.PI / 2 && angle2 < (Math.PI / 2 + beginAngle))
                            connectionValue = 1;
                        else if (angle2 > (Math.PI / 2 + beginAngle) && angle2 < Math.PI)
                            connectionValue = 6;
                    }
                    else if (angle1 > Math.PI / 2 && angle1 < (Math.PI / 2 + bendAngle))
                    {
                        if (angle2 > 0 && angle2 < (Math.PI / 2 - beginAngle))
                            connectionValue = 6;
                        else if (angle2 > (Math.PI / 2 - beginAngle) && angle2 < Math.PI / 2)
                            connectionValue = 1;
                        else if (angle2 > Math.PI / 2 && angle2 < (Math.PI / 2 + beginAngle))
                            connectionValue = 2;
                        else if (angle2 > (Math.PI / 2 + beginAngle) && angle2 < Math.PI)
                            connectionValue = 7;
                    }
                    else if (angle1 > (Math.PI / 2 + bendAngle) && angle1 < Math.PI)
                    {
                        if (angle2 > 0 && angle2 < (Math.PI / 2 - beginAngle))
                            connectionValue = 8;
                        else if (angle2 > (Math.PI / 2 - beginAngle) && angle2 < Math.PI / 2)
                            connectionValue = 1;
                        else if (angle2 > Math.PI / 2 && angle2 < (Math.PI / 2 + beginAngle))
                            connectionValue = 2;
                        else if (angle2 > (Math.PI / 2 + beginAngle) && angle2 < Math.PI)
                            connectionValue = 7;
                    }
                }

                switch (connectionValue)
                {
                    case 1:
                        routeTurn = new TrayShipAseemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.X, Axis.X);
                        break;
                    case 2:
                        routeTurn = new TrayShipAseemblyServices.ConfigIndex(Plane.XY, Plane.NegativeXY, Axis.X, Axis.X);
                        break;
                    case 3:
                        routeTurn = new TrayShipAseemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.X, Axis.NegativeX);
                        break;
                    case 4:
                        routeTurn = new TrayShipAseemblyServices.ConfigIndex(Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeX);
                        break;
                    case 5:
                        routeTurn = new TrayShipAseemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.X, Axis.Y);
                        break;
                    case 6:
                        routeTurn = new TrayShipAseemblyServices.ConfigIndex(Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y);
                        break;
                    case 7:
                        routeTurn = new TrayShipAseemblyServices.ConfigIndex(Plane.XY, Plane.XY, Axis.X, Axis.NegativeY);
                        break;
                    case 8:
                        routeTurn = new TrayShipAseemblyServices.ConfigIndex(Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeY);
                        break;
                }

                if (connectionValue == 1 || connectionValue == 3 || connectionValue == 5 || connectionValue == 7)
                {
                    if (shape == 5)          //Rectangle
                    {
                        legTurnPltPlaneOff = 0;
                        routeTurnPltPlaneOff = plateHeight + height / 2;
                        boltPlaneOffset = plateHeight / 2;
                    }
                    else if (shape == 1)      //Round
                    {
                        legTurnPltPlaneOff = -plateHeight / 4;
                        routeTurnPltPlaneOff = plateHeight * 3 / 4 + height / 2;
                        boltPlaneOffset = plateHeight / 2;
                    }

                    legPort1 = "BeginCap";
                    legPort2 = "EndCap";

                    legFacePort1 = "BeginFace";
                    legFacePort2 = "EndFace";

                    overLength1 = "BeginOverLength";
                    overLength2 = "EndOverLength";

                    padPort1 = "BotStructure";
                    padPort2 = "TopStructure";
                }
                else if (connectionValue == 2 || connectionValue == 4 || connectionValue == 6 || connectionValue == 8)
                {
                    if (shape == 5)          //Rectangle
                    {
                        legTurnPltPlaneOff = plateHeight;
                        routeTurnPltPlaneOff = -height / 2;
                        boltPlaneOffset = -plateHeight / 2;
                    }
                    else if (shape == 1)      //Round
                    {
                        legTurnPltPlaneOff = plateHeight * 3 / 4;
                        routeTurnPltPlaneOff = -plateHeight / 4 - height / 2;
                        boltPlaneOffset = -plateHeight / 2;
                    }

                    legPort1 = "EndCap";
                    legPort2 = "BeginCap";

                    legFacePort1 = "EndFace";
                    legFacePort2 = "BeginFace";

                    overLength1 = "EndOverLength";
                    overLength2 = "BeginOverLength";

                    padPort1 = "TopStructure";
                    padPort2 = "BotStructure";
                }

                double extraLength;
                int lowestIndex;
                int[] routeIndex = new int[SupportHelper.SupportedObjects.Count+1];

                double tempLength = TrayShipAseemblyServices.GetMaxRouteStructDistance(this, SupportHelper.SupportedObjects.Count, ref routeIndex);
                double structTurnDis = RefPortHelper.DistanceBetweenPorts("Structure", "TurnRef", PortDistanceType.Vertical);
                lowestIndex = routeIndex[1];

                if (numOfRoutes > 1)
                {
                    if (!(lowestIndex == 1))
                        extraLength = tempLength - structTurnDis;
                    else
                        extraLength = 0;
                }
                else
                    extraLength = 0;

                for (int index = LegBegin; index <= legEnd; index++)
                {
                    PropertyValueCodelist braceCodelist = (PropertyValueCodelist)componentDictionary[part[LegBegin]].GetPropertyValue("IJOAhsSteelCP", "CP1");
                    braceCodelist.PropValue = 1;
                    componentDictionary[part[index]].SetPropertyValue(braceCodelist.PropValue, "IJOAhsSteelCP", "CP1");
                    componentDictionary[part[index]].SetPropertyValue(0.0, "IJOAhsBeginCap", "BeginCapRotZ");
                    componentDictionary[part[index]].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                    componentDictionary[part[index]].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");

                    if (includePad == false)
                        padThickness = 0;

                    if (index == LegBegin || index == LegBegin + 1)
                    {
                        componentDictionary[part[index]].SetPropertyValue(underLength + extraLength, "IJUAHgrOccOverLength", overLength2);
                        componentDictionary[part[index]].SetPropertyValue(-padThickness, "IJUAHgrOccOverLength", overLength1);
                    }
                    else
                    {
                        componentDictionary[part[index]].SetPropertyValue(underLength + extraLength, "IJUAHgrOccOverLength", overLength1);
                        componentDictionary[part[index]].SetPropertyValue(-padThickness, "IJUAHgrOccOverLength", overLength2);
                    }
                }

                //=============
                //Create Joints
                //=============
                //Create a Joint Factory
                //Create an object to hold the Joint

                JointHelper.CreateRigidJoint(part[1], "Route", "-1", "TurnRef", routeTurn.A, routeTurn.B, routeTurn.C, routeTurn.D, routeTurnPltPlaneOff, routeTurnPltAxisOff, routeTurnPltOriginOff);

                for (int index = LegBegin; index <= legEnd; index++)
                {
                    //Add Joint Between the Ports of Vertical Beam 1
                    JointHelper.CreatePrismaticJoint(part[index], "BeginCap", part[index], "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                }

                //Add Joint Between the Horizontal and Vertical Beams 9380
                JointHelper.CreateRigidJoint(part[1], "CableTray1", part[LegBegin], legPort2, Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, legTurnPltPlaneOff, 0, width / 2);

                //Add Joint Between the Ports of Vertical Beam 1
                JointHelper.CreatePointOnPlaneJoint(part[LegBegin], legPort1, "-1", leftStructPort, Plane.XY);

                //Add Joint Between the Horizontal and Vertical Beams 2212
                JointHelper.CreateRigidJoint(part[1], "CableTray1", part[LegBegin + 1], legPort2, Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeY, legTurnPltPlaneOff, 0, -width / 2);

                //Add Joint Between the Ports of Vertical Beam 2
                JointHelper.CreatePointOnPlaneJoint(part[LegBegin + 1], legPort1, "-1", rightStructPort, Plane.XY);

                //Add Joint Between the Horizontal and Vertical Beams 1188
                JointHelper.CreateRigidJoint(part[1], "CableTray2", part[LegBegin + 2], legPort1, Plane.XY, Plane.XY, Axis.X, Axis.X, legTurnPltPlaneOff, 0, width / 2);

                //Add Joint Between the Ports of Vertical Beam 1
                JointHelper.CreatePointOnPlaneJoint(part[LegBegin + 2], legPort2, "-1", leftStructPort, Plane.XY);

                //Add Joint Between the Horizontal and Vertical Beams 2212
                JointHelper.CreateRigidJoint(part[1], "CableTray2", part[LegBegin + 3], legPort1, Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, legTurnPltPlaneOff, 0, -width / 2);

                //Add Joint Between the Ports of Vertical Beam 2
                JointHelper.CreatePointOnPlaneJoint(part[LegBegin + 3], legPort2, "-1", rightStructPort, Plane.XY);

                if (noOfLegs == 5)
                {
                    //Add Joint Between the Horizontal and Vertical Beams 2212
                    JointHelper.CreateRigidJoint(part[1], "CableTray3", part[LegBegin + 4], legPort1, Plane.XY, Plane.XY, Axis.X, Axis.X, legTurnPltPlaneOff, 0, width / 2);

                    //Add Joint Between the Ports of Vertical Beam 2
                    JointHelper.CreatePointOnPlaneJoint(part[LegBegin + 4], legPort2, "-1", rightStructPort, Plane.XY);
                }
                else if (noOfLegs == 6)
                {
                    //Add Joint Between the Horizontal and Vertical Beams 2212
                    JointHelper.CreateRigidJoint(part[1], "CableTray3", part[LegBegin + 4], legPort1, Plane.XY, Plane.XY, Axis.X, Axis.X, legTurnPltPlaneOff, 0, width / 2);

                    //Add Joint Between the Ports of Vertical Beam 2
                    JointHelper.CreatePointOnPlaneJoint(part[LegBegin + 4], legPort2, "-1", rightStructPort, Plane.XY);

                    //Add Joint Between the Horizontal and Vertical Beams 2212
                    JointHelper.CreateRigidJoint(part[1], "CableTray3", part[LegBegin + 5], legPort1, Plane.XY, Plane.XY, Axis.X, Axis.X, legTurnPltPlaneOff, 0, -width / 2);

                    //Add Joint Between the Ports of Vertical Beam 2
                    JointHelper.CreatePointOnPlaneJoint(part[LegBegin + 5], legPort2, "-1", leftStructPort, Plane.XY);
                }

                string pipePortName = string.Empty;

                for (int idxTray = 2; idxTray <= numOfRoutes; idxTray++)
                {
                    pipePortName = "Route_" + idxTray;
                    trayPlaneOffset = hangerPortOriginZ_Route[1] - hangerPortOriginZ_Route[idxTray];
                    if (connectionValue == 1 || connectionValue == 3 || connectionValue == 5 || connectionValue == 7)
                        trayPlaneOffset = -trayPlaneOffset;
                    else
                        trayPlaneOffset = hangerPortOriginZ_Route[1] - hangerPortOriginZ_Route[idxTray];

                    JointHelper.CreateRigidJoint(part[1], "Route", part[idxTray], "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, trayPlaneOffset, 0, 0);
                }

                //Add the Joints between the Clamp or Bolts and the Legs

                double boltPlaneOffset2 = 0;
                for (int idxTray = 1; idxTray <= numOfRoutes; idxTray++)
                {
                    boltPlaneOffset2 = hangerPortOriginZ_Route[1] - hangerPortOriginZ_Route[idxTray];
                    if (connectionValue == 1 || connectionValue == 3 || connectionValue == 5 || connectionValue == 7)
                        boltPlaneOffset2 = -boltPlaneOffset2;
                    else if (Configuration == 2)
                        boltPlaneOffset2 = hangerPortOriginZ_Route[1] - hangerPortOriginZ_Route[idxTray];

                    if (connection == 1)
                    {
                        //Add Joint Between the Horizontal and Vertical Beams
                        JointHelper.CreateRigidJoint(clamp[clampBegin + (noOfLegs * (idxTray - 1))], "Route", part[LegBegin], legPort2, Plane.ZX, Plane.XY, Axis.Z, Axis.NegativeX, boltPlaneOffset + boltPlaneOffset2, sectionWidth / 2, -clampOD / 2);

                        //Add Joint Between the Horizontal and Vertical Beams
                        JointHelper.CreateRigidJoint(clamp[clampBegin + (noOfLegs * (idxTray - 1)) + 1], "Route", part[LegBegin + 1], legPort2, Plane.XY, Plane.NegativeZX, Axis.X, Axis.X, -clampOD / 2, boltPlaneOffset + boltPlaneOffset2, -sectionWidth / 2);

                        //Add Joint Between the Horizontal and Vertical Beams 2277
                        JointHelper.CreateRigidJoint(clamp[clampBegin + (noOfLegs * (idxTray - 1)) + 2], "Route", part[LegBegin + 2], legPort1, Plane.ZX, Plane.NegativeXY, Axis.X, Axis.Y, boltPlaneOffset + boltPlaneOffset2, -clampOD / 2, -sectionWidth / 2);

                        //Add Joint Between the Horizontal and Vertical Beams
                        JointHelper.CreateRigidJoint(clamp[clampBegin + (noOfLegs * (idxTray - 1)) + 3], "Route", part[LegBegin + 3], legPort1, Plane.ZX, Plane.NegativeXY, Axis.Z, Axis.NegativeY, boltPlaneOffset + boltPlaneOffset2, sectionWidth / 2, -clampOD / 2);

                        if (noOfLegs == 5)
                            //Add Joint Between the Horizontal and Vertical Beams
                            JointHelper.CreateRigidJoint(clamp[clampBegin + (noOfLegs * (idxTray - 1)) + 4], "Route", part[LegBegin + 4], legPort1, Plane.ZX, Plane.NegativeXY, Axis.X, Axis.Y, boltPlaneOffset + boltPlaneOffset2, -clampOD / 2, 0);
                        else if (noOfLegs == 6)
                        {
                            //Add Joint Between the Horizontal and Vertical Beams
                            JointHelper.CreateRigidJoint(clamp[clampBegin + (noOfLegs * (idxTray - 1)) + 4], "Route", part[LegBegin + 4], legPort1, Plane.ZX, Plane.NegativeXY, Axis.X, Axis.Y, boltPlaneOffset + boltPlaneOffset2, -clampOD / 2, 0);

                            //Add Joint Between the Horizontal and Vertical Beams
                            JointHelper.CreateRigidJoint(clamp[clampBegin + (noOfLegs * (idxTray - 1)) + 5], "Route", part[LegBegin + 5], legPort1, Plane.ZX, Plane.NegativeXY, Axis.X, Axis.NegativeY, boltPlaneOffset + boltPlaneOffset2, -clampOD / 2, 0);
                        }
                    }
                    else if (connection == 2)
                    {
                        //Add Joint Between the Horizontal and Vertical Beams
                        JointHelper.CreateRigidJoint(bolt[boltsBegin + (noOfLegs * (idxTray - 1))], "StartOther", part[LegBegin], legPort2, Plane.ZX, Plane.XY, Axis.Z, Axis.X, boltPlaneOffset + boltPlaneOffset2, -sectionWidth / 2, boltLen / 2);

                        //Add Joint Between the Horizontal and Vertical Beams
                        JointHelper.CreateRigidJoint(bolt[boltsBegin + (noOfLegs * (idxTray - 1)) + 1], "EndOther", part[LegBegin + 1], legPort2, Plane.XY, Plane.NegativeZX, Axis.X, Axis.X, -boltLen / 2, boltPlaneOffset + boltPlaneOffset2, -sectionWidth / 2);

                        //Add Joint Between the Horizontal and Vertical Beams 2277
                        JointHelper.CreateRigidJoint(bolt[boltsBegin + (noOfLegs * (idxTray - 1)) + 2], "EndOther", part[LegBegin + 2], legPort1, Plane.ZX, Plane.NegativeXY, Axis.X, Axis.Y, boltPlaneOffset + boltPlaneOffset2, -boltLen / 2, -sectionWidth / 2);

                        //Add Joint Between the Horizontal and Vertical Beams
                        JointHelper.CreateRigidJoint(bolt[boltsBegin + (noOfLegs * (idxTray - 1)) + 3], "StartOther", part[LegBegin + 3], legPort1, Plane.ZX, Plane.NegativeXY, Axis.Z, Axis.Y, boltPlaneOffset + boltPlaneOffset2, -sectionWidth / 2, boltLen / 2);

                        if (noOfLegs == 5)
                            //Add Joint Between the Horizontal and Vertical Beams
                            JointHelper.CreateRigidJoint(bolt[boltsBegin + (noOfLegs * (idxTray - 1)) + 4], "EndOther", part[LegBegin + 4], legPort1, Plane.ZX, Plane.NegativeXY, Axis.X, Axis.Y, boltPlaneOffset + boltPlaneOffset2, -boltLen / 2, 0);
                        else if (noOfLegs == 6)
                        {
                            //Add Joint Between the Horizontal and Vertical Beams
                            JointHelper.CreateRigidJoint(bolt[boltsBegin + (noOfLegs * (idxTray - 1)) + 4], "EndOther", part[LegBegin + 4], legPort1, Plane.ZX, Plane.NegativeXY, Axis.X, Axis.Y, boltPlaneOffset + boltPlaneOffset2, -boltLen / 2, 0);

                            //Add Joint Between the Horizontal and Vertical Beams
                            JointHelper.CreateRigidJoint(bolt[boltsBegin + (noOfLegs * (idxTray - 1)) + 5], "StartOther", part[LegBegin + 5], legPort1, Plane.ZX, Plane.NegativeXY, Axis.X, Axis.Y, boltPlaneOffset + boltPlaneOffset2, boltLen / 2, 0);
                        }
                    }
                }

                int dxLeg;
                if (includePad == true)
                {
                    for (int index = padBegin; index <= padEnd; index++)
                    {
                        dxLeg = index - noOfLegs;
                        if (dxLeg == LegBegin || dxLeg == LegBegin + 1)
                            JointHelper.CreateRigidJoint(part[index], padPort1, part[dxLeg], legFacePort1, Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, sectionWidth / 4, -sectionWidth / 4);
                        else
                            JointHelper.CreateRigidJoint(part[index], padPort2, part[dxLeg], legFacePort2, Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, sectionWidth / 4, -sectionWidth / 4);
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
