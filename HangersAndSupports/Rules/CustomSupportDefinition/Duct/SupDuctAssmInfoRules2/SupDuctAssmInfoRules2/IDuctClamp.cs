//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   IDuctClamp.cs
//   SupDuctAssmInfoRules2,Ingr.SP3D.Content.Support.Rules.IDuctClamp
//   Author       : Vinay
//   Creation Date:  27/11/2015
//   Description:    

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  21.Mar.2016     PVK      TR-CP-288920	Issues found in HS_Assembly_V2
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Support.Middle;

namespace Ingr.SP3D.Content.Support.Rules
{
    //-----------------------------------------------------------------------------------
    //Namespace of this class is Ingr.SP3D.Content.Support.Rules
    //It is recommended that customers specify namespace of their symbols to be
    //CompanyName.SP3D.Content.Specialization.
    //It is also recommended that if customers want to change this symbol to suit their
    //requirements, they should change namespace/symbol name so the identity of the modified
    //symbol will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------
    public class IDuctClamp : CustomSupportDefinition
    {
        private const string HGRBEAM = "HGRBEAM"; //Hanger beam object
        private const string DUCTCLAMP = "DUCTCLAMP";          //Hanger beam object as angle guide
        private const string LEFTPAD = "LEFTPAD";
        private const string RIGHTPAD = "RIGHTPAD";
        private const string CONNECTION1 = "CONNECTION1";
        private const string CONNECTION2 = "CONNECTION2";
        private const string CONNECTION3 = "CONNECTION3";
        private const string CONNECTION4 = "CONNECTION4";
        //Constants
        bool leftPad = false, rightPad = false;
        double D = 0, G = 0;
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
                    object[] AttributeColl = SupDuctAssemblyServices.GetDuctStructuralASMAttributes(this);
                    D = (double)AttributeColl[0];
                    G = (double)AttributeColl[1];
                    leftPad = (bool)AttributeColl[2];
                    rightPad = (bool)AttributeColl[3];


                    //define pads
                    string sectionPartClass = (string)((PropertyValueString)support.SupportDefinition.GetPropertyValue("IJUAHSA_SecPartClass", "SecPartClass")).PropValue;

                    parts.Add(new PartInfo(HGRBEAM, sectionPartClass, "PartSelectionRule,Ingr.SP3D.Content.Support.Rules.CPartByCrossSection"));
                    parts.Add(new PartInfo(CONNECTION1, "Connection"));
                    parts.Add(new PartInfo(CONNECTION2, "Connection"));
                    parts.Add(new PartInfo(CONNECTION3, "Connection"));
                    parts.Add(new PartInfo(CONNECTION4, "Connection"));

                    DuctObjectInfo duct = (DuctObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                    if (duct.CrossSectionShape == CrossSectionShape.Rectangular)
                        parts.Add(new PartInfo(DUCTCLAMP, "S3Dhs_DuctClamp_1"));
                    else
                        parts.Add(new PartInfo(DUCTCLAMP, "S3Dhs_DuctClamp_2"));

                    //define pads
                    if (leftPad)
                        parts.Add(new PartInfo(LEFTPAD, SupDuctAssemblyServices.GetPadPartNameByCrossSectionType(this, parts[0])));
                    if (rightPad)
                        parts.Add(new PartInfo(RIGHTPAD, SupDuctAssemblyServices.GetPadPartNameByCrossSectionType(this, parts[0])));


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
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
                //==========================
                //1. Load standard bounding box definition
                //==========================
                BoundingBoxHelper.CreateStandardBoundingBoxes(false);
                BoundingBox boundingBox;
                //====== ======
                //2. retrieve dimension of the bounding box
                //====== ======
                // Get route box geometry

                //  ____________________
                // |                    |
                // |  ROUTE BOX BOUND   | dHeight
                // |____________________|
                //        dWidth

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    boundingBox = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.SupportedAndSupporting);
                else
                    boundingBox = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.Supported);
                bool[] offsetApplied = new bool[3];

                double boundingBoxWidth = boundingBox.Width, boundingBoxHeight = boundingBox.Height;

                if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                {
                    offsetApplied[1] = false;
                    offsetApplied[2] = true;
                }
                else
                {
                    if (SupportHelper.SupportingObjects.Count > 1)
                        offsetApplied[2] = false;
                    else
                        offsetApplied[2] = true;
                }

                //            '====================================
                //3. Retrieve route/structure configuration and
                //check if offset value is to be applied on lug end
                //If the leg is attached to a stiffener, no offset needs to be specified
                //If the leg is attached to a slab, the offest value will be retrieved from
                //predefined offset rule
                //====================================
                // get cross section dimensions
                BusinessObject horizontalSectionPart = componentDictionary[HGRBEAM].GetRelationship("madeFrom", "part").TargetObjects[0];
                CrossSection crosssection = (CrossSection)horizontalSectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];

                //Get SteelWidth and SteelDepth 
                double steelWidth = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructCrossSectionDimensions", "Width")).PropValue;
                double steelDepth = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructCrossSectionDimensions", "Depth")).PropValue;
                BusinessObject ductClampPart = componentDictionary[DUCTCLAMP].GetRelationship("madeFrom", "part").TargetObjects[0];
                double boltDia = (double)((PropertyValueDouble)ductClampPart.GetPropertyValue("IJUAhsPin1", "Pin1Diameter")).PropValue;
                double clampThickness = (double)((PropertyValueDouble)ductClampPart.GetPropertyValue("IJUAhsStrap", "StrapThickness")).PropValue;
                double[] lugOffset = new double[3];
                lugOffset[2] = lugOffset[1] = D ;

                componentDictionary[DUCTCLAMP].SetPropertyValue(boundingBoxWidth, "IJOAhsPipeOD", "PipeOD");
                componentDictionary[DUCTCLAMP].SetPropertyValue(boundingBoxHeight + G, "IJUAhsStrap", "StrapHeightInside");
                componentDictionary[DUCTCLAMP].SetPropertyValue(boundingBoxWidth, "IJUAhsStrap", "StrapWidthInside");
                componentDictionary[DUCTCLAMP].SetPropertyValue(boundingBoxWidth + 4 * (boltDia + clampThickness), "IJUAhsStrap", "StrapWidthWings");
                componentDictionary[DUCTCLAMP].SetPropertyValue(boundingBoxWidth/2 + boltDia + clampThickness, "IJUAhsOffset1", "Offset1");

                DuctObjectInfo duct = (DuctObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                if (duct.CrossSectionShape == CrossSectionShape.Rectangular)
                    componentDictionary[DUCTCLAMP].SetPropertyValue(boundingBoxWidth, "IJUAhsStrap", "StrapFlatSpot");

                //set pad Locations
                double rightPadLength = 0, rightPadWidth = 0, leftPadLength = 0, leftPadWidth = 0;
                Part padPart = null;
                string rightPadPartnumber = "", leftPadPartnumber = "";
                string rightPadtype = "", leftPadtype = "";
                if (rightPad)
                {
                    padPart = (Part)componentDictionary[RIGHTPAD].GetRelationship("madeFrom", "part").TargetObjects[0];
                    rightPadPartnumber = padPart.PartNumber;
                    rightPadtype = SupDuctAssemblyServices.GetPadPartType(this, rightPadPartnumber);
                    if (rightPadtype == "Triangular")
                    {
                        rightPadLength = (double)((PropertyValueDouble)componentDictionary[RIGHTPAD].GetPropertyValue("IJUAhsLength1", "Length1")).PropValue;
                        rightPadWidth = (double)((PropertyValueDouble)componentDictionary[RIGHTPAD].GetPropertyValue("IJUAhsWidth1", "Width1")).PropValue;
                        componentDictionary[RIGHTPAD].SetPropertyValue(-rightPadLength / 6.0, "IJUAhsStandardPort1", "P1xOffset");
                        componentDictionary[RIGHTPAD].SetPropertyValue(-rightPadWidth / 6.0, "IJUAhsStandardPort1", "P1yOffset");
                        componentDictionary[RIGHTPAD].SetPropertyValue(-rightPadLength / 6.0, "IJUAhsStandardPort2", "P2xOffset");
                        componentDictionary[RIGHTPAD].SetPropertyValue(-rightPadWidth / 6.0, "IJUAhsStandardPort2", "P2yOffset");
                    }

                }
                if (leftPad)
                {
                    padPart = (Part)componentDictionary[LEFTPAD].GetRelationship("madeFrom", "part").TargetObjects[0];
                    leftPadPartnumber = padPart.PartNumber;
                    leftPadtype = SupDuctAssemblyServices.GetPadPartType(this, leftPadPartnumber);
                    if (leftPadtype == "Triangular")
                    {
                        leftPadLength = (double)((PropertyValueDouble)componentDictionary[LEFTPAD].GetPropertyValue("IJUAhsLength1", "Length1")).PropValue;
                        leftPadWidth = (double)((PropertyValueDouble)componentDictionary[LEFTPAD].GetPropertyValue("IJUAhsWidth1", "Width1")).PropValue;
                        componentDictionary[LEFTPAD].SetPropertyValue(-leftPadLength / 6.0, "IJUAhsStandardPort1", "P1xOffset");
                        componentDictionary[LEFTPAD].SetPropertyValue(-leftPadWidth / 6.0, "IJUAhsStandardPort1", "P1yOffset");
                        componentDictionary[LEFTPAD].SetPropertyValue(-leftPadLength / 6.0, "IJUAhsStandardPort2", "P2xOffset");
                        componentDictionary[LEFTPAD].SetPropertyValue(-leftPadWidth / 6.0, "IJUAhsStandardPort2", "P2yOffset");
                    }
                }
                //====== ======
                //4. Create Joints
                //====== ======
                string bbxLow, bbxHigh, structure, struct_2 = string.Empty;
                string highIndex, struct2Index = "-1", lowIndex, structureIndex;
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    bbxLow = "BBSR_Low";
                    bbxHigh = "BBSR_High";
                }
                else
                {
                    bbxLow = "BBR_Low";
                    bbxHigh = "BBR_High";
                }

                if (Configuration == 1)
                {
                    structureIndex = "-1";
                    structure = "Structure";
                    lowIndex = "-1";
                    highIndex = "-1";
                }
                else
                {
                    JointHelper.CreateRigidJoint("-1", "Structure", CONNECTION1, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.NegativeX, 0, 0, 0);
                    JointHelper.CreateRigidJoint("-1", bbxLow, CONNECTION2, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.NegativeX, 0, boundingBoxWidth, 0);
                    JointHelper.CreateRigidJoint("-1", bbxHigh, CONNECTION3, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.NegativeX, 0, -boundingBoxWidth, 0);

                    if (!offsetApplied[2])
                    {
                        JointHelper.CreateRigidJoint("-1", "Struct_2", CONNECTION4, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.NegativeX, 0, 0, 0);
                        struct2Index = CONNECTION4;
                        struct_2 = "Connection";
                    }
                    structureIndex = CONNECTION1;
                    structure = "Connection";
                    lowIndex = CONNECTION2;
                    highIndex = CONNECTION3;
                    bbxHigh = bbxLow = "Connection";
                }
                //---- (Flexible Structural Part)----
                JointHelper.CreatePrismaticJoint(HGRBEAM, "BeginCap", HGRBEAM, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                string topStructurePort, botStructurePort, topPartIndex, botPartIndex;
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    // =============By Structure Command ====================
                    //route ---- endcap
                    JointHelper.CreatePlanarJoint(highIndex, bbxHigh, HGRBEAM, "EndCap", Plane.ZX, Plane.NegativeYZ, -boundingBoxWidth - G);

                    if (!leftPad)
                    {
                        topStructurePort = "EndCap";
                        topPartIndex = HGRBEAM;
                    }
                    else
                    {
                        JointHelper.CreateRigidJoint(HGRBEAM, "EndCap", LEFTPAD, "Port1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, steelDepth / 2, steelWidth / 2);
                        topStructurePort = "Port2";
                        topPartIndex = LEFTPAD;
                    }
                    //==============Top Structure Connection=============
                    //endcap ---- structure
                    JointHelper.CreatePrismaticJoint(structureIndex, structure, topPartIndex, topStructurePort, Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeX, 0, -steelDepth / 2);
                    //============Bottom Structure Connection ==============
                    //begincap --- route(structure)
                    if (!rightPad)
                    {
                        botStructurePort = "BeginCap";
                        botPartIndex = HGRBEAM;
                    }
                    else
                    {
                        JointHelper.CreateRigidJoint(HGRBEAM, "BeginCap", RIGHTPAD, "Port1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, steelDepth / 2, steelWidth / 2);

                        botStructurePort = "Port2";
                        botPartIndex = RIGHTPAD;
                    }

                    if (offsetApplied[2])
                        JointHelper.CreatePlanarJoint(lowIndex, bbxLow, botPartIndex, botStructurePort, Plane.XY, Plane.XY, -lugOffset[2]);
                    else
                        //==============By Point Command======================
                        //route ---- endcap
                        JointHelper.CreatePointOnPlaneJoint(botPartIndex, botStructurePort, struct2Index, struct_2, Plane.XY);

                }
                else
                {
                    JointHelper.CreatePrismaticJoint(highIndex, bbxHigh, HGRBEAM, "EndCap", Plane.ZX, Plane.NegativeYZ, Axis.Z, Axis.Z, -boundingBoxWidth - G, -steelDepth / 2);

                    if (!leftPad)
                    {
                        topStructurePort = "EndCap";
                        topPartIndex = HGRBEAM;
                    }
                    else
                    {
                        JointHelper.CreateRigidJoint(HGRBEAM, "EndCap", LEFTPAD, "Port1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, steelDepth / 2, steelWidth / 2);

                        topStructurePort = "Port2";
                        topPartIndex = LEFTPAD;
                    }

                    JointHelper.CreatePointOnPlaneJoint(topPartIndex, topStructurePort, structureIndex, structure, Plane.XY);

                    if (!rightPad)
                    {
                        botStructurePort = "BeginCap";
                        botPartIndex = HGRBEAM;
                    }
                    else
                    {
                        JointHelper.CreateRigidJoint(HGRBEAM, "BeginCap", RIGHTPAD, "Port1", Plane.XY, Plane.XY, Axis.X, Axis.X, steelDepth / 2, steelWidth / 2, 0);

                        botStructurePort = "Port2";
                        botPartIndex = RIGHTPAD;
                    }

                    if (offsetApplied[2])
                        JointHelper.CreatePlanarJoint(lowIndex, bbxLow, botPartIndex, botStructurePort, Plane.XY, Plane.XY, -lugOffset[2]);
                    else
                        JointHelper.CreatePointOnPlaneJoint(botPartIndex, botStructurePort, struct2Index, struct_2, Plane.XY);
                }
                //---- (Profile part -- Duct Clamp -- Duct) ----
                if(Configuration == 1)
                    JointHelper.CreateRigidJoint(DUCTCLAMP, "Route", "-1", "Route", Plane.ZX, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                else
                    JointHelper.CreateRigidJoint(DUCTCLAMP, "Route", "-1", "Route", Plane.ZX, Plane.NegativeXY, Axis.X, Axis.X, 0, 0, 0);
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

                    routeConnections.Add(new ConnectionInfo(DUCTCLAMP, 1)); // partindex, routeindex

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
                    //Create a collection to hold the ALL Structure connection information
                    Collection<ConnectionInfo> structConnections = new Collection<ConnectionInfo>();
                    if (leftPad)
                        structConnections.Add(new ConnectionInfo(LEFTPAD, 1)); // partindex, structureindex
                    else
                        structConnections.Add(new ConnectionInfo(HGRBEAM, 1));
                    if (rightPad != leftPad)
                    {
                        if (rightPad)
                            structConnections.Add(new ConnectionInfo(RIGHTPAD, 1)); // partindex, structureindex
                        else
                            structConnections.Add(new ConnectionInfo(HGRBEAM, 1));
                    }
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

