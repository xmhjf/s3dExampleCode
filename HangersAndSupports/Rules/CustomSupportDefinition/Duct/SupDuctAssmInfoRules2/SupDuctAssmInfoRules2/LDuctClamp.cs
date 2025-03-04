//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   LDuctClamp.cs
//   SupDuctAssmInfoRules2,Ingr.SP3D.Content.Support.Rules.LDuctClamp
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
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.ReferenceData.Middle;

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
    public class LDuctClamp : CustomSupportDefinition
    {
        //Constants
        private const string HGRBEAM1 = "HGRBEAM1"; //Hanger beam object
        private const string HGRBEAM2 = "HGRBEAM2"; //Hanger beam object
        private const string DUCTCLAMP = "DUCTCLAMP";          //Hanger beam object as angle guide
        private const string LEFTPAD = "LEFTPAD";
        private const string RIGHTPAD = "RIGHTPAD";
        double D, G;
        bool leftPad, rightPad;

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

                    //Get the attributes from assembly
                    object[] attributeCollection = SupDuctAssemblyServices.GetDuctStructuralASMAttributes(this);
                    D = (double)attributeCollection[0];
                    G = (double)attributeCollection[1];
                    leftPad = (bool)attributeCollection[2];
                    rightPad = (bool)attributeCollection[3];

                    string sectionPartClass = (string)((PropertyValueString)support.SupportDefinition.GetPropertyValue("IJUAHSA_SecPartClass", "SecPartClass")).PropValue;

                    parts.Add(new PartInfo(HGRBEAM1, sectionPartClass, "PartSelectionRule,Ingr.SP3D.Content.Support.Rules.CPartByCrossSection"));
                    parts.Add(new PartInfo(HGRBEAM2, sectionPartClass, "PartSelectionRule,Ingr.SP3D.Content.Support.Rules.CPartByCrossSection"));
                    DuctObjectInfo duct = (DuctObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                    if (duct.CrossSectionShape == CrossSectionShape.Rectangular)
                        parts.Add(new PartInfo(DUCTCLAMP, "S3Dhs_DuctClamp_1"));
                    else
                        parts.Add(new PartInfo(DUCTCLAMP, "S3Dhs_DuctClamp_2"));

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
                //load standard bounding box.
                BoundingBoxHelper.CreateStandardBoundingBoxes(false);
                BoundingBox boundingBox;

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    boundingBox = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.SupportedAndSupporting);
                else
                    boundingBox = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.Supported);

                double boundingBoxWidth = boundingBox.Width, boundingBoxHeight = boundingBox.Height;

                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;

                double[] lugOffset = new double[2];
                bool[] isOffsetApplied = new bool[2];
                int routeStructureConfiguration;
                if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                {
                    routeStructureConfiguration = -1;
                    isOffsetApplied[0] = true;
                    isOffsetApplied[1] = true;     //For Place By reference
                }
                else
                    routeStructureConfiguration = SupDuctAssemblyServices.GetDuctLSupportConfiguration(this, isOffsetApplied);

                if (routeStructureConfiguration == -1)
                    routeStructureConfiguration = Configuration;

                //===========================
                //7. Get the input objects information
                //   Dimensions of the pipe/duct cross section
                //   (Applicable for Pipe/Duct/Cableway)
                //==========================

                //===================================
                //8. Set symbol occurrence attributes (optional)
                //===================================
                double pcsSectionWidth = 0, pcsSectionDepth = 0;

                BusinessObject sectionPart = componentDictionary[HGRBEAM2].GetRelationship("madeFrom", "part").TargetObjects[0];
                CrossSection crossSection = (CrossSection)sectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];

                pcsSectionWidth = crossSection.Width;
                pcsSectionDepth = crossSection.Depth;

                BusinessObject ductClampPart = componentDictionary[DUCTCLAMP].GetRelationship("madeFrom", "part").TargetObjects[0];
                double boltDia = (double)((PropertyValueDouble)ductClampPart.GetPropertyValue("IJUAhsPin1", "Pin1Diameter")).PropValue;
                double clampThickness = (double)((PropertyValueDouble)ductClampPart.GetPropertyValue("IJUAhsStrap", "StrapThickness")).PropValue;
                lugOffset[0] = D;
                lugOffset[1] = lugOffset[0];

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
                //10. Create Joints
                //====== ======
                //Create a collection to hold the joints

                //Add a Prismatic Joint defining the flexible bottom member
                JointHelper.CreatePrismaticJoint(HGRBEAM2, "BeginCap", HGRBEAM2, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                //Add prismatic joint for the flexible memeber
                JointHelper.CreatePrismaticJoint(HGRBEAM1, "BeginCap", HGRBEAM1, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                string bottomStructPort = string.Empty, bottomPartIndex = string.Empty, topStructPort = string.Empty, topPartIndex;
                double axisOffset;
                if (routeStructureConfiguration == 1)
                {
                    //Create the Joint between the RteLow Reference Port and the Left Bottom Symbol
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        //Beam Structure... Add a Prismatic Joint
                        JointHelper.CreatePrismaticJoint("-1", "BBSR_High", HGRBEAM2, "EndCap", Plane.XY, Plane.YZ, Axis.X, Axis.Y, G, lugOffset[0]);
                    else
                    {
                        //Plate Structure... Add a Rigid Joint
                        if (isOffsetApplied[0])
                            JointHelper.CreateRigidJoint("-1", "BBR_High", HGRBEAM2, "EndCap", Plane.XY, Plane.YZ, Axis.X, Axis.Y, G, lugOffset[0], -pcsSectionDepth/2);
                        else
                            JointHelper.CreatePrismaticJoint("-1", "BBR_High", HGRBEAM2, "EndCap", Plane.XY, Plane.YZ, Axis.Y, Axis.Z, 0.0, 0.0);
                    }

                    //----------------------------------------------------
                    //Create the Plane Joint between the RteHigh Reference Port
                    //and the Right Bottom Symbol

                    if (rightPad)
                    {
                        JointHelper.CreatePlanarJoint(HGRBEAM2, "BeginCap", RIGHTPAD, "Port2", Plane.XY, Plane.XY, 0);
                        JointHelper.CreatePrismaticJoint(RIGHTPAD, "Port2", HGRBEAM2, "Neutral", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                        bottomStructPort = "Port1";
                        bottomPartIndex = RIGHTPAD;
                    }
                    else
                    {
                        bottomStructPort = "BeginCap";
                        bottomPartIndex = HGRBEAM2;
                    }

                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    {
                        if (isOffsetApplied[1])
                            JointHelper.CreatePlanarJoint("-1", "BBSR_Low", HGRBEAM2, "BeginCap", Plane.ZX, Plane.XY, -lugOffset[1]);
                        else
                            JointHelper.CreatePointOnPlaneJoint(bottomPartIndex, bottomStructPort, "-1", "Struct_2", Plane.XY);
                    }
                    else
                    {
                        if (isOffsetApplied[1])
                            JointHelper.CreatePlanarJoint("-1", "BBR_Low", HGRBEAM2, "BeginCap", Plane.ZX, Plane.XY, -lugOffset[1]);
                        else
                            JointHelper.CreatePointOnPlaneJoint(bottomPartIndex, bottomStructPort, "-1", "Struct_2", Plane.XY);
                    }
                    //----------------------------------------------------
                    //Add rigid joint between beam 1 and beam 2
                    JointHelper.CreateRigidJoint(HGRBEAM2, "EndCap", HGRBEAM1, "BeginCap", Plane.XY, Plane.NegativeYZ, Axis.X, Axis.Z, 0, 0, 0);

                    //----------------------------------------------------
                    //Add prismatic joint between beam 1 and structure

                    if (leftPad)
                    {
                        JointHelper.CreatePlanarJoint(HGRBEAM1, "EndCap", LEFTPAD, "Port1", Plane.XY, Plane.XY, 0);
                        JointHelper.CreatePrismaticJoint(LEFTPAD, "Port1", HGRBEAM1, "Neutral", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                        topStructPort = "Port2";
                        topPartIndex = LEFTPAD;
                        axisOffset = 0;
                    }
                    else
                    {
                        topStructPort = "EndCap";
                        topPartIndex = HGRBEAM1;
                        axisOffset = -pcsSectionWidth / 2.0;
                    }

                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        JointHelper.CreatePrismaticJoint("-1", "Structure", topPartIndex, topStructPort, Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeX, 0.0, axisOffset);
                    else
                    {
                        if (isOffsetApplied[0])
                            JointHelper.CreatePlanarJoint("-1", "Structure", topPartIndex, topStructPort, Plane.XY, Plane.NegativeXY, 0);
                        else
                            JointHelper.CreatePlanarSlotJoint("-1", "Structure", topPartIndex, topStructPort, Plane.XY, Plane.NegativeXY, Axis.X, 0, 0);
                    }
                }
                else
                {
                    //----------------------------------------------------
                    //Create the Joint between the RteLow Reference Port and the Left Bottom Symbol
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        //Beam Structure... Add a Prismatic Joint
                        JointHelper.CreatePrismaticJoint("-1", "BBSR_Low", HGRBEAM2, "BeginCap", Plane.XY, Plane.YZ, Axis.X, Axis.Y, boundingBoxHeight + G, -lugOffset[0]);
                    else
                    {
                        //Plate Structure... Add a Rigid Joint
                        if (isOffsetApplied[0])
                            JointHelper.CreateRigidJoint("-1", "BBR_Low", HGRBEAM2, "BeginCap", Plane.XY, Plane.YZ, Axis.X, Axis.Y, boundingBoxHeight + G, -lugOffset[0], -pcsSectionDepth/2);
                        else
                            JointHelper.CreatePrismaticJoint("-1", "BBR_Low", HGRBEAM2, "BeginCap", Plane.XY, Plane.YZ, Axis.Y, Axis.Z, boundingBoxHeight, 0);
                    }

                    //----------------------------------------------------
                    //Create the Plane Joint between the RteHigh Reference Port
                    //and the Right Bottom Symbol
                    if (rightPad)
                    {
                        JointHelper.CreatePlanarJoint(HGRBEAM2, "EndCap", RIGHTPAD, "Port1", Plane.XY, Plane.XY, 0);
                        JointHelper.CreatePrismaticJoint(RIGHTPAD, "Port1", HGRBEAM2, "Neutral", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                        bottomStructPort = "Port2";
                        bottomPartIndex = RIGHTPAD; ;
                    }
                    else
                    {
                        bottomStructPort = "EndCap";
                        bottomPartIndex = HGRBEAM2;
                    }
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    {
                        if (isOffsetApplied[1])
                            JointHelper.CreatePlanarJoint("-1", "BBSR_High", HGRBEAM2, "EndCap", Plane.ZX, Plane.XY, lugOffset[1]);
                        else
                            JointHelper.CreatePointOnPlaneJoint(bottomPartIndex, bottomStructPort, "-1", "Struct_2", Plane.XY);
                    }
                    else
                    {
                        if (isOffsetApplied[1])
                            JointHelper.CreatePlanarJoint("-1", "BBR_High", HGRBEAM2, "EndCap", Plane.ZX, Plane.XY, lugOffset[1]);
                        else
                            JointHelper.CreatePointOnPlaneJoint(bottomPartIndex, bottomStructPort, "-1", "Struct_2", Plane.XY);
                    }
                    //----------------------------------------------------
                    //Add rigid joint between beam 1 and beam 2
                    JointHelper.CreateRigidJoint(HGRBEAM1, "EndCap", HGRBEAM2, "BeginCap", Plane.XY, Plane.NegativeYZ, Axis.X, Axis.Z, 0, 0, 0);

                    //----------------------------------------------------
                    //Add prismatic joint between beam 1 and structure
                    if (leftPad)
                    {
                        JointHelper.CreatePlanarJoint(HGRBEAM1, "BeginCap", LEFTPAD, "Port1", Plane.XY, Plane.XY, 0);
                        JointHelper.CreatePrismaticJoint(LEFTPAD, "Port1", HGRBEAM1, "Neutral", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                        topStructPort = "Port2";
                        topPartIndex = LEFTPAD;
                        axisOffset = 0;
                    }
                    else
                    {
                        topStructPort = "BeginCap";
                        topPartIndex = HGRBEAM1;
                        axisOffset = -pcsSectionWidth / 2.0;
                    }
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        JointHelper.CreatePrismaticJoint("-1", "Structure", topPartIndex, topStructPort, Plane.XY, Plane.XY, Axis.X, Axis.X, 0.0, axisOffset);
                    else
                    {
                        if (isOffsetApplied[0])
                            JointHelper.CreatePlanarJoint("-1", "Structure", topPartIndex, topStructPort, Plane.XY, Plane.XY, 0);
                        else
                            JointHelper.CreatePlanarSlotJoint("-1", "Structure", topPartIndex, topStructPort, Plane.XY, Plane.XY, Axis.X, 0, 0);
                    }
                }
                //---- (Profile part -- Duct Clamp -- Duct) ----
                JointHelper.CreateRigidJoint(DUCTCLAMP, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
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

                    routeConnections.Add(new ConnectionInfo(DUCTCLAMP, 1)); // Duct Clamp, routeindex

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
                        structConnections.Add(new ConnectionInfo(LEFTPAD, 1)); // LEFTPAD, Structureindex
                    else
                        structConnections.Add(new ConnectionInfo(HGRBEAM1, 1));// HgrBeam, Structureindex

                    if (rightPad)
                        structConnections.Add(new ConnectionInfo(RIGHTPAD, 1)); // RIGHTPAD, Structureindex
                    else
                        structConnections.Add(new ConnectionInfo(HGRBEAM2, 1));// HgrBeam, Structureindex

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





