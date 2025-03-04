//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   AngularClamp.cs
//   SupDuctAssmInfoRules2,Ingr.SP3D.Content.Support.Rules.AngularClamp
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
    public class AngularClamp : CustomSupportDefinition
    {
        //Constants
        private const string HGRBEAM1 = "HgrBeam_1"; //Horizontal structural part
        private const string CONNECTION1 = "Connection_1";    //Logical connection
        private const string HGRBEAM2 = "HgrBeam_2";          //supporting strucutral part
        private const string CONNECTION2 = "Connection_2";    //Logical connection
        private const string DuctClamp = "DuctClamp";      //Duct clamp
        private const string CONNECTION3 = "Connection_3";   //bounds to structure, make it easy for toggle
        private const string CONNECTION4 = "Connection_4";   //bounds to BBSR_Low/BBR_Low, make it easy for toggle
        private const string CONNECTION5 = "Connection_5";   //bounds to BBSR_High/BBR_High, make it easy for toggle
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
                    object[] AttributeColl = SupDuctAssemblyServices.GetDuctStructuralASMAttributes(this);
                    D = (double)AttributeColl[0];
                    G = (double)AttributeColl[1];
                    leftPad = (bool)AttributeColl[2];
                    rightPad = (bool)AttributeColl[3];
                    string sectionPartClass = (string)((PropertyValueString)support.SupportDefinition.GetPropertyValue("IJUAHSA_SecPartClass", "SecPartClass")).PropValue;

                    parts.Add(new PartInfo(HGRBEAM1, sectionPartClass, "PartSelectionRule,Ingr.SP3D.Content.Support.Rules.CPartByCrossSection"));
                    parts.Add(new PartInfo(HGRBEAM2, sectionPartClass, "PartSelectionRule,Ingr.SP3D.Content.Support.Rules.CPartByCrossSection"));
                    parts.Add(new PartInfo(CONNECTION1, "Connection"));
                    parts.Add(new PartInfo(CONNECTION2, "Connection"));
                    parts.Add(new PartInfo(CONNECTION3, "Connection"));
                    parts.Add(new PartInfo(CONNECTION4, "Connection"));
                    parts.Add(new PartInfo(CONNECTION5, "Connection"));

                    DuctObjectInfo duct = (DuctObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                    if (duct.CrossSectionShape == CrossSectionShape.Rectangular)
                        parts.Add(new PartInfo(DuctClamp, "S3Dhs_DuctClamp_1"));
                    else
                        parts.Add(new PartInfo(DuctClamp, "S3Dhs_DuctClamp_2"));

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

                double boundingBoxWidth = boundingBox.Width, boundingBoxHeight = boundingBox.Height;

                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;

                BusinessObject sectionPart = componentDictionary[HGRBEAM2].GetRelationship("madeFrom", "part").TargetObjects[0];
                CrossSection crossSection = (CrossSection)sectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];

                double sectionWidth = crossSection.Width, sectionDepth = crossSection.Depth;

                Part part = (Part)componentDictionary[DuctClamp].GetRelationship("madeFrom", "part").TargetObjects[0];
                double boltDia = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAhsPin1", "Pin1Diameter")).PropValue;
                double clampThickness = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAhsStrap", "StrapThickness")).PropValue;
                double lugOffset = D;

                componentDictionary[DuctClamp].SetPropertyValue(boundingBoxWidth, "IJOAhsPipeOD", "PipeOD");
                componentDictionary[DuctClamp].SetPropertyValue(boundingBoxHeight + G, "IJUAhsStrap", "StrapHeightInside");
                componentDictionary[DuctClamp].SetPropertyValue(boundingBoxWidth, "IJUAhsStrap", "StrapWidthInside");
                componentDictionary[DuctClamp].SetPropertyValue(boundingBoxWidth + 4 * (boltDia + clampThickness), "IJUAhsStrap", "StrapWidthWings");
                componentDictionary[DuctClamp].SetPropertyValue(boundingBoxWidth/2 + boltDia + clampThickness, "IJUAhsOffset1", "Offset1");

                DuctObjectInfo duct = (DuctObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                if (duct.CrossSectionShape == CrossSectionShape.Rectangular)
                    componentDictionary[DuctClamp].SetPropertyValue(boundingBoxWidth, "IJUAhsStrap", "StrapFlatSpot");

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
                //3. Create Joints
                //====== ======
                //Create a collection to hold the joints

                string bbxLow = string.Empty;
                string bbxHigh = string.Empty;
                string structure = string.Empty;
                string structureIndex = string.Empty;
                string lowIndex = string.Empty;
                string highIndex = string.Empty;

                //Make it easy for toggle
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
                    JointHelper.CreatePrismaticJoint("-1", "Structure", CONNECTION3, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.NegativeX, 0, 0);
                    JointHelper.CreateRigidJoint("-1", bbxLow, CONNECTION4, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.NegativeX, 0, boundingBoxWidth, 0);
                    JointHelper.CreateRigidJoint("-1", bbxHigh, CONNECTION5, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.NegativeX, 0, -boundingBoxWidth, 0);

                    structure = "Connection";
                    structureIndex = CONNECTION3;
                    bbxLow = "Connection";
                    bbxHigh = "Connection";
                    lowIndex = CONNECTION4;
                    highIndex = CONNECTION5;
                }
                else
                {
                    structure = "Structure";
                    structureIndex = "-1";
                    lowIndex = "-1";
                    highIndex = "-1";
                }

                //-----------------------------------------------------------------------------
                //---------Structure <--> (Pad) <--> Profile part <--> Logical Connection-----

                //Add a prismatic joint between structure and part 1

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    if (leftPad)
                    {
                        JointHelper.CreatePrismaticJoint(structureIndex, structure, LEFTPAD, "Port1", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0);
                        JointHelper.CreateRigidJoint(LEFTPAD, "Port2", HGRBEAM1, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, -sectionDepth / 2.0, -sectionWidth / 2.0);
                    }
                    else
                        JointHelper.CreatePrismaticJoint(structureIndex, structure, HGRBEAM1, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, sectionDepth / 2.0);
                }
                else
                {
                    if (leftPad)
                    {
                        JointHelper.CreatePlanarJoint(structureIndex, structure, LEFTPAD, "Port1", Plane.XY, Plane.XY, 0);
                        JointHelper.CreateRigidJoint(LEFTPAD, "Port2", HGRBEAM1, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, -sectionDepth / 2.0, -sectionWidth / 2.0);
                    }
                    else
                        JointHelper.CreatePlanarJoint(structureIndex, structure, HGRBEAM1, "BeginCap", Plane.XY, Plane.XY, 0);
                }

                double offset = D;

                //----------------------------------------
                //Route --- HgrBeam 1
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    JointHelper.CreatePlanarJoint(lowIndex, bbxLow, HGRBEAM1, "EndCap", Plane.XY, Plane.NegativeXY, -offset);
                    JointHelper.CreatePlanarJoint(highIndex, bbxHigh, HGRBEAM1, "EndCap", Plane.ZX, Plane.ZX, G);
                }
                else
                    JointHelper.CreateRigidJoint(highIndex, bbxHigh, HGRBEAM1, "EndCap", Plane.XY, Plane.NegativeXY, Axis.Y, Axis.Y, -offset-boundingBoxHeight, sectionWidth/2, G);

                //Flexible Member
                JointHelper.CreatePrismaticJoint(HGRBEAM1, "BeginCap", HGRBEAM1, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                //-------------------------------------------------------------------------
                //------structure <--> Logical connection <--> Pad <--> profile part 2-----

                if (leftPad)
                    JointHelper.CreateRigidJoint(CONNECTION1, "Connection", LEFTPAD, "Port1", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, -lugOffset * 5.0);
                else
                    JointHelper.CreateRigidJoint(CONNECTION1, "Connection", HGRBEAM1, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, -lugOffset * 5.0);

                //Logical conection between horizontal beam and logical part to locate the second beam
                JointHelper.CreateRigidJoint(HGRBEAM1, "EndCap", CONNECTION2, "Connection", Plane.YZ, Plane.YZ, Axis.Z, Axis.Z, 0, 0, -offset - boundingBoxHeight);

                //Cylindrical joint between logical connection and beam 2
                if (rightPad)
                {
                    JointHelper.CreateRigidJoint(CONNECTION1, "Connection", RIGHTPAD, "Port1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    JointHelper.CreateCylindricalJoint(RIGHTPAD, "Port2", HGRBEAM2, "BeginCap", Axis.Y, Axis.NegativeX, 0);
                }
                else
                    JointHelper.CreateCylindricalJoint(CONNECTION1, "Connection", HGRBEAM2, "BeginCap", Axis.Y, Axis.NegativeX, 0);

                //Spherical joint between logical connection and beam 2
                JointHelper.CreateSphericalJoint(CONNECTION2, "Connection", HGRBEAM2, "EndCap");

                //Flexible member
                JointHelper.CreatePrismaticJoint(HGRBEAM2, "BeginCap", HGRBEAM2, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                //--------------------------------------------------
                //------Profile part 1 <--> Duct Clamp <--> Duct
                if (Configuration == 1)
                    JointHelper.CreateRigidJoint(DuctClamp, "Route", "-1", "Route", Plane.ZX, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                else
                    JointHelper.CreateRigidJoint(DuctClamp, "Route", "-1", "Route", Plane.ZX, Plane.NegativeXY, Axis.X, Axis.X, 0, 0, 0);
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

                    routeConnections.Add(new ConnectionInfo(DuctClamp, 1)); // partindex, routeindex

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
                        structConnections.Add(new ConnectionInfo(HGRBEAM1, 1));// partindex, structureindex    
                    if (rightPad)
                        structConnections.Add(new ConnectionInfo(RIGHTPAD, 1)); // partindex, structureindex    
                    else
                        structConnections.Add(new ConnectionInfo(CONNECTION2, 1));// partindex, structureindex    

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



