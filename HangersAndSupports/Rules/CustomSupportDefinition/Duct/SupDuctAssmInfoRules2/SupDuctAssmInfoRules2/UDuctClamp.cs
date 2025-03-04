//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   UDuctClamp.cs
//   SupDuctAssmInfoRules2,Ingr.SP3D.Content.Support.Rules.UDuctClamp
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
    public class UDuctClamp : CustomSupportDefinition
    {
        private const string HGRBEAM1 = "HGRBEAM1"; //Hanger beam object
        private const string HGRBEAM2 = "HGRBEAM2"; //Hanger beam object
        private const string HGRBEAM3 = "HGRBEAM3"; //Hanger beam object
        private const string DUCTCLAMP = "DUCTCLAMP";          //Hanger beam object as angle guide
        private const string LEFTPAD = "LEFTPAD";
        private const string RIGHTPAD = "RIGHTPAD";
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
                    //=============================
                    //Retrieve the support occurrence data (attributes)
                    //All structural assemblies will share the same attribute set.
                    //=============================
                    object[] AttributeColl = SupDuctAssemblyServices.GetDuctStructuralASMAttributes(this);
                    D = (double)AttributeColl[0];
                    G = (double)AttributeColl[1];
                    leftPad = (bool)AttributeColl[2];
                    rightPad = (bool)AttributeColl[3];

                    string sectionPartClass = (string)((PropertyValueString)support.SupportDefinition.GetPropertyValue("IJUAHSA_SecPartClass", "SecPartClass")).PropValue;

                    parts.Add(new PartInfo(HGRBEAM1, sectionPartClass, "PartSelectionRule,Ingr.SP3D.Content.Support.Rules.CPartByCrossSection"));
                    parts.Add(new PartInfo(HGRBEAM2, sectionPartClass, "PartSelectionRule,Ingr.SP3D.Content.Support.Rules.CPartByCrossSection"));
                    parts.Add(new PartInfo(HGRBEAM3, sectionPartClass, "PartSelectionRule,Ingr.SP3D.Content.Support.Rules.CPartByCrossSection"));
                    DuctObjectInfo duct = (DuctObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                    if (duct.CrossSectionShape == CrossSectionShape.Rectangular)
                        parts.Add(new PartInfo(DUCTCLAMP,"S3Dhs_DuctClamp_1"));
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
                return 1;
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

                double boundingBoxWidth = boundingBox.Width, boundingBoxHeight = boundingBox.Height;

                //====================================
                //3. Retrieve route/structure configuration and
                // check if offset value is to be applied on lug end
                //If the leg is attached to a stiffener, no offset needs to be specified
                //If the leg is attached to a slab, the offest value will be retrieved from
                //predefined offset rule
                //====================================
                bool[] bIsOffsetApplied = SupDuctAssemblyServices.GetIsLugEndOffsetApplied(this);
                // get cross section dimensions
                BusinessObject horizontalSectionPart = componentDictionary[HGRBEAM3].GetRelationship("madeFrom", "part").TargetObjects[0];
                CrossSection crosssection = (CrossSection)horizontalSectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];

                //Get SteelWidth and SteelDepth 
                double steelWidth = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructCrossSectionDimensions", "Width")).PropValue;
                double steelDepth = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructCrossSectionDimensions", "Depth")).PropValue;
                BusinessObject ductClampPart = componentDictionary[DUCTCLAMP].GetRelationship("madeFrom", "part").TargetObjects[0];
                double boltDia = (double)((PropertyValueDouble)ductClampPart.GetPropertyValue("IJUAhsPin1", "Pin1Diameter")).PropValue;
                double clampThickness = (double)((PropertyValueDouble)ductClampPart.GetPropertyValue("IJUAhsStrap", "StrapThickness")).PropValue;


                double[] lugOffset = new double[3];
                lugOffset[2] = lugOffset[1] = D;

                componentDictionary[DUCTCLAMP].SetPropertyValue(boundingBoxWidth, "IJOAhsPipeOD", "PipeOD");
                componentDictionary[DUCTCLAMP].SetPropertyValue(boundingBoxHeight + G, "IJUAhsStrap", "StrapHeightInside");
                componentDictionary[DUCTCLAMP].SetPropertyValue(boundingBoxWidth, "IJUAhsStrap", "StrapWidthInside");
                componentDictionary[DUCTCLAMP].SetPropertyValue(boundingBoxWidth + 4 * (boltDia + clampThickness), "IJUAhsStrap", "StrapWidthWings");
                componentDictionary[DUCTCLAMP].SetPropertyValue(boundingBoxWidth/2 + boltDia + clampThickness, "IJUAhsOffset1", "Offset1");

                DuctObjectInfo duct = (DuctObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                if (duct.CrossSectionShape == CrossSectionShape.Rectangular)
                    componentDictionary[DUCTCLAMP].SetPropertyValue(boundingBoxWidth, "IJUAhsStrap", "StrapFlatSpot");

                //====================
                //4.  Get Indexed structural reference ports
                //This method will differentiate "Left" structure input from "Right"
                //structur input.
                //"Left" structure is always on the negtive Y direction of the bounding box coord. sys
                //"Right" structure is on the positive Y direction of the bounding box coord. sys.
                //'====================
                string[] idxstructPort = SupDuctAssemblyServices.GetIndexedStructPortName(this, bIsOffsetApplied);
                string leftStructPort = idxstructPort[0];
                string rightStructPort = idxstructPort[1];

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
                //----------------------------------------------------
                //Create the Joint between the RteLow Reference Port and the Bottom Symbol
                //Bounding box --- Horizontal Beam
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    //Beam Structure... Add a Prismatic Joint
                    JointHelper.CreatePrismaticJoint("-1", "BBSR_Low", HGRBEAM2, "BeginCap", Plane.XY, Plane.YZ, Axis.X, Axis.Y, boundingBoxHeight + G, -lugOffset[1]);
                else
                {
                    if (bIsOffsetApplied[0])
                        //Add a Rigid Joint
                        JointHelper.CreateRigidJoint("-1", "BBR_Low", HGRBEAM2, "BeginCap", Plane.XY, Plane.YZ, Axis.X, Axis.Y, boundingBoxHeight + G, -lugOffset[1], -steelDepth/2);
                    else
                        //Add a prismatic joint
                        JointHelper.CreatePrismaticJoint("-1", "BBR_Low", HGRBEAM2, "BeginCap", Plane.XY, Plane.YZ, Axis.Y, Axis.Z, boundingBoxHeight + G, 0);
                }
                //----------------------------------------------------
                //Create the Plane Joint between the RteHigh Reference Port
                //and the Right Bottom Symbol
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    JointHelper.CreatePlanarJoint("-1", "BBSR_High", HGRBEAM2, "EndCap", Plane.ZX, Plane.XY, lugOffset[2]);
                else
                {
                    if (bIsOffsetApplied[1])
                        //Add a planar joint is offset is to be applied
                        JointHelper.CreatePlanarJoint("-1", "BBR_High", HGRBEAM2, "EndCap", Plane.ZX, Plane.XY, lugOffset[2]);
                }
                //Add a Prismatic Joint defining the flexible bottom member
                JointHelper.CreatePrismaticJoint(HGRBEAM2, "BeginCap", HGRBEAM2, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                //Add rigid joint between beam 1 and beam 2
                JointHelper.CreateRigidJoint(HGRBEAM1, "EndCap", HGRBEAM2, "BeginCap", Plane.XY, Plane.NegativeYZ, Axis.X, Axis.Z, 0, 0, 0);
                //Add rigid joint between beam 2 and beam 3
                JointHelper.CreateRigidJoint(HGRBEAM3, "BeginCap", HGRBEAM2, "EndCap", Plane.XY, Plane.YZ, Axis.X, Axis.NegativeZ, 0, 0, 0);
                //Add prismatic joint for the flexible memeber
                JointHelper.CreatePrismaticJoint(HGRBEAM1, "BeginCap", HGRBEAM1, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                //Add prismatic joint for the flexible memeber
                JointHelper.CreatePrismaticJoint(HGRBEAM3, "BeginCap", HGRBEAM3, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                //structure -- leftPad -- Profile part 1
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    if (leftPad)
                    {
                        JointHelper.CreatePrismaticJoint("-1", leftStructPort, LEFTPAD, "Port1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0);
                        JointHelper.CreateRigidJoint(LEFTPAD, "Port2", HGRBEAM1, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, -steelDepth / 2, -steelWidth / 2);
                    }
                    else
                        //add a prismatic joint
                        JointHelper.CreatePrismaticJoint("-1", leftStructPort, HGRBEAM1, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, -steelDepth / 2);
                }
                else
                {
                    if (leftPad)
                    {
                        if (bIsOffsetApplied[0])
                            JointHelper.CreatePlanarJoint("-1", leftStructPort, LEFTPAD, "Port1", Plane.XY, Plane.XY, 0);
                        else
                            JointHelper.CreatePlanarSlotJoint("-1", leftStructPort, LEFTPAD, "Port1", Plane.XY, Plane.XY, Axis.X, 0, 0);

                        JointHelper.CreateRigidJoint(LEFTPAD, "Port2", HGRBEAM1, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, -steelDepth / 2, -steelWidth / 2);
                    }
                    else
                    {
                        if (bIsOffsetApplied[0])
                            JointHelper.CreatePlanarJoint("-1", leftStructPort, HGRBEAM1, "BeginCap", Plane.XY, Plane.XY, 0);
                        else
                            //Add a planar joint
                            JointHelper.CreatePlanarSlotJoint("-1", leftStructPort, HGRBEAM1, "BeginCap", Plane.XY, Plane.XY, Axis.X, 0, 0);
                    }
                }
                //structure -- rightPad -- profile part 3
                if (rightPad)
                {
                    if (bIsOffsetApplied[1])
                        JointHelper.CreatePlanarJoint("-1", rightStructPort, RIGHTPAD, "Port2", Plane.XY, Plane.NegativeXY, 0);
                    else
                        JointHelper.CreatePlanarSlotJoint("-1", rightStructPort, RIGHTPAD, "Port2", Plane.XY, Plane.NegativeXY, Axis.X, 0, 0); ;

                    JointHelper.CreateRigidJoint(RIGHTPAD, "Port1", HGRBEAM3, "EndCap", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, -steelDepth / 2, -steelWidth / 2);
                }
                else
                {
                    if (bIsOffsetApplied[1])
                        JointHelper.CreatePlanarJoint("-1", rightStructPort, HGRBEAM3, "EndCap", Plane.XY, Plane.NegativeXY, 0);
                    else
                        JointHelper.CreatePlanarSlotJoint("-1", rightStructPort, HGRBEAM3, "EndCap", Plane.XY, Plane.NegativeXY, Axis.X, 0, 0);
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
                    if (leftPad == true)
                        structConnections.Add(new ConnectionInfo(LEFTPAD, 1)); // LEFTPAD, structureindex
                    else
                        structConnections.Add(new ConnectionInfo(HGRBEAM1, 1));// HgrBeam, structureindex

                    if (rightPad == true)
                        structConnections.Add(new ConnectionInfo(RIGHTPAD, 1)); // RIGHTPAD, structureindex
                    else
                        structConnections.Add(new ConnectionInfo(HGRBEAM2, 1));// HgrBeam, structureindex

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

