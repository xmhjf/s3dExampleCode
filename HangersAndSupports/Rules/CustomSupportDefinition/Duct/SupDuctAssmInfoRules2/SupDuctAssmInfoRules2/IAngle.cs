//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   IAngle.cs
//   SupDuctAssmInfoRules2,Ingr.SP3D.Content.Support.Rules.IAngle
//   Author       : Vinay
//   Creation Date:  27/11/2015
//   Description:    

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------

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
    public class IAngle : CustomSupportDefinition
    {
        //Constants
        private const string HGRBEAM1 = "HgrBeam_1";          //Hanger beam object
        private const string CONNECTION1 = "Connection_1";    //bounds to structure, make it easy for toggle
        private const string HGRBEAM2 = "HgrBeam_2";          //Hanger beam object as angle guide
        private const string CONNECTION2 = "Connection_2";    //bounds to BBSR_Low/BBR_Low, make it easy for toggle
        private const string HGRBEAM3 = "HgrBeam_3";          //Hanger beam object as angle guide
        private const string CONNECTION3 = "Connection_3";    //bounds to BBSR_High/BBR_High, make it easy for toggle
        private const string CONNECTION4 = "Connection_4";    //bounds to structure 2 if applicable, make it easy for toggle
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
                    parts.Add(new PartInfo(HGRBEAM2, sectionPartClass, "HgrSupSecondaryCS"));
                    parts.Add(new PartInfo(HGRBEAM3, sectionPartClass, "HgrSupSecondaryCS"));
                    parts.Add(new PartInfo(CONNECTION1, "Connection"));
                    parts.Add(new PartInfo(CONNECTION2, "Connection"));
                    parts.Add(new PartInfo(CONNECTION3, "Connection"));
                    parts.Add(new PartInfo(CONNECTION4, "Connection"));

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

                double[] lugOffset = new double[2];
                bool[] isOffsetApplied = new bool[2];

                if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                {
                    isOffsetApplied[0] = false;
                    isOffsetApplied[1] = true;     //For Place By reference
                }
                else
                {
                    if (SupportHelper.SupportingObjects.Count > 1)
                        isOffsetApplied[1] = false;
                    else
                        isOffsetApplied[1] = true;
                }
                //===========================
                //   Get the input objects information
                //   Dimensions of the pipe/duct cross section
                //   (Applicable for Pipe/Duct/Cableway)
                //==========================

                //===================================
                //3. Set symbol occurrence attributes (optional)
                //===================================
                double pcsSectionWidth = 0, pcsSectionDepth = 0;
                double scSectionWidth = 0, scSectionDepth = 0;
                for (int i = 1; i <= 3; i++)
                {
                    BusinessObject sectionPart = componentDictionary["HgrBeam_" + i].GetRelationship("madeFrom", "part").TargetObjects[0];
                    CrossSection crossSection = (CrossSection)sectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];

                    if (i == 1)
                    {
                        pcsSectionWidth = crossSection.Width;
                        pcsSectionDepth = crossSection.Depth;
                    }
                    else if (i == 3)
                    {
                        scSectionWidth = crossSection.Width;
                        scSectionDepth = crossSection.Depth;
                    }

                    //Set the length of guide angle
                    if (i == 2 || i == 3)
                        componentDictionary["HgrBeam_" + i].SetPropertyValue(2.0 * pcsSectionDepth, "IJUAHgrOccLength", "Length");
                }

                lugOffset[0] = D + scSectionWidth;
                lugOffset[1] = lugOffset[0];

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
                //Create Joints
                //====== ======
                //Create a collection to hold the joints

                string bbxLow = string.Empty, bbxHigh = string.Empty, structure = string.Empty, structureIndex = string.Empty, lowIndex = string.Empty, highIndex = string.Empty, struct2 = string.Empty, struct2Index = string.Empty;

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

                struct2Index = "-1";
                struct2 = "Struct_2";

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

                    if (!(isOffsetApplied[1]))
                    {
                        JointHelper.CreateRigidJoint("-1", "Struct_2", CONNECTION4, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.NegativeX, 0, 0, 0);
                        struct2Index = CONNECTION4;
                        struct2 = "Connection";
                    }
                    structure = "Connection";
                    structureIndex = CONNECTION1;
                    lowIndex = CONNECTION2;
                    highIndex = CONNECTION3;

                    bbxLow = "Connection";
                    bbxHigh = "Connection";
                }

                //---- (Flexible Structural Part)----
                JointHelper.CreatePrismaticJoint(HGRBEAM1, "BeginCap", HGRBEAM1, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                string bottomStructPort = string.Empty, bottomPartIndex = string.Empty, topStructPort = string.Empty, topPartIndex = string.Empty;

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    //=============By Structure Command ====================
                    //route ---- endcap
                    JointHelper.CreatePlanarJoint(highIndex, bbxHigh, HGRBEAM1, "EndCap", Plane.ZX, Plane.NegativeYZ, -boundingBoxWidth - G);
                    //==============Top Structure Connection=============
                    //endcap ---- structure
                    if (leftPad == false)
                    {
                        topStructPort = "EndCap";
                        topPartIndex = HGRBEAM1;
                    }
                    else
                    {
                        JointHelper.CreateRigidJoint(HGRBEAM1, "EndCap", LEFTPAD, "Port1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, pcsSectionDepth / 2.0, pcsSectionWidth / 2.0);
                        topStructPort = "Port2";
                        topPartIndex = LEFTPAD;
                    }
                    JointHelper.CreatePrismaticJoint(structureIndex, structure, topPartIndex, topStructPort, Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeX, 0, -pcsSectionDepth / 2.0);
                    //============Bottom Structure Connection ==============
                    //begincap --- route(structure)
                    if (rightPad == false)
                    {
                        bottomStructPort = "BeginCap";
                        bottomPartIndex = HGRBEAM1;
                    }
                    else
                    {
                        JointHelper.CreateRigidJoint(HGRBEAM1, "BeginCap", RIGHTPAD, "Port1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0.0, pcsSectionDepth / 2.0, pcsSectionWidth / 2.0);
                        bottomStructPort = "Port2";
                        bottomPartIndex = RIGHTPAD;
                    }
                    if (isOffsetApplied[1])
                        JointHelper.CreatePlanarJoint(lowIndex, bbxLow, bottomPartIndex, bottomStructPort, Plane.XY, Plane.XY, -lugOffset[1]);
                    else
                        JointHelper.CreatePointOnPlaneJoint(bottomPartIndex, bottomStructPort, struct2Index, struct2, Plane.XY);
                }
                else
                {
                    //==============By Point Command======================
                    //route ---- endcap
                    JointHelper.CreatePrismaticJoint(highIndex, bbxHigh, HGRBEAM1, "EndCap", Plane.ZX, Plane.NegativeYZ, Axis.Z, Axis.Z, -boundingBoxWidth - G, -pcsSectionDepth / 2.0);
                    //==============Top Structure Connection=============
                    //endcap ---- structure

                    if (leftPad == false)
                    {
                        topStructPort = "EndCap";
                        topPartIndex = HGRBEAM1;
                    }
                    else
                    {
                        JointHelper.CreateRigidJoint(HGRBEAM1, "EndCap", LEFTPAD, "Port1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0.0, pcsSectionDepth / 2.0, pcsSectionWidth / 2.0);
                        topStructPort = "Port2";
                        topPartIndex = LEFTPAD;
                    }
                    JointHelper.CreatePointOnPlaneJoint(topPartIndex, topStructPort, structureIndex, structure, Plane.XY);
                    //============Bottom Structure Connection ==============
                    //begincap --- route(structure)
                    if (rightPad == false)
                    {
                        bottomStructPort = "BeginCap";
                        bottomPartIndex = HGRBEAM1;
                    }
                    else
                    {
                        JointHelper.CreateRigidJoint(HGRBEAM1, "BeginCap", RIGHTPAD, "Port1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0.0, pcsSectionDepth / 2.0, pcsSectionWidth / 2.0);
                        bottomStructPort = "Port2";
                        bottomPartIndex = RIGHTPAD;
                    }
                    if (isOffsetApplied[1])
                        JointHelper.CreatePlanarJoint(lowIndex, bbxLow, bottomPartIndex, bottomStructPort, Plane.XY, Plane.XY, -lugOffset[1]);
                    else
                        JointHelper.CreatePointOnPlaneJoint(bottomPartIndex, bottomStructPort, struct2Index, struct2, Plane.XY);
                }
                //structure support part --- angle ---- duct
                JointHelper.CreatePrismaticJoint(HGRBEAM1, "EndCap", HGRBEAM2, "EndCap", Plane.YZ, Plane.NegativeZX, Axis.Z, Axis.X, 0.0, -pcsSectionDepth / 2.0);

                JointHelper.CreatePlanarJoint(highIndex, bbxHigh, HGRBEAM2, "EndCap", Plane.XY, Plane.YZ, 0);

                //structure support part --- angle ---- duct
                JointHelper.CreatePrismaticJoint(HGRBEAM1, "EndCap", HGRBEAM3, "BeginCap", Plane.YZ, Plane.NegativeZX, Axis.Z, Axis.NegativeX, 0.0, -pcsSectionDepth / 2.0);
                JointHelper.CreatePlanarJoint(lowIndex, bbxLow, HGRBEAM3, "BeginCap", Plane.XY, Plane.NegativeYZ, 0);
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

                    routeConnections.Add(new ConnectionInfo(HGRBEAM2, 1)); // Left guide angle,Connects to First Route Input (i.e. the pipe)
                    routeConnections.Add(new ConnectionInfo(HGRBEAM3, 1)); // Right guide angle,Connects to First Route Input (i.e. the pipe)

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
                        structConnections.Add(new ConnectionInfo(LEFTPAD, 1)); // partindex, Structureindex
                    else
                        structConnections.Add(new ConnectionInfo(HGRBEAM1, 1));
                    if (rightPad)
                        structConnections.Add(new ConnectionInfo(RIGHTPAD, 1)); // partindex, Structureindex
                    else
                        structConnections.Add(new ConnectionInfo(HGRBEAM1, 1));

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



