﻿//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   LDuctClamp.cs
//   SupDuctAssmInfoRules,Ingr.SP3D.Content.Support.Rules.LDuctClamp
//   Author       :Hema
//   Creation Date:05.Sep.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  05.Sep.2013     Hema   CR-CP-224482 Convert HgrSupDuctAssmInfoRules to C# .Net
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

                    parts.Add(new PartInfo(HGRBEAM1, "HgrBeam"));
                    parts.Add(new PartInfo(HGRBEAM2, "HgrBeam"));
                    parts.Add(new PartInfo(DUCTCLAMP, "DuctClamp"));

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
                double clampLength = (double)((PropertyValueDouble)ductClampPart.GetPropertyValue("IJUAHgrDuctClamp", "ClampLength")).PropValue;

                lugOffset[0] = D + (clampLength - boundingBoxWidth) / 2.0;
                lugOffset[1] = lugOffset[0];

                //define physical connection information
                //if (routeStructureConfiguration == 1)
                    //CreatePhysicalConnection(HGRBEAM1, "BeginCap", HGRBEAM2, "EndCap", 301);
                //else
                    //CreatePhysicalConnection(HGRBEAM1, "EndCap", HGRBEAM2, "BeginCap", 301);
                //
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
                            JointHelper.CreateRigidJoint("-1", "BBR_High", HGRBEAM2, "EndCap", Plane.XY, Plane.YZ, Axis.X, Axis.Y, G, lugOffset[0], 0);
                        else
                            JointHelper.CreatePrismaticJoint("-1", "BBR_High", HGRBEAM2, "EndCap", Plane.XY, Plane.YZ, Axis.Y, Axis.Z, 0.0, 0.0);
                    }

                    //----------------------------------------------------
                    //Create the Plane Joint between the RteHigh Reference Port
                    //and the Right Bottom Symbol

                    if (rightPad)
                    {
                        JointHelper.CreateRigidJoint(HGRBEAM2, "BeginCap", RIGHTPAD, "HgrPort_1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0.0, pcsSectionDepth / 2.0, pcsSectionWidth / 2.0);
                        bottomStructPort = "HgrPort_2";
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
                        JointHelper.CreateRigidJoint(HGRBEAM1, "EndCap", LEFTPAD, "HgrPort_1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, pcsSectionDepth / 2.0, pcsSectionWidth / 2.0);
                        topStructPort = "HgrPort_2";
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
                            JointHelper.CreateRigidJoint("-1", "BBR_Low", HGRBEAM2, "BeginCap", Plane.XY, Plane.YZ, Axis.X, Axis.Y, boundingBoxHeight + G, -lugOffset[0], 0);
                        else
                            JointHelper.CreatePrismaticJoint("-1", "BBR_Low", HGRBEAM2, "BeginCap", Plane.XY, Plane.YZ, Axis.Y, Axis.Z, boundingBoxHeight, 0);
                    }

                    //----------------------------------------------------
                    //Create the Plane Joint between the RteHigh Reference Port
                    //and the Right Bottom Symbol
                    if (rightPad)
                    {
                        JointHelper.CreateRigidJoint(HGRBEAM2, "EndCap", RIGHTPAD, "HgrPort_1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, pcsSectionDepth / 2.0, pcsSectionWidth / 2.0);
                        bottomStructPort = "HgrPort_2";
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
                        JointHelper.CreateRigidJoint(HGRBEAM1, "BeginCap", LEFTPAD, "HgrPort_1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, pcsSectionWidth / 2.0, pcsSectionDepth / 2.0);
                        topStructPort = "HgrPort_2";
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
                JointHelper.CreatePrismaticJoint(HGRBEAM2, "BeginCap", DUCTCLAMP, "HgrPort_2", Plane.YZ, Plane.XY, Axis.Z, Axis.Y, 0.0, pcsSectionDepth / 2.0);
                JointHelper.CreatePointOnAxisJoint(DUCTCLAMP, "HgrPort_1", "-1", "Route", Axis.X);
                JointHelper.CreateCylindricalJoint(DUCTCLAMP, "HgrPort_2", DUCTCLAMP, "HgrPort_1", Axis.Z, Axis.Z, 0);
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





