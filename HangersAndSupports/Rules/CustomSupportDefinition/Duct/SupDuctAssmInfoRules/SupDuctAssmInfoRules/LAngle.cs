//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   LAngle.cs
//   SupDuctAssmInfoRules,Ingr.SP3D.Content.Support.Rules.LAngle
//   Author       :Vijay
//   Creation Date:10.Sep.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   10.Sep.2013     Vijay   CR-CP-224482 Convert HgrSupDuctAssmInfoRules to C# .Net
//   01.Jun.2016     Vinay   TR-CP-295807	L Shape w/angle guide duct support assembly gets misaligned when ‘placed by ref’
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
    public class LAngle : CustomSupportDefinition
    {
        //Constants
        private const string HGRBEAM1 = "HgrBeam_1"; //Hanger beam object
        private const string HGRBEAM2 = "HgrBeam_2"; //Hanger beam object
        private const string HGRBEAM3 = "HgrBeam_3";          //Hanger beam object as angle guide
        private const string HGRBEAM4 = "HgrBeam_4";      //Hanger beam object as angle guide
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
                    parts.Add(new PartInfo(HGRBEAM3, "HgrBeam", "HgrSupSecondaryCS"));
                    parts.Add(new PartInfo(HGRBEAM4, "HgrBeam", "HgrSupSecondaryCS"));

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
                int routeStructConfiguration;
                if (SupportHelper.PlacementType == PlacementType.PlaceByReference)
                {
                    routeStructConfiguration = -1;
                    isOffsetApplied[0] = true;
                    isOffsetApplied[1] = true;     //For Place By reference
                }
                else
                    routeStructConfiguration = SupDuctAssemblyServices.GetDuctLSupportConfiguration(this, isOffsetApplied);

                if (routeStructConfiguration == -1)
                    routeStructConfiguration = Configuration;

                //===========================
                // 3.  Get the input objects information
                //   Dimensions of the pipe/duct cross section
                //   (Applicable for Pipe/Duct/Cableway)
                //==========================

                //===================================
                //4. Set symbol occurrence attributes (optional)
                //===================================

                double pcsSectionWidth = 0, pcsSectionDepth = 0;
                double scsSectionWidth = 0, scsSectionDepth = 0;
                for (int i = 1; i <= 4; i++)
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
                        scsSectionWidth = crossSection.Width;
                        scsSectionDepth = crossSection.Depth;
                    }

                    //Set the length of guide angle
                    if (i == 3 || i == 4)
                        componentDictionary["HgrBeam_" + i].SetPropertyValue(2.0 * pcsSectionDepth, "IJUAHgrOccLength", "Length");
                }

                lugOffset[0] = D + scsSectionWidth;
                lugOffset[1] = lugOffset[0];

                //define physical connection information
                //if (routeStructConfiguration == 1)
                    //CreatePhysicalConnection(HGRBEAM1, "BeginCap", HGRBEAM2, "EndCap", 301);
                //else
                    //CreatePhysicalConnection(HGRBEAM1, "EndCap", HGRBEAM2, "BeginCap", 301);
                //
                //====== ======
                // Create Joints
                //====== ======
                //Create a collection to hold the joints

                //Add a Prismatic Joint defining the flexible bottom member
                JointHelper.CreatePrismaticJoint(HGRBEAM2, "BeginCap", HGRBEAM2, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                //Add prismatic joint for the flexible memeber
                JointHelper.CreatePrismaticJoint(HGRBEAM1, "BeginCap", HGRBEAM1, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                string bottomStructPort = string.Empty,bottomPartIndex = string.Empty,topStructPort = string.Empty,topPartIndex = string.Empty;
                double axisOffset;
                if (routeStructConfiguration == 1)
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
                            JointHelper.CreatePrismaticJoint("-1", "BBR_High", HGRBEAM2, "EndCap", Plane.XY, Plane.YZ, Axis.Y, Axis.Z, G, 0);
                    }

                    //----------------------------------------------------
                    //Create the Plane Joint between the RteHigh Reference Port
                    //and the Right Bottom Symbol

                    if (rightPad)
                    {
                        JointHelper.CreateRigidJoint(HGRBEAM2, "BeginCap", RIGHTPAD, "HgrPort_1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
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
                        JointHelper.CreatePointOnAxisJoint(topPartIndex, topStructPort, "-1", "Structure", Axis.X);
                    else
                    {
                        if (isOffsetApplied[0])
                            JointHelper.CreatePointOnPlaneJoint(topPartIndex, topStructPort, "-1", "Structure", Plane.XY);
                        else
                            JointHelper.CreatePointOnAxisJoint(topPartIndex, topStructPort, "-1", "Structure", Axis.X);
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
                            JointHelper.CreatePrismaticJoint("-1", "BBR_Low", HGRBEAM2, "BeginCap", Plane.XY, Plane.YZ, Axis.Y, Axis.Z, boundingBoxHeight + G, 0);

                    }

                    //----------------------------------------------------
                    //Create the Plane Joint between the RteHigh Reference Port
                    //and the Right Bottom Symbol
                    if (rightPad)
                    {
                        JointHelper.CreateRigidJoint(HGRBEAM2, "EndCap", RIGHTPAD, "HgrPort_1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, pcsSectionDepth / 2.0, pcsSectionWidth / 2.0);
                        bottomStructPort = "HgrPort_2";
                        bottomPartIndex = RIGHTPAD;
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
                        JointHelper.CreateRigidJoint(HGRBEAM1, "BeginCap", LEFTPAD, "HgrPort_1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, pcsSectionDepth / 2.0, pcsSectionWidth / 2.0);
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
                        JointHelper.CreatePointOnAxisJoint(topPartIndex, topStructPort, "-1", "Structure", Axis.X);
                    else
                    {
                        if (isOffsetApplied[0])
                            JointHelper.CreatePointOnPlaneJoint(topPartIndex, topStructPort, "-1", "Structure", Plane.XY);
                        else
                            JointHelper.CreatePointOnAxisJoint(topPartIndex, topStructPort, "-1", "Structure", Axis.X);
                    }
                }

                //structure support part --- angle ---- duct
                JointHelper.CreatePrismaticJoint(HGRBEAM2, "EndCap", HGRBEAM3, "EndCap", Plane.YZ, Plane.NegativeZX, Axis.Z, Axis.X, 0, -pcsSectionDepth / 2.0);
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    JointHelper.CreatePlanarJoint("-1", "BBSR_High", HGRBEAM3, "EndCap", Plane.ZX, Plane.YZ, 0);
                else
                    JointHelper.CreatePlanarJoint("-1", "BBR_High", HGRBEAM3, "EndCap", Plane.ZX, Plane.YZ, 0);

                //structure support part --- angle ---- duct
                JointHelper.CreatePrismaticJoint(HGRBEAM2, "EndCap", HGRBEAM4, "BeginCap", Plane.YZ, Plane.NegativeZX, Axis.Z, Axis.NegativeX, 0, -pcsSectionDepth / 2.0);
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    JointHelper.CreatePlanarJoint("-1", "BBSR_Low", HGRBEAM4, "BeginCap", Plane.ZX, Plane.NegativeYZ, 0);
                else
                    JointHelper.CreatePlanarJoint("-1", "BBR_Low", HGRBEAM4, "BeginCap", Plane.ZX, Plane.NegativeYZ, 0);
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

                    routeConnections.Add(new ConnectionInfo(HGRBEAM3, 1)); // partindex, routeindex
                    routeConnections.Add(new ConnectionInfo(HGRBEAM4, 1)); // partindex, routeindex

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
                        structConnections.Add(new ConnectionInfo(LEFTPAD, 1)); // partindex, routeindex
                    else
                        structConnections.Add(new ConnectionInfo(HGRBEAM1, 1));// partindex, routeindex
                    if (rightPad)
                        structConnections.Add(new ConnectionInfo(RIGHTPAD, 1)); // partindex, routeindex
                    else
                        structConnections.Add(new ConnectionInfo(HGRBEAM2, 1));// partindex, routeindex

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



