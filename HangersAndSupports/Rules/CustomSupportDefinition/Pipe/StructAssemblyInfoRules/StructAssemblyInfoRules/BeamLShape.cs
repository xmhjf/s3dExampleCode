﻿//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   BeamLShape.cs
//   SupportStructAssemblyInfoRules,Ingr.SP3D.Content.Support.Rules.BeamLShape
//   Author       :Vijaya
//   Creation Date:30.Aug.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  30.Aug.2013     Vijaya   CR-CP-224488  Convert HgrSupStructAssmInfoRules to C# .Net  
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
    public class BeamLShape : CustomSupportDefinition
    {
        //Constants
        private const string HGRBEAM1 = "HGRBEAM1";
        private const string HGRBEAM2 = "HGRBEAM2";
        private const string LEFTPAD = "LEFTPAD";
        private const string RIGHTPAD = "RIGHTPAD";

        bool leftPad, rightPad;
        double d, gap;
     
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
                    object[] attributeCollection = SupportStructAssemblyServices.GetPipeStructuralASMAttributes(this);

                    //Get the attributes from assembly
                    d = (double)attributeCollection[0];
                    gap = (double)attributeCollection[1];
                    leftPad = (bool)attributeCollection[2];
                    rightPad = (bool)attributeCollection[3];

                    //Use the default selection rule to get a catalog part for each part class  
                    parts.Add(new PartInfo(HGRBEAM1, "HgrBeam", "PartSelectionRule,Ingr.SP3D.Content.Support.Rules.PartByLoadFactor"));
                    parts.Add(new PartInfo(HGRBEAM2, "HgrBeam", "PartSelectionRule,Ingr.SP3D.Content.Support.Rules.PartByLoadFactor"));
                  
                    if (rightPad)
                        parts.Add(new PartInfo(RIGHTPAD, SupportStructAssemblyServices.GetPadPartNameByCrossSectionType(this, parts[0])));
                    if (leftPad)
                        parts.Add(new PartInfo(LEFTPAD, SupportStructAssemblyServices.GetPadPartNameByCrossSectionType(this, parts[0])));
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

                int routeStructureConfiguration = 0;
                bool[] isOffsetApplied = new bool[2];
                double[] lugOffset = new double[2];
                routeStructureConfiguration = SupportStructAssemblyServices.GetLSupportConfiguration(this, ref isOffsetApplied, ref lugOffset);
                if (routeStructureConfiguration == -1)
                    routeStructureConfiguration = 2;
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
               
                //define physical connection information
                //if (routeStructureConfiguration == 1)
                //    CreatePhysicalConnection(HGRBEAM1, "BeginCap", HGRBEAM2, "EndCap", 301);
                //else
                //    CreatePhysicalConnection(HGRBEAM1, "EndCap", HGRBEAM2, "BeginCap", 301);


                BusinessObject horizontalSectionPart = componentDictionary[HGRBEAM1].GetRelationship("madeFrom", "part").TargetObjects[0];
                CrossSection crosssection = (CrossSection)horizontalSectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];
                //Get SteelWidth and SteelDepth 
                double steelWidth = 0.0, steelDepth = 0.0, topOffset = 0.0;
                steelWidth = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructCrossSectionDimensions", "Width")).PropValue;
                steelDepth = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructCrossSectionDimensions", "Depth")).PropValue;
                bool toggleRouteConfiguration = (SupportHelper.SupportingObjects.Count == 1 && Configuration != 1);
                string bottomStructPort = string.Empty, bottomPart, topPart, topStructPort;

                //Create Joints
                JointHelper.CreatePrismaticJoint(HGRBEAM2, "BeginCap", HGRBEAM2, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                JointHelper.CreatePrismaticJoint(HGRBEAM1, "BeginCap", HGRBEAM1, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                if (routeStructureConfiguration == 1)
                {
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        JointHelper.CreatePrismaticJoint("-1", "BBSR_High", HGRBEAM2, "EndCap", Plane.XY, Plane.YZ, Axis.X, Axis.Y, gap, lugOffset[0]);
                    else
                        if (isOffsetApplied[0])
                            JointHelper.CreateRigidJoint("-1", "BBR_High", HGRBEAM2, "EndCap", Plane.XY, Plane.YZ, Axis.X, Axis.Y, gap, lugOffset[0], -steelDepth / 2);
                        else
                            JointHelper.CreatePrismaticJoint("-1", "BBR_High", HGRBEAM2, "EndCap", Plane.XY, Plane.YZ, Axis.Y, Axis.Z, gap, 0);

                    if (rightPad)
                    {
                        JointHelper.CreatePlanarJoint(HGRBEAM2, "BeginCap", RIGHTPAD, "HgrPort_2", Plane.XY, Plane.XY, 0);
                        JointHelper.CreatePrismaticJoint(RIGHTPAD, "HgrPort_2", HGRBEAM2, "Neutral", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                        bottomStructPort = "HgrPort_1";
                        bottomPart = RIGHTPAD;
                    }
                    else
                    {
                        bottomStructPort = "BeginCap";
                        bottomPart = HGRBEAM2;
                    }

                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        if (isOffsetApplied[1])
                            JointHelper.CreatePlanarJoint("-1", "BBSR_Low", HGRBEAM2, "BeginCap", Plane.ZX, Plane.XY, -lugOffset[1]);
                        else
                            JointHelper.CreatePointOnPlaneJoint(bottomPart, bottomStructPort, "-1", "Struct_2", Plane.XY);
                    else
                        if (isOffsetApplied[1])
                            JointHelper.CreatePlanarJoint("-1", "BBR_Low", HGRBEAM2, "BeginCap", Plane.ZX, Plane.XY, -lugOffset[1]);
                        else
                            JointHelper.CreatePointOnPlaneJoint(bottomPart, bottomStructPort, "-1", "Struct_2", Plane.XY);

                    JointHelper.CreateRigidJoint(HGRBEAM2, "EndCap", HGRBEAM1, "BeginCap", Plane.XY, Plane.NegativeYZ, Axis.X, Axis.Z, 0, 0, 0);

                    if (leftPad)
                    {
                        JointHelper.CreatePlanarJoint(HGRBEAM1, "EndCap", LEFTPAD, "HgrPort_1", Plane.XY, Plane.XY, 0);
                        JointHelper.CreatePrismaticJoint(LEFTPAD, "HgrPort_1", HGRBEAM1, "Neutral", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                        topStructPort = "HgrPort_2";
                        topPart = LEFTPAD;
                        topOffset = 0;
                    }
                    else
                    {
                        topOffset = steelDepth / 2;
                        topStructPort = "EndCap";
                        topPart = HGRBEAM1;
                    }
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        JointHelper.CreatePrismaticJoint("-1", "Structure", topPart, topStructPort, Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeX, 0, -topOffset);

                    else
                        JointHelper.CreatePlanarJoint("-1", "Structure", topPart, topStructPort, Plane.XY, Plane.NegativeXY, 0);
                }

                else
                {
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        if (!toggleRouteConfiguration)
                            JointHelper.CreatePrismaticJoint("-1", "BBSR_Low", HGRBEAM2, "BeginCap", Plane.XY, Plane.YZ, Axis.X, Axis.Y, boundingBoxHeight + gap, -lugOffset[0]);
                        else
                            JointHelper.CreatePrismaticJoint("-1", "BBSR_Low", HGRBEAM2, "BeginCap", Plane.XY, Plane.YZ, Axis.X, Axis.NegativeY, boundingBoxHeight + gap, lugOffset[0] + boundingBoxWidth);
                    else
                        if (isOffsetApplied[0])
                            if (!toggleRouteConfiguration)
                                JointHelper.CreateRigidJoint("-1", "BBR_Low", HGRBEAM2, "BeginCap", Plane.XY, Plane.YZ, Axis.X, Axis.Y, boundingBoxHeight + gap, -lugOffset[0], -steelDepth / 2);
                            else
                                JointHelper.CreateRigidJoint("-1", "BBR_Low", HGRBEAM2, "BeginCap", Plane.XY, Plane.YZ, Axis.X, Axis.NegativeY, boundingBoxHeight + gap, lugOffset[0] + boundingBoxWidth, steelDepth / 2);
                        else
                            JointHelper.CreatePrismaticJoint("-1", "BBR_Low", HGRBEAM2, "BeginCap", Plane.XY, Plane.YZ, Axis.Y, Axis.Z, boundingBoxHeight + gap, 0);

                    if (rightPad)
                    {
                        JointHelper.CreatePlanarJoint(HGRBEAM2, "EndCap", RIGHTPAD, "HgrPort_1", Plane.XY, Plane.XY, 0);
                        JointHelper.CreatePrismaticJoint(RIGHTPAD, "HgrPort_1", HGRBEAM2, "Neutral", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                        bottomStructPort = "HgrPort_2";
                        bottomPart = RIGHTPAD;
                    }
                    else
                    {
                        bottomStructPort = "EndCap";
                        bottomPart = HGRBEAM2;
                    }
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    {
                        if (isOffsetApplied[1])
                            if (!toggleRouteConfiguration)
                                JointHelper.CreatePlanarJoint("-1", "BBSR_High", HGRBEAM2, "EndCap", Plane.ZX, Plane.XY, lugOffset[1]);
                            else
                                JointHelper.CreatePlanarJoint("-1", "BBSR_High", HGRBEAM2, "EndCap", Plane.ZX, Plane.NegativeXY, -(lugOffset[1] + boundingBoxWidth));
                        else
                            JointHelper.CreatePointOnPlaneJoint(bottomPart, bottomStructPort, "-1", "Struct_2", Plane.XY);
                    }
                    else
                    {
                        if (isOffsetApplied[1])
                            if (!toggleRouteConfiguration)
                                JointHelper.CreatePlanarJoint("-1", "BBR_High", HGRBEAM2, "EndCap", Plane.ZX, Plane.XY, lugOffset[1]);
                            else
                                JointHelper.CreatePlanarJoint("-1", "BBR_High", HGRBEAM2, "EndCap", Plane.ZX, Plane.NegativeXY, -(lugOffset[1] + boundingBoxWidth));
                        else
                            JointHelper.CreatePointOnPlaneJoint(bottomPart, bottomStructPort, "-1", "Struct_2", Plane.XY);
                    }
                    JointHelper.CreateRigidJoint(HGRBEAM1, "EndCap", HGRBEAM2, "BeginCap", Plane.XY, Plane.NegativeYZ, Axis.X, Axis.Z, 0, 0, 0);
                    if (leftPad)
                    {
                        JointHelper.CreatePlanarJoint(HGRBEAM1, "BeginCap", LEFTPAD, "HgrPort_2", Plane.XY, Plane.XY, 0);
                        JointHelper.CreatePrismaticJoint(LEFTPAD, "HgrPort_2", HGRBEAM1, "Neutral", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                        topStructPort = "HgrPort_1";
                        topPart = LEFTPAD;
                        topOffset = 0;
                    }
                    else
                    {
                        topStructPort = "BeginCap";
                        topPart = HGRBEAM1;
                        topOffset = steelDepth / 2;
                    }

                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    {
                        if (!toggleRouteConfiguration)
                            JointHelper.CreatePrismaticJoint("-1", "Structure", topPart, topStructPort, Plane.XY, Plane.XY, Axis.X, Axis.X, 0, -topOffset);
                        else
                            JointHelper.CreatePrismaticJoint("-1", "Structure", topPart, topStructPort, Plane.XY, Plane.XY, Axis.X, Axis.NegativeX, 0, topOffset);
                    }
                    else
                    {
                        if (isOffsetApplied[0])
                            JointHelper.CreatePlanarJoint("-1", "Structure", topPart, topStructPort, Plane.XY, Plane.XY, 0);
                        else
                            JointHelper.CreatePlanarSlotJoint("-1", "Structure", topPart, topStructPort, Plane.XY, Plane.XY, Axis.X, 0, 0);
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

                    for (int index = 1; index <= SupportHelper.SupportedObjects.Count; index++)
                        routeConnections.Add(new ConnectionInfo(HGRBEAM2, index)); // partindex, routeindex

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
                        structConnections.Add(new ConnectionInfo(HGRBEAM2, 1)); // partindex, Structureindex

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

