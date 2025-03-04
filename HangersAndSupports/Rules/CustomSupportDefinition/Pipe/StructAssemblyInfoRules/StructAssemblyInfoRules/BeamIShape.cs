﻿//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   BeamIShape.cs
//   SupportStructAssemblyInfoRules,Ingr.SP3D.Content.Support.Rules.BeamIShape
//   Author       :Vijaya
//   Creation Date:3.Sep.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  03.Sep.2013     Vijaya   CR-CP-224488  Convert HgrSupStructAssmInfoRules to C# .Net  
//  31.Oct.2014     PVK      TR-CP-260301	Resolve coverity issues found in August 2014 report
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
    public class BeamIShape : CustomSupportDefinition
    {
        //Constants
        private const string HGRBEAM = "HGRBEAM";
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
                    parts.Add(new PartInfo(HGRBEAM, "HgrBeam", "PartSelectionRule,Ingr.SP3D.Content.Support.Rules.PartByLoadFactor"));                   
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
                BoundingBox boundingBox = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.Supported);

                double[] boxOffset = SupportStructAssemblyServices.GetBoundaryObjectDimension(this, boundingBox), lugOffset = new double[2];
                double boundingBoxWidth = boundingBox.Width, boundingBoxHeight = boundingBox.Height;

                Collection<object> angleConfig = null;
                bool varUBoltPart = GenericHelper.GetDataByRule("HgrSupAngleByLF", (BusinessObject)support, out angleConfig);
                if (angleConfig != null)
                {
                    lugOffset[0] = (double)angleConfig[3] - boxOffset[0] / 2;
                    lugOffset[1] = (double)angleConfig[3] - boxOffset[3] / 2;
                }

                bool[] isOffsetApplied = new bool[2];
                if (SupportHelper.SupportingObjects.Count > 1)
                    isOffsetApplied[1] = false;
                else
                    isOffsetApplied[1] = true;

                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
            
                //Create Joints
                BusinessObject horizontalSectionPart = componentDictionary[HGRBEAM].GetRelationship("madeFrom", "part").TargetObjects[0];
                CrossSection crosssection = (CrossSection)horizontalSectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];
                //Get SteelWidth and SteelDepth 
                double steelWidth=0.0,steelDepth=0.0,topOffset=0.0;
                 steelWidth = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructCrossSectionDimensions", "Width")).PropValue;
                 steelDepth = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructCrossSectionDimensions", "Depth")).PropValue;
                string bottomStructPort = string.Empty, bottomPart, topPart, topStructPort;

                //Flexible Structural Part
                JointHelper.CreatePrismaticJoint(HGRBEAM, "BeginCap", HGRBEAM, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    if (Configuration == 1)
                        JointHelper.CreatePlanarJoint("-1", "BBSR_High", HGRBEAM, "EndCap", Plane.ZX, Plane.NegativeYZ, -boundingBoxWidth - gap);
                    else
                        JointHelper.CreatePlanarJoint("-1", "BBSR_High", HGRBEAM, "EndCap", Plane.ZX, Plane.YZ, gap);

                    if (!leftPad)
                    {
                        topStructPort = "EndCap";
                        topPart = HGRBEAM;
                        topOffset = steelDepth / 2;
                    }
                    else
                    {
                        JointHelper.CreatePlanarJoint(LEFTPAD, "HgrPort_1", HGRBEAM, "EndCap", Plane.XY, Plane.XY, 0);
                        JointHelper.CreatePrismaticJoint(LEFTPAD, "HgrPort_1", HGRBEAM, "Neutral", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                        topStructPort = "HgrPort_2";
                        topPart = LEFTPAD;
                        topOffset = 0;
                    }
                    if (Configuration == 1)
                        JointHelper.CreatePrismaticJoint("-1", "Structure", topPart, topStructPort, Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeX, 0, -topOffset);
                    else
                        JointHelper.CreatePrismaticJoint("-1", "Structure", topPart, topStructPort, Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, topOffset);

                    if (!rightPad)
                    {
                        bottomStructPort = "BeginCap";
                        bottomPart = HGRBEAM;
                    }
                    else
                    {
                        JointHelper.CreatePlanarJoint(RIGHTPAD, "HgrPort_2", HGRBEAM, "BeginCap", Plane.XY, Plane.XY, 0);
                        JointHelper.CreatePrismaticJoint(RIGHTPAD, "HgrPort_2", HGRBEAM, "Neutral", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                        bottomStructPort = "HgrPort_1";
                        bottomPart = RIGHTPAD;
                    }
                    if (isOffsetApplied[1])
                        JointHelper.CreatePlanarJoint("-1", "BBSR_Low", bottomPart, bottomStructPort, Plane.XY, Plane.XY, -(lugOffset[1] + d / 2));
                    else
                        JointHelper.CreatePointOnPlaneJoint(bottomPart, bottomStructPort, "-1", "Struct_2", Plane.XY);
                }
                else
                {
                    if (Configuration == 1)
                        JointHelper.CreatePrismaticJoint("-1", "BBR_High", HGRBEAM, "EndCap", Plane.ZX, Plane.NegativeYZ, Axis.Z, Axis.Z, -boundingBoxWidth - gap, -steelDepth / 2);
                    else
                        JointHelper.CreatePrismaticJoint("-1", "BBR_High", HGRBEAM, "EndCap", Plane.ZX, Plane.YZ, Axis.Z, Axis.Z, gap, steelDepth / 2);
                    
                    if (!leftPad)
                    {
                        topStructPort = "EndCap";
                        topPart = HGRBEAM;
                    }
                    else
                    {
                        JointHelper.CreatePlanarJoint(LEFTPAD, "HgrPort_1", HGRBEAM, "EndCap", Plane.XY, Plane.XY, 0);
                        JointHelper.CreatePrismaticJoint(LEFTPAD, "HgrPort_1", HGRBEAM, "Neutral", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                        topStructPort = "HgrPort_2";
                        topPart = LEFTPAD;
                    }
                    JointHelper.CreatePointOnPlaneJoint(topPart, topStructPort, "-1", "Structure", Plane.XY);

                    if (!rightPad)
                    {
                        bottomStructPort = "BeginCap";
                        bottomPart = HGRBEAM;
                    }
                    else
                    {
                        JointHelper.CreatePlanarJoint(RIGHTPAD, "HgrPort_2", HGRBEAM, "BeginCap", Plane.XY, Plane.XY, 0);
                        JointHelper.CreatePrismaticJoint(RIGHTPAD, "HgrPort_2", HGRBEAM, "Neutral", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                        bottomStructPort = "HgrPort_1";
                        bottomPart = RIGHTPAD;
                    }

                    if (isOffsetApplied[1])
                        JointHelper.CreatePlanarJoint("-1", "BBR_Low", bottomPart, bottomStructPort, Plane.XY, Plane.XY, -(lugOffset[1] + d / 2));
                    else
                        JointHelper.CreatePointOnPlaneJoint(bottomPart, bottomStructPort, "-1", "Struct_2", Plane.XY);
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
                        routeConnections.Add(new ConnectionInfo(HGRBEAM, index)); // partindex, routeindex

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
                    if (rightPad)
                        structConnections.Add(new ConnectionInfo(RIGHTPAD, 1)); // partindex, Structureindex
                    else
                        structConnections.Add(new ConnectionInfo(HGRBEAM, 1)); // partindex, Structureindex

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



