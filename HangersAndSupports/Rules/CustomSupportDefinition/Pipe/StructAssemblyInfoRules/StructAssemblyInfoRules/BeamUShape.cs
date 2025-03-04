//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   BeamUShape.cs
//   SupportStructAssemblyInfoRules,Ingr.SP3D.Content.Support.Rules.BeamUShape
//   Author       :Vijaya
//   Creation Date:5.Aug.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  5.Aug.2013     Vijaya   CR-CP-224488  Convert HgrSupStructAssmInfoRules to C# .Net  
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.Generic;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;

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
    public class BeamUShape : CustomSupportDefinition
    {
        //Constants
        private const string HGRBEAM1 = "HGRBEAM1";
        private const string HGRBEAM2 = "HGRBEAM2";
        private const string HGRBEAM3 = "HGRBEAM3";       
        private const string LEFTPAD = "LEFTPAD";
        private const string RIGHTPAD = "RIGHTPAD";
        private static bool dimensionPort = false;
        string[] partKeys=new string[4];
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

                    CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                    PartClass partClass = (PartClass)catalogBaseHelper.GetPartClass("HSDimensionPort");
                    
                    //Use the default selection rule to get a catalog part for each part class
                    for (int index = 1; index <= 3; index++)
                        parts.Add(new PartInfo("HGRBEAM" + index, "HgrBeam", "PartSelectionRule,Ingr.SP3D.Content.Support.Rules.PartByLoadFactor"));

                    if (partClass != null)
                    {
                        for (int index = 1; index <= partKeys.Length; index++)
                        {
                            partKeys[index-1] = "HSDIMENSIONPORT" + index;
                            parts.Add(new PartInfo(partKeys[index-1], "HSDimensionPort_1"));
                        }
                        dimensionPort=true;
                    }
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
                //load standard bounding box.
                BoundingBoxHelper.CreateStandardBoundingBoxes(false);

                BoundingBox boundingBox;
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    boundingBox = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.SupportedAndSupporting);
                else
                    boundingBox = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.Supported);

                double boundingBoxWidth = boundingBox.Width, boundingBoxHeight = boundingBox.Height;

                double[] boxOffset = SupportStructAssemblyServices.GetBoundaryObjectDimension(this, boundingBox);
                double[] lugoffset = new double[2];
                lugoffset[0] = d / 2.0 - boxOffset[1] / 2.0;
                lugoffset[1] = d / 2.0 - boxOffset[3] / 2.0;
                bool[] isOffsetApplied = SupportStructAssemblyServices.GetIsLugEndOffsetApplied(this);
                string[] structPort = SupportStructAssemblyServices.GetIndexedStructPortName(this, isOffsetApplied);
                string leftStructPort = structPort[0],rightStructPort = structPort[1];                

                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
                double beginOverLength, endOverLength, leftVerticalDistance, rightVerticalDistance;
                try
                {
                    beginOverLength = (double)((PropertyValueDouble)componentDictionary[HGRBEAM1].GetPropertyValue("IJUAHgrOccOverLength", "BeginOverLength")).PropValue;
                }
                catch
                {
                    beginOverLength = 0.0;
                }
                try
                {
                    endOverLength = (double)((PropertyValueDouble)componentDictionary[HGRBEAM1].GetPropertyValue("IJUAHgrOccOverLength", "EndOverLength")).PropValue;
                }
                catch
                {
                    endOverLength = 0.0;
                }             

                //define physical connection information
                //CreatePhysicalConnection(HGRBEAM1, "EndCap", HGRBEAM2, "BeginCap", 301);
                //CreatePhysicalConnection(HGRBEAM2, "EndCap", HGRBEAM3, "BeginCap", 301);

                //Create Joints               
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    JointHelper.CreatePrismaticJoint("-1", "BBSR_Low", HGRBEAM2, "BeginCap", Plane.XY, Plane.YZ, Axis.X, Axis.Y, boundingBoxHeight + gap, -lugoffset[0]);
                else
                    if (isOffsetApplied[0])
                        JointHelper.CreateRigidJoint("-1", "BBR_Low", HGRBEAM2, "BeginCap", Plane.XY, Plane.YZ, Axis.X, Axis.Y, boundingBoxHeight + gap, -lugoffset[0], 0.0);
                    else
                        JointHelper.CreatePrismaticJoint("-1", "BBR_Low", HGRBEAM2, "BeginCap", Plane.XY, Plane.YZ, Axis.Y, Axis.Z, boundingBoxHeight + gap, 0);

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    JointHelper.CreatePlanarJoint("-1", "BBSR_High", HGRBEAM2, "EndCap", Plane.ZX, Plane.XY, lugoffset[1]);
                else
                    if (isOffsetApplied[1])
                        JointHelper.CreatePlanarJoint("-1", "BBR_High", HGRBEAM2, "EndCap", Plane.ZX, Plane.XY, lugoffset[1]);

                leftVerticalDistance = RefPortHelper.DistanceBetweenPorts("BBRV_High", leftStructPort, PortDistanceType.Vertical);
                rightVerticalDistance = RefPortHelper.DistanceBetweenPorts("BBRV_High", rightStructPort, PortDistanceType.Vertical);

                JointHelper.CreatePrismaticJoint(HGRBEAM2, "BeginCap", HGRBEAM2, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                if (dimensionPort)
                {
                    JointHelper.CreateRigidJoint(HGRBEAM2, "BeginCap", partKeys[0], "Dimension", Plane.XY, Plane.XY, Axis.X, Axis.X, 0.0, 0.0, 0.0);
                    JointHelper.CreateRigidJoint(partKeys[0], "Dimension", partKeys[1], "Dimension", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, boundingBoxWidth + beginOverLength + endOverLength + lugoffset[0] + lugoffset[1], 0.0);
                    JointHelper.CreateRigidJoint(partKeys[0], "Dimension", partKeys[2], "Dimension", Plane.XY, Plane.XY, Axis.X, Axis.X, 0.0, 0.0, leftVerticalDistance - gap);
                    JointHelper.CreateRigidJoint(partKeys[0], "Dimension", partKeys[3], "Dimension", Plane.XY, Plane.XY, Axis.X, Axis.X, 0.0, 0.0, rightVerticalDistance - gap);

                    SupportStructAssemblyServices.CreateDimensionNote(this, "L_Start", componentDictionary[partKeys[0]], "Dimension");
                    SupportStructAssemblyServices.CreateDimensionNote(this, "L_End", componentDictionary[partKeys[1]], "Dimension");
                    SupportStructAssemblyServices.CreateDimensionNote(this, "L_Height", componentDictionary[partKeys[2]], "Dimension");
                    SupportStructAssemblyServices.CreateDimensionNote(this, "L_Height1", componentDictionary[partKeys[3]], "Dimension");
                }

                JointHelper.CreateRigidJoint(HGRBEAM1, "EndCap", HGRBEAM2, "BeginCap", Plane.XY, Plane.NegativeYZ, Axis.X, Axis.Z, 0.0, 0.0, 0.0);
                JointHelper.CreateRigidJoint(HGRBEAM3, "BeginCap", HGRBEAM2, "EndCap", Plane.XY, Plane.YZ, Axis.X, Axis.NegativeZ, 0.0, 0.0, 0.0);
                JointHelper.CreatePrismaticJoint(HGRBEAM1, "BeginCap", HGRBEAM1, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0.0, 0.0);
                JointHelper.CreatePrismaticJoint(HGRBEAM3, "BeginCap", HGRBEAM3, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0.0, 0.0);

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    if (leftPad)
                    {
                        JointHelper.CreatePrismaticJoint("-1", leftStructPort, LEFTPAD, "HgrPort_1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0);
                        JointHelper.CreatePlanarJoint(LEFTPAD, "HgrPort_2", HGRBEAM1, "BeginCap", Plane.XY, Plane.XY, 0.0);
                        JointHelper.CreatePrismaticJoint(LEFTPAD, "HgrPort_2", HGRBEAM1, "Neutral", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                    }
                    else
                        JointHelper.CreatePrismaticJoint("-1", leftStructPort, HGRBEAM1, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.X, 0.0, 0.0);
                }
                else
                {
                    if (leftPad)
                    {
                        if (isOffsetApplied[1])
                            JointHelper.CreatePlanarJoint("-1", leftStructPort, LEFTPAD, "HgrPort_1", Plane.XY, Plane.XY, 0);
                        else
                            JointHelper.CreatePlanarSlotJoint("-1", leftStructPort, LEFTPAD, "HgrPort_1", Plane.XY, Plane.XY, Axis.X, 0, 0);

                        JointHelper.CreatePlanarJoint(LEFTPAD, "HgrPort_2", HGRBEAM1, "BeginCap", Plane.XY, Plane.XY, 0);
                        JointHelper.CreatePrismaticJoint(LEFTPAD, "HgrPort_2", HGRBEAM1, "Neutral", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                    }
                    else
                    {
                        if (isOffsetApplied[1])
                            JointHelper.CreatePlanarJoint("-1", leftStructPort, HGRBEAM1, "BeginCap", Plane.XY, Plane.XY, 0);
                        else
                            JointHelper.CreatePlanarSlotJoint("-1", leftStructPort, HGRBEAM1, "BeginCap", Plane.XY, Plane.XY, 0, 0, 0);
                    }

                }
                if (rightPad)
                {
                    if (isOffsetApplied[1])
                        JointHelper.CreatePlanarJoint("-1", rightStructPort, RIGHTPAD, "HgrPort_2", Plane.XY, Plane.NegativeXY, 0);
                    else
                        JointHelper.CreatePointOnAxisJoint(RIGHTPAD, "HgrPort_1", "-1", rightStructPort, Axis.X);

                    JointHelper.CreatePlanarJoint(RIGHTPAD, "HgrPort_1", HGRBEAM3, "EndCap", Plane.XY, Plane.XY, 0);
                    JointHelper.CreatePrismaticJoint(RIGHTPAD, "HgrPort_1", HGRBEAM3, "Neutral", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                }
                else
                {
                    if (isOffsetApplied[1])
                        JointHelper.CreatePlanarJoint("-1", rightStructPort, HGRBEAM3, "EndCap", Plane.XY, Plane.NegativeXY, 0);
                    else
                        JointHelper.CreatePlanarSlotJoint("-1", rightStructPort, HGRBEAM3, "EndCap", Plane.XY, Plane.NegativeXY, Axis.X, 0, 0);

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

