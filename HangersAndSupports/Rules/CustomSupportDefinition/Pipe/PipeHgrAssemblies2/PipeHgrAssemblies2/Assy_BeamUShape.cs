//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   BeamUShape.cs
//   PipeHgrAssemblies2,Ingr.SP3D.Content.Support.Rules.BeamUShape
//   Author       :Vinay
//   Creation Date:27-11-2015
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
                    object[] attributeCollection = SupportAssemblyServices.GetPipeStructuralASMAttributes(this);

                    //Get the attributes from assembly
                    d = (double)attributeCollection[0];
                    gap = (double)attributeCollection[1];
                    leftPad = (bool)attributeCollection[2];
                    rightPad = (bool)attributeCollection[3];

                    string sectionPartClass = (string)((PropertyValueString)support.SupportDefinition.GetPropertyValue("IJUAHSA_SecPartClass", "SecPartClass")).PropValue;
                 
                    //Use the default selection rule to get a catalog part for each part class
                    for (int index = 1; index <= 3; index++)
                        parts.Add(new PartInfo("HGRBEAM" + index, sectionPartClass, "PartSelectionRule,Ingr.SP3D.Content.Support.Rules.PartByLoadFactor"));

                    if (rightPad)
                        parts.Add(new PartInfo(RIGHTPAD, SupportAssemblyServices.GetPadPartNameByCrossSectionType(this, parts[0])));
                    if (leftPad)
                        parts.Add(new PartInfo(LEFTPAD, SupportAssemblyServices.GetPadPartNameByCrossSectionType(this, parts[0])));

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

                double[] boxOffset = SupportAssemblyServices.GetBoundaryObjectDimension(this, boundingBox);
                double[] lugoffset = new double[2];
                lugoffset[0] = d;
                lugoffset[1] = d;

                bool[] isOffsetApplied = SupportAssemblyServices.GetIsLugEndOffsetApplied(this);
                string[] structPort = SupportAssemblyServices.GetIndexedStructPortName(this, isOffsetApplied);
                string leftStructPort = structPort[0],rightStructPort = structPort[1];                

                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
                double beginOverLength, endOverLength, leftVerticalDistance=0, rightVerticalDistance=0;
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

                //Add notes
                Note note;
                ControlPoint controlPoint;
                leftVerticalDistance = RefPortHelper.DistanceBetweenPorts("BBRV_High", leftStructPort, PortDistanceType.Vertical);
                rightVerticalDistance = RefPortHelper.DistanceBetweenPorts("BBRV_High", rightStructPort, PortDistanceType.Vertical);
                note = CreateNote("L_Start", HGRBEAM2, "BeginCap", new Position(0, 0, 0), " ", true, 2, 1, out controlPoint);
                note = CreateNote("L_End", HGRBEAM2, "BeginCap", new Position(0, 0, boundingBoxWidth + beginOverLength + endOverLength + lugoffset[0] + lugoffset[1]), " ", true, 2, 1, out controlPoint);
                note = CreateNote("L_1", HGRBEAM2, "BeginCap", new Position(leftVerticalDistance - gap, 0, 0), " ", true, 2, 1, out controlPoint);
                note = CreateNote("L_2", HGRBEAM2, "BeginCap", new Position(rightVerticalDistance - gap, 0, boundingBoxWidth + beginOverLength + endOverLength + lugoffset[0] + lugoffset[1]), " ", true, 2, 1, out controlPoint);

                //Set the Pad Port Positions
                double rightPadLength = 0, rightPadWidth = 0, leftPadLength = 0, leftPadWidth = 0;
                Part padPart = null;
                string rightPadPartnumber = "", leftPadPartnumber = "";
                string rightPadtype = "", leftPadtype = "";
                if (rightPad)
                {
                    padPart = (Part)componentDictionary[RIGHTPAD].GetRelationship("madeFrom", "part").TargetObjects[0];
                    rightPadPartnumber = padPart.PartNumber;
                    rightPadtype = SupportAssemblyServices.GetPadPartType(this, rightPadPartnumber);
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
                    leftPadtype = SupportAssemblyServices.GetPadPartType(this, leftPadPartnumber);
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

                //Create Joints   
                JointHelper.CreateRigidJoint(HGRBEAM1, "EndCap", HGRBEAM2, "BeginCap", Plane.XY, Plane.NegativeYZ, Axis.X, Axis.Z, 0.0, 0.0, 0.0);
                JointHelper.CreateRigidJoint(HGRBEAM3, "BeginCap", HGRBEAM2, "EndCap", Plane.XY, Plane.YZ, Axis.X, Axis.NegativeZ, 0.0, 0.0, 0.0);
                JointHelper.CreatePrismaticJoint(HGRBEAM1, "BeginCap", HGRBEAM1, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0.0, 0.0);
                JointHelper.CreatePrismaticJoint(HGRBEAM3, "BeginCap", HGRBEAM3, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0.0, 0.0);

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

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    if (leftPad)
                    {
                        JointHelper.CreatePrismaticJoint("-1", leftStructPort, LEFTPAD, "Port1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0);
                        JointHelper.CreatePlanarJoint(LEFTPAD, "Port2", HGRBEAM1, "BeginCap", Plane.XY, Plane.XY, 0.0);
                        JointHelper.CreatePrismaticJoint(LEFTPAD, "Port2", HGRBEAM1, "Neutral", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                    }
                    else
                        JointHelper.CreatePrismaticJoint("-1", leftStructPort, HGRBEAM1, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.X, 0.0, 0.0);
                }
                else
                {
                    if (leftPad)
                    {
                        if (isOffsetApplied[1])
                            JointHelper.CreatePlanarJoint("-1", leftStructPort, LEFTPAD, "Port1", Plane.XY, Plane.XY, 0);
                        else
                            JointHelper.CreatePlanarSlotJoint("-1", leftStructPort, LEFTPAD, "Port1", Plane.XY, Plane.XY, Axis.X, 0, 0);

                        JointHelper.CreatePlanarJoint(LEFTPAD, "Port2", HGRBEAM1, "BeginCap", Plane.XY, Plane.XY, 0);
                        JointHelper.CreatePrismaticJoint(LEFTPAD, "Port2", HGRBEAM1, "Neutral", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
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
                        JointHelper.CreatePlanarJoint("-1", rightStructPort, RIGHTPAD, "Port2", Plane.XY, Plane.NegativeXY, 0);
                    else
                        JointHelper.CreatePointOnAxisJoint(RIGHTPAD, "Port1", "-1", rightStructPort, Axis.X);

                    JointHelper.CreatePlanarJoint(RIGHTPAD, "Port1", HGRBEAM3, "EndCap", Plane.XY, Plane.XY, 0);
                    JointHelper.CreatePrismaticJoint(RIGHTPAD, "Port1", HGRBEAM3, "Neutral", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
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

