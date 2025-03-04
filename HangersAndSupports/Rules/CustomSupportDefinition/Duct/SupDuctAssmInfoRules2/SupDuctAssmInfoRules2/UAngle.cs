//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   UAngle.cs
//   SupDuctAssmInfoRules2,Ingr.SP3D.Content.Support.Rules.UAngle
//   Author       : Vinay
//   Creation Date:  27/11/2015
//   Description:    

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------

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
    public class UAngle : CustomSupportDefinition
    {
        //Constants
        private const string HGRBEAM1 = "HgrBeam_1"; //Hanger beam object
        private const string HGRBEAM2 = "HgrBeam_2"; //Hanger beam object
        private const string HGRBEAM3 = "HgrBeam_3"; //Hanger beam object
        private const string HGRBEAM4 = "HgrBeam_4"; //Hanger beam object as angle guide
        private const string HGRBEAM5 = "HgrBeam_5"; //Hanger beam object as angle guide
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
                    parts.Add(new PartInfo(HGRBEAM3, sectionPartClass, "PartSelectionRule,Ingr.SP3D.Content.Support.Rules.CPartByCrossSection"));
                    parts.Add(new PartInfo(HGRBEAM4, sectionPartClass, "HgrSupSecondaryCS"));
                    parts.Add(new PartInfo(HGRBEAM5, sectionPartClass, "HgrSupSecondaryCS"));

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
                bool[] isOffsetApplied = SupDuctAssemblyServices.GetIsLugEndOffsetApplied(this);

                //===========================
                //3.Get the input objects information
                //   Dimensions of the pipe/duct cross section
                //   (Applicable for Pipe/Duct/Cableway)
                //==========================

                //===================================
                //4.Set symbol occurrence attributes (optional)
                //===================================
                double pcsSectionWidth = 0, pcsSectionDepth = 0;
                double scsSectionWidth = 0, scsSectionDepth = 0;
                for (int i = 1; i <= 5; i++)
                {
                    BusinessObject sectionPart = componentDictionary["HgrBeam_" + i].GetRelationship("madeFrom", "part").TargetObjects[0];
                    CrossSection crossSection = (CrossSection)sectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];

                    if (i == 1)
                    {
                        pcsSectionWidth = crossSection.Width;
                        pcsSectionDepth = crossSection.Depth;
                    }
                    else if (i == 4)
                    {
                        scsSectionWidth = crossSection.Width;
                        scsSectionDepth = crossSection.Depth;
                    }

                    //Set the length of guide angle
                    if (i == 4 || i == 5)
                        componentDictionary["HgrBeam_" + i].SetPropertyValue(2.0 * pcsSectionDepth, "IJUAHgrOccLength", "Length");
                }

                lugOffset[0] = D + scsSectionWidth;
                lugOffset[1] = lugOffset[0];

                //====================
                //    Get Indexed structural reference ports
                //    This method will differentiate "Left" structure input from "Right"
                //    structur input.
                //    "Left" structure is always on the negtive Y direction of the bounding box coord. sys
                //    "Right" structure is on the positive Y direction of the bounding box coord. sys.
                //====================

                string[] structPort = SupDuctAssemblyServices.GetIndexedStructPortName(this, isOffsetApplied);
                string leftStructPort = structPort[0], rightStructPort = structPort[1];

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
                // Create Joints
                //====== ======
                //Create a collection to hold the joints
                //----------------------------------------------------
                //Create the Joint between the RteLow Reference Port and the Bottom Symbol
                //Bounding box --- Horizontal Beam

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    JointHelper.CreatePrismaticJoint("-1", "BBSR_Low", HGRBEAM2, "BeginCap", Plane.XY, Plane.YZ, Axis.X, Axis.Y, boundingBoxHeight + G, -lugOffset[0]);
                else
                {
                    if (isOffsetApplied[0])
                        JointHelper.CreateRigidJoint("-1", "BBR_Low", HGRBEAM2, "BeginCap", Plane.XY, Plane.YZ, Axis.X, Axis.Y, boundingBoxHeight + G, -lugOffset[0], 0);
                    else
                        JointHelper.CreatePrismaticJoint("-1", "BBR_Low", HGRBEAM2, "BeginCap", Plane.XY, Plane.YZ, Axis.Y, Axis.Z, boundingBoxHeight + G, 0);
                }

                //----------------------------------------------------
                //Create the Plane Joint between the RteHigh Reference Port
                //and the Right Bottom Symbol

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    JointHelper.CreatePlanarJoint("-1", "BBSR_High", HGRBEAM2, "EndCap", Plane.ZX, Plane.XY, lugOffset[1]);
                else
                {
                    if (isOffsetApplied[1])
                        JointHelper.CreatePlanarJoint("-1", "BBR_High", HGRBEAM2, "EndCap", Plane.ZX, Plane.XY, lugOffset[1]);
                }

                //----------------------------------------------------
                //Add a Prismatic Joint defining the flexible bottom member
                JointHelper.CreatePrismaticJoint(HGRBEAM2, "BeginCap", HGRBEAM2, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                //----------------------------------------------------
                //Add rigid joint between beam 1 and beam 2
                JointHelper.CreateRigidJoint(HGRBEAM1, "EndCap", HGRBEAM2, "BeginCap", Plane.XY, Plane.NegativeYZ, Axis.X, Axis.Z, 0, 0, 0);
                //Add rigid joint between beam 2 and beam 3
                JointHelper.CreateRigidJoint(HGRBEAM3, "BeginCap", HGRBEAM2, "EndCap", Plane.XY, Plane.YZ, Axis.X, Axis.NegativeZ, 0, 0, 0);

                //----------------------------------------------------
                //Add prismatic joint for the flexible memeber
                JointHelper.CreatePrismaticJoint(HGRBEAM1, "BeginCap", HGRBEAM1, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                //Add prismatic joint for the flexible memeber
                JointHelper.CreatePrismaticJoint(HGRBEAM3, "BeginCap", HGRBEAM3, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                //structure -- leftPad -- Profile part 1
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    if (leftPad)
                    {
                        JointHelper.CreatePointOnAxisJoint(LEFTPAD, "Port1", "-1", leftStructPort, Axis.X);
                        JointHelper.CreateRigidJoint(LEFTPAD, "Port2", HGRBEAM1, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, -pcsSectionDepth / 2.0, -pcsSectionWidth / 2.0);
                    }
                    else
                        JointHelper.CreatePointOnAxisJoint(HGRBEAM1, "BeginCap", "-1", leftStructPort, Axis.X);
                }
                else
                {
                    if (leftPad)
                    {
                        if (isOffsetApplied[0])
                            JointHelper.CreatePointOnPlaneJoint(LEFTPAD, "Port1", "-1", leftStructPort, Plane.XY);
                        else
                            JointHelper.CreatePointOnAxisJoint(LEFTPAD, "Port1", "-1", leftStructPort, Axis.X);
                        JointHelper.CreateRigidJoint(LEFTPAD, "Port2", HGRBEAM1, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, -pcsSectionDepth / 2.0, -pcsSectionWidth / 2.0);
                    }
                    else
                    {
                        if (isOffsetApplied[0])
                            JointHelper.CreatePointOnPlaneJoint(HGRBEAM1, "BeginCap", "-1", leftStructPort, Plane.XY);
                        else
                            JointHelper.CreatePointOnAxisJoint(HGRBEAM1, "BeginCap", "-1", leftStructPort, Axis.X);
                    }
                }

                //structure -- rightPad -- profile part 3
                if (rightPad)
                {
                    if (isOffsetApplied[1])
                        JointHelper.CreatePointOnPlaneJoint(RIGHTPAD, "Port2", "-1", rightStructPort, Plane.XY);
                    else
                        JointHelper.CreatePointOnAxisJoint(RIGHTPAD, "Port2", "-1", rightStructPort, Axis.X);

                    JointHelper.CreateRigidJoint(RIGHTPAD, "Port1", HGRBEAM3, "EndCap", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, -pcsSectionDepth / 2.0, -pcsSectionWidth / 2.0);
                }
                else
                {
                    if (isOffsetApplied[1])
                        JointHelper.CreatePointOnPlaneJoint(HGRBEAM3, "EndCap", "-1", rightStructPort, Plane.XY);
                    else
                        JointHelper.CreatePointOnAxisJoint(HGRBEAM3, "EndCap", "-1", rightStructPort, Axis.X);
                }

                //structure support part --- angle ---- duct
                JointHelper.CreatePrismaticJoint(HGRBEAM2, "EndCap", HGRBEAM4, "EndCap", Plane.YZ, Plane.NegativeZX, Axis.Z, Axis.X, 0, -pcsSectionDepth / 2.0);
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    JointHelper.CreatePlanarJoint("-1", "BBSR_High", HGRBEAM4, "EndCap", Plane.ZX, Plane.YZ, 0);
                else
                    JointHelper.CreatePlanarJoint("-1", "BBR_High", HGRBEAM4, "EndCap", Plane.ZX, Plane.YZ, 0);

                //structure support part --- angle ---- duct
                JointHelper.CreatePrismaticJoint(HGRBEAM2, "EndCap", HGRBEAM5, "BeginCap", Plane.YZ, Plane.NegativeZX, Axis.Z, Axis.NegativeX, 0, -pcsSectionDepth / 2.0);
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    JointHelper.CreatePlanarJoint("-1", "BBSR_Low", HGRBEAM5, "BeginCap", Plane.ZX, Plane.NegativeYZ, 0);
                else
                    JointHelper.CreatePlanarJoint("-1", "BBR_Low", HGRBEAM5, "BeginCap", Plane.ZX, Plane.NegativeYZ, 0);
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

                    routeConnections.Add(new ConnectionInfo(HGRBEAM4, 1)); // partindex, routeindex
                    routeConnections.Add(new ConnectionInfo(HGRBEAM5, 1)); // partindex, routeindex

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
                        structConnections.Add(new ConnectionInfo(RIGHTPAD, 1));// partindex, structureindex
                    else
                        structConnections.Add(new ConnectionInfo(HGRBEAM3, 1));// partindex, structureindex
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

