//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   AngularAngle.cs
//   SupDuctAssmInfoRules2,Ingr.SP3D.Content.Support.Rules.AngularAngle
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
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;
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
    public class AngularAngle : CustomSupportDefinition
    {

        bool leftPad = false, rightPad = false;
        double D = 0, G = 0;
        //Constants
        private const string HGRBEAM1 = "HGRBEAM1"; //Horizontal structural part
        private const string CONNECTION1 = "CONNECTION1";    //Logical connection
        private const string HGRBEAM2 = "HGRBEAM2";          //supporting strucutral part
        private const string CONNECTION2 = "CONNECTION2";    //Logical connection
        private const string HGRBEAM3 = "HGRBEAM3";      //Hanger beam object as angle guide
        private const string HGRBEAM4 = "HGRBEAM4";      //Hanger beam object as angle guide
        private const string CONNECTION3 = "CONNECTION3";   //bounds to structure, make it easy for toggle
        private const string CONNECTION4 = "CONNECTION4";   //bounds to BBSR_Low/BBR_Low, make it easy for toggle
        private const string CONNECTION5 = "CONNECTION5";   //bounds to BBSR_High/BBR_High, make it easy for toggle
        private const string CONNECTION6 = "CONNECTION6";  //connection for HgrBeam(3) to pad/structure
        string[] Dimensionkeys = new string[5];

        private const string LEFTPAD = "LEFTPAD";
        private const string RIGHTPAD = "RIGHTPAD";
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
                    parts.Add(new PartInfo(CONNECTION2, "Connection"));
                    parts.Add(new PartInfo(CONNECTION1, "Connection"));
                    parts.Add(new PartInfo(HGRBEAM3, sectionPartClass, "HgrSupSecondaryCS"));
                    parts.Add(new PartInfo(HGRBEAM4, sectionPartClass, "HgrSupSecondaryCS"));
                    parts.Add(new PartInfo(CONNECTION3, "Connection"));
                    parts.Add(new PartInfo(CONNECTION4, "Connection"));
                    parts.Add(new PartInfo(CONNECTION5, "Connection"));
                    parts.Add(new PartInfo(CONNECTION6, "Connection"));

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
                bool[] isOffsetApplied = SupDuctAssemblyServices.GetIsLugEndOffsetApplied(this);
                //====================================
                //3. Retrieve route/structure configuration and
                //check if offset value is to be applied on lug end
                //If the leg is attached to a stiffener, no offset needs to be specified
                //If the leg is attached to a slab, the offest value will be retrieved from
                //predefined offset rule
                //====================================
                Collection<object> colllection = new Collection<object>();
                bool value = GenericHelper.GetDataByRule("HgrSupAngleByLF", support, out colllection);

                double steelDepth1, steelDepth2, steelWidth1, steelWidth2;
                BusinessObject horizontalSectionPart = componentDictionary[HGRBEAM1].GetRelationship("madeFrom", "part").TargetObjects[0];
                CrossSection crosssection = (CrossSection)horizontalSectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];

                BusinessObject SectionPart = componentDictionary[HGRBEAM3].GetRelationship("madeFrom", "part").TargetObjects[0];
                CrossSection crosssection1 = (CrossSection)SectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];

                //===========================
                //4. Get the input objects information
                //Dimensions of the pipe/duct cross section
                //(Applicable for Pipe/Duct/Cableway)
                //==========================
                steelWidth1 = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructCrossSectionDimensions", "Width")).PropValue;
                steelDepth1 = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructCrossSectionDimensions", "Depth")).PropValue;
                steelWidth2 = (double)((PropertyValueDouble)crosssection1.GetPropertyValue("IStructCrossSectionDimensions", "Width")).PropValue;
                steelDepth2 = (double)((PropertyValueDouble)crosssection1.GetPropertyValue("IStructCrossSectionDimensions", "Depth")).PropValue;
                double steelThickness = (double)((PropertyValueDouble)crosssection1.GetPropertyValue("IStructFlangedSectionDimensions", "tf")).PropValue;
                double lugOffset = D + steelWidth2;

                string BBX_Low, BBX_High, Structure = string.Empty, Structure_idx = string.Empty, Low_idx = string.Empty, High_idx = string.Empty;

                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    BBX_Low = "BBSR_Low";
                    BBX_High = "BBSR_High";
                }
                else
                {
                    BBX_Low = "BBR_Low";
                    BBX_High = "BBR_High";
                }

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
                //5. Create Joints
                //====== ======
                double lengthhor2 = 0;
                if (Configuration == 1)
                {
                    JointHelper.CreatePrismaticJoint("-1", "Structure", CONNECTION3, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.NegativeX, 0, 0);

                    JointHelper.CreateRigidJoint("-1", BBX_Low, CONNECTION4, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.NegativeX, 0, boundingBoxWidth, 0);

                    JointHelper.CreateRigidJoint("-1", BBX_High, CONNECTION5, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.NegativeX, 0, -boundingBoxWidth, 0);

                    lengthhor2 = RefPortHelper.DistanceBetweenPorts(BBX_Low, "Structure", PortDistanceType.Horizontal) + ((boundingBoxWidth / 2) + 1.5 * (steelWidth2));
                    Structure = "Connection";
                    Structure_idx = CONNECTION3;
                    BBX_Low = "Connection";
                    BBX_High = "Connection";
                    Low_idx = CONNECTION4;
                    High_idx = CONNECTION5;
                    
                }
                else
                {
                    lengthhor2 = RefPortHelper.DistanceBetweenPorts(BBX_Low, "Structure", PortDistanceType.Horizontal) +((boundingBoxWidth / 2) + 1.5 * (steelWidth2)) + steelWidth2 + (boundingBoxWidth / 2 - steelWidth2);
                    Structure = "Structure";
                    Structure_idx = "-1";
                    Low_idx = "-1";
                    High_idx = "-1";


                }
                (componentDictionary[HGRBEAM3]).SetPropertyValue(2 * steelDepth1, "IJUAHgrOccLength", "Length");
                (componentDictionary[HGRBEAM4]).SetPropertyValue(2 * steelDepth1, "IJUAHgrOccLength", "Length");

                //Add Notes to the support
                ControlPoint controlPoint;
                Note note1 = CreateNote("L_Start", HGRBEAM1, "BeginCap", new Position(0, 0, 0), " ", true, 2, 1, out controlPoint);
                note1.SetPropertyValue("", "IJGeneralNote", "Text");
                PropertyValueCodelist note1PropertyValueCL = (PropertyValueCodelist)note1.GetPropertyValue("IJGeneralNote", "Purpose");
                CodelistItem codeList1 = note1PropertyValueCL.PropertyInfo.CodeListInfo.GetCodelistItem((int)3);   //value 3 means fabrication
                note1.SetPropertyValue(codeList1, "IJGeneralNote", "Purpose");
                note1.SetPropertyValue(true, "IJGeneralNote", "Dimensioned");

                Note note2 = CreateNote("L_End", HGRBEAM2, "BeginCap", new Position(0, 0, 0), " ", true, 2, 1, out controlPoint);
                note2.SetPropertyValue("", "IJGeneralNote", "Text");
                PropertyValueCodelist note2PropertyValueCL = (PropertyValueCodelist)note2.GetPropertyValue("IJGeneralNote", "Purpose");
                CodelistItem codeList2 = note2PropertyValueCL.PropertyInfo.CodeListInfo.GetCodelistItem((int)3);   //value 3 means fabrication
                note2.SetPropertyValue(codeList2, "IJGeneralNote", "Purpose");
                note2.SetPropertyValue(true, "IJGeneralNote", "Dimensioned");

                Note note3 = CreateNote("L_1", "-1", "Route", new Position(0, 0, 0), " ", true, 2, 1, out controlPoint);
                note3.SetPropertyValue("", "IJGeneralNote", "Text");
                PropertyValueCodelist note3PropertyValueCL = (PropertyValueCodelist)note3.GetPropertyValue("IJGeneralNote", "Purpose");
                CodelistItem codeList3 = note3PropertyValueCL.PropertyInfo.CodeListInfo.GetCodelistItem((int)3);   //value 3 means fabrication
                note3.SetPropertyValue(codeList2, "IJGeneralNote", "Purpose");

                //Add a prismatic joint between structure and part 1
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    if (leftPad)
                    {
                        JointHelper.CreatePointOnAxisJoint(LEFTPAD, "Port1", Structure_idx, Structure, Axis.X);

                        JointHelper.CreateRigidJoint(LEFTPAD, "Port2", HGRBEAM1, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, -steelDepth1 / 2, -steelWidth1 / 2);
                    }
                    else
                        JointHelper.CreatePointOnAxisJoint(HGRBEAM1, "BeginCap", Structure_idx, Structure, Axis.X);
                }
                else
                {
                    if (leftPad)
                    {
                        JointHelper.CreatePointOnPlaneJoint(LEFTPAD, "Port1", Structure_idx, Structure, Plane.XY);

                        JointHelper.CreateRigidJoint(LEFTPAD, "Port2", HGRBEAM1, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, -steelDepth1 / 2, -steelWidth1 / 2);
                    }
                    else
                        JointHelper.CreatePointOnPlaneJoint(HGRBEAM1, "BeginCap", Structure_idx, Structure, Plane.XY);
                }
                //----------------------------------------
                //Route --- HgrBeam 1
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    JointHelper.CreatePlanarJoint(Low_idx, BBX_Low, HGRBEAM1, "EndCap", Plane.XY, Plane.NegativeXY, -lugOffset);

                    JointHelper.CreatePlanarJoint(High_idx, BBX_High, HGRBEAM1, "EndCap", Plane.ZX, Plane.ZX, G);
                }
                else
                    JointHelper.CreateRigidJoint(High_idx, BBX_High, HGRBEAM1, "EndCap", Plane.XY, Plane.NegativeXY, Axis.Y, Axis.Y, -lugOffset - boundingBoxHeight, 0, G);

                JointHelper.CreatePrismaticJoint(HGRBEAM1, "BeginCap", HGRBEAM1, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                if (leftPad)
                    JointHelper.CreatePrismaticJoint(CONNECTION1, "Connection", LEFTPAD, "Port1", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, -lugOffset * 5, 0);
                else
                    JointHelper.CreatePrismaticJoint(CONNECTION1, "Connection", HGRBEAM1, "BeginCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, -lugOffset * 5, 0);

                //Logical conection between horizontal beam and logical part to locate the second beam
                JointHelper.CreateRigidJoint(HGRBEAM1, "EndCap", CONNECTION2, "Connection", Plane.YZ, Plane.YZ, Axis.Z, Axis.Z, 0, 0, -lugOffset * 2);
                //Cylindrical joint between logical connection and beam 2
                if (rightPad)
                {
                    JointHelper.CreateRigidJoint(CONNECTION1, "Connection", RIGHTPAD, "Port1", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, -0.5 * steelDepth1, 0);

                    JointHelper.CreateRigidJoint(CONNECTION6, "Connection", RIGHTPAD, "Port2", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0.5 * steelDepth1, 0);

                    JointHelper.CreateCylindricalJoint(CONNECTION6, "Connection", HGRBEAM2, "BeginCap", Axis.X, Axis.X, 0);
                }
                else
                    JointHelper.CreateCylindricalJoint(CONNECTION1, "Connection", HGRBEAM2, "BeginCap", Axis.X, Axis.X, 0);

                JointHelper.CreatePointOnPlaneJoint(CONNECTION1, "Connection", Structure_idx, Structure, Plane.XY);
                //Spherical joint between logical connection and beam 2
                JointHelper.CreateSphericalJoint(CONNECTION2, "Connection", HGRBEAM2, "EndCap");
                //Flexible member
                JointHelper.CreatePrismaticJoint(HGRBEAM2, "BeginCap", HGRBEAM2, "EndCap", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);
                //----------------------------------------------------
                //---- profile part 1 <---> angle guide <--> duct----
                JointHelper.CreatePrismaticJoint(HGRBEAM1, "EndCap", HGRBEAM3, "BeginCap", Plane.ZX, Plane.NegativeZX, Axis.Z, Axis.X, 0, -0.5 * steelDepth1);

                JointHelper.CreatePlanarJoint(Low_idx, BBX_Low, HGRBEAM3, "EndCap", Plane.XY, Plane.NegativeYZ, 0);

                if (Configuration == 1)
                {
                    JointHelper.CreatePrismaticJoint(HGRBEAM1, "EndCap", HGRBEAM4, "EndCap", Plane.ZX, Plane.NegativeZX, Axis.Z, Axis.NegativeX, 0, -0.5 * steelDepth1);

                    JointHelper.CreatePlanarJoint(High_idx, BBX_Low, HGRBEAM4, "EndCap", Plane.XY, Plane.YZ, 0);
                }
                else
                {
                    JointHelper.CreatePrismaticJoint(HGRBEAM1, "EndCap", HGRBEAM4, "EndCap", Plane.ZX, Plane.NegativeZX, Axis.Z, Axis.NegativeX, 0, -0.5 * steelDepth1);

                    JointHelper.CreatePlanarJoint(High_idx, BBX_Low, HGRBEAM4, "EndCap", Plane.XY, Plane.YZ,boundingBoxHeight);
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

                    //Add the LeftAngleGuide-Route Connections to the Route Collection of Connections.
                    routeConnections.Add(new ConnectionInfo(HGRBEAM3, 1)); // Left angle guide, Connects to First Route Input
                    //Add the RightAngleGuide-Route Connections to the Route Collection of Connections.
                    routeConnections.Add(new ConnectionInfo(HGRBEAM4, 1)); // Right angle guide, Connects to First Route Input

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
