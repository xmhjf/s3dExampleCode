//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Assy_GD_AN.cs
//   PipeHgrAssemblies2,Ingr.SP3D.Content.Support.Rules.Assy_GD_AN
//   Author       : Vinay
//   Creation Date: 09-Sep-2015
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   01-10-2015      Vinay   DI-CP-276996	Update HS_Assembly2 to use RichSteel
//   27-11-2015      Vinay   DI-CP-276798	Replace the use of any HS_Utility parts
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
    public class Assy_GD_AN : CustomSupportDefinition
    {
        //Constants
        private const string LEFT_SECTION = "LEFT_SECTION";
        private const string RIGHT_SECTION = "RIGHT_SECTION";
        private const string LEFT_WEB = "LEFT_WEB";
        private const string LEFT_FLANGE = "LEFT_FLANGE";
        private const string RIGHT_WEB = "RIGHT_WEB";
        private const string RIGHT_FLANGE = "RIGHT_FLANGE";

        double shoeWidth, restWidth, restLength, gap;
        string sectionOrPlate, sectionOrPlatePart;
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
                    BusinessObject supportPart = support.GetRelationship("OccAssyHasPart", "OccAssyHasPart_Part").TargetObjects[0];

                    shoeWidth = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrAssyShoeW", "SHOE_W")).PropValue;
                    sectionOrPlate = (string)((PropertyValueString)supportPart.GetPropertyValue("IJUAHSA_SecPlate", "SecPlate")).PropValue;
                    sectionOrPlatePart = (string)((PropertyValueString)supportPart.GetPropertyValue("IJUAHSA_SecPlatePart", "SecPlatePart")).PropValue;
                    restLength = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssyGD_AN", "REST_L")).PropValue;
                    restWidth = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssyGD_AN", "REST_W")).PropValue;
                    gap = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssyGD", "GAP")).PropValue;

                    if (sectionOrPlate == "Plate")
                    {
                        parts.Add(new PartInfo(LEFT_WEB, sectionOrPlatePart));
                        parts.Add(new PartInfo(LEFT_FLANGE, sectionOrPlatePart));
                        parts.Add(new PartInfo(RIGHT_WEB, sectionOrPlatePart));
                        parts.Add(new PartInfo(RIGHT_FLANGE, sectionOrPlatePart));
                    }
                    else
                    {
                        parts.Add(new PartInfo(LEFT_SECTION, sectionOrPlatePart));
                        parts.Add(new PartInfo(RIGHT_SECTION, sectionOrPlatePart));
                    }
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
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;

                double length = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Vertical);
                double plateThickness, webWidth;

                if (sectionOrPlate == "Plate")
                {
                    BusinessObject leftWebPart = componentDictionary[LEFT_WEB].GetRelationship("madeFrom", "part").TargetObjects[0];

                    plateThickness = (double)((PropertyValueDouble)leftWebPart.GetPropertyValue("IJUAhsThickness1", "Thickness1")).PropValue;
                    webWidth = restWidth - plateThickness;

                    componentDictionary[LEFT_WEB].SetPropertyValue(webWidth, "IJUAhsLength1", "Length1");
                    componentDictionary[LEFT_WEB].SetPropertyValue(restLength, "IJUAhsWidth1", "Width1");
                    componentDictionary[RIGHT_WEB].SetPropertyValue(webWidth, "IJUAhsLength1", "Length1");
                    componentDictionary[RIGHT_WEB].SetPropertyValue(restLength, "IJUAhsWidth1", "Width1");
                    componentDictionary[LEFT_FLANGE].SetPropertyValue(restWidth, "IJUAhsLength1", "Length1");
                    componentDictionary[LEFT_FLANGE].SetPropertyValue(restLength, "IJUAhsWidth1", "Width1");
                    componentDictionary[RIGHT_FLANGE].SetPropertyValue(restWidth, "IJUAhsLength1", "Length1");
                    componentDictionary[RIGHT_FLANGE].SetPropertyValue(restLength, "IJUAhsWidth1", "Width1");


                    //Add Connection for the end of the Angled Beam
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        JointHelper.CreateRigidJoint("-1", "Structure", LEFT_FLANGE, "Port1", Plane.XY, Plane.YZ, Axis.Y, Axis.Y, restLength / 2, -shoeWidth / 2 - gap, 0);
                    else
                        JointHelper.CreateRigidJoint("-1", "Route", LEFT_FLANGE, "Port1", Plane.XY, Plane.YZ, Axis.Y, Axis.Z, length - restLength / 2, 0, shoeWidth / 2 + gap);

                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        JointHelper.CreateRigidJoint("-1", "Structure", RIGHT_FLANGE, "Port1", Plane.XY, Plane.YZ, Axis.Y, Axis.NegativeY, restLength / 2, shoeWidth / 2 + gap, 0);
                    else
                        JointHelper.CreateRigidJoint("-1", "Route", RIGHT_FLANGE, "Port1", Plane.XY, Plane.YZ, Axis.Y, Axis.NegativeZ, length - restLength / 2, 0, -shoeWidth / 2 - gap);

                    JointHelper.CreateRigidJoint(LEFT_FLANGE, "Port1", LEFT_WEB, "Port1", Plane.YZ, Plane.YZ, Axis.Y, Axis.NegativeZ, 0, plateThickness + (webWidth / 2), plateThickness / 2);

                    JointHelper.CreateRigidJoint(RIGHT_FLANGE, "Port1", RIGHT_WEB, "Port1", Plane.YZ, Plane.YZ, Axis.Y, Axis.NegativeZ, 0, plateThickness + (webWidth / 2), plateThickness / 2);
                }
                else
                {
                    componentDictionary[LEFT_SECTION].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                    componentDictionary[LEFT_SECTION].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");

                    componentDictionary[RIGHT_SECTION].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                    componentDictionary[RIGHT_SECTION].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");

                    BusinessObject leftSectionPart = componentDictionary[LEFT_SECTION].GetRelationship("madeFrom", "part").TargetObjects[0];
                    CrossSection crosssection = (CrossSection)leftSectionPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];
                    double steelWidth = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructCrossSectionDimensions", "Width")).PropValue;
                    double steelDepth = (double)((PropertyValueDouble)crosssection.GetPropertyValue("IStructCrossSectionDimensions", "Depth")).PropValue;

                    componentDictionary[LEFT_SECTION].SetPropertyValue(restLength, "IJUAHgrOccLength", "Length");
                    componentDictionary[RIGHT_SECTION].SetPropertyValue(restLength, "IJUAHgrOccLength", "Length");

                    //Add Connection for the end of the Angled Beam
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        JointHelper.CreateRigidJoint("-1", "Structure", LEFT_SECTION, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, steelWidth / 2, -steelDepth - shoeWidth / 2 - gap);
                    else
                        JointHelper.CreateRigidJoint("-1", "Route", LEFT_SECTION, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.NegativeX, length - restLength, steelDepth + shoeWidth / 2 + gap, steelWidth / 2);

                    //Add Connection for the end of the Angled Beam
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        JointHelper.CreateRigidJoint("-1", "Structure", RIGHT_SECTION, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, -steelWidth / 2, steelDepth + shoeWidth / 2 + gap);
                    else
                        JointHelper.CreateRigidJoint("-1", "Route", RIGHT_SECTION, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.X, length - restLength, -steelDepth - shoeWidth / 2 - gap, -steelWidth / 2);
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
                    if (sectionOrPlate == "Plate")
                        routeConnections.Add(new ConnectionInfo(LEFT_WEB, 1)); // partindex, routeindex
                    else
                        routeConnections.Add(new ConnectionInfo(LEFT_SECTION, 1));

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
                    if (sectionOrPlate == "Plate")
                        structConnections.Add(new ConnectionInfo(LEFT_WEB, 1)); // partindex, routeindex
                    else
                        structConnections.Add(new ConnectionInfo(LEFT_SECTION, 1));

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

