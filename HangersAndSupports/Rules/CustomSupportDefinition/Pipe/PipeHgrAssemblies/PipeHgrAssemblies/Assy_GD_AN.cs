//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Assy_GD_AN.cs
//   PipeHgrAssemblies,Ingr.SP3D.Content.Support.Rules.Assy_GD_AN
//   Author       :Vijaya
//   Creation Date:05.Apr.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  05.Apr.2013     Vijaya   CR-CP-224484-Initial Creation
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
        string sectionSize, plateSize;
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

                    shoeWidth = (double)((PropertyValueDouble)support.GetPropertyValue("IJOAHgrAssyShoeW", "SHOE_W")).PropValue;
                    sectionSize = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrWTSize", "WTSize")).PropValue;
                    plateSize = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrAssyPlateSize", "PLATE_SIZE")).PropValue;
                    restLength = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssyGD_AN", "REST_L")).PropValue;
                    restWidth = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssyGD_AN", "REST_W")).PropValue;
                    gap = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssyGD", "GAP")).PropValue;

                    if (sectionSize == "Plate")
                    {
                        parts.Add(new PartInfo(LEFT_WEB, "Utility_PLATE_" + plateSize));
                        parts.Add(new PartInfo(LEFT_FLANGE, "Utility_PLATE_" + plateSize));
                        parts.Add(new PartInfo(RIGHT_WEB, "Utility_PLATE_" + plateSize));
                        parts.Add(new PartInfo(RIGHT_FLANGE, "Utility_PLATE_" + plateSize));
                    }
                    else
                    {
                        parts.Add(new PartInfo(LEFT_SECTION, sectionSize));
                        parts.Add(new PartInfo(RIGHT_SECTION, sectionSize));
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

                if (sectionSize == "Plate")
                {
                    BusinessObject leftWebPart = componentDictionary[LEFT_WEB].GetRelationship("madeFrom", "part").TargetObjects[0];

                    plateThickness = (double)((PropertyValueDouble)leftWebPart.GetPropertyValue("IJUAHgrUtility_PLATE", "THICKNESS")).PropValue;
                    webWidth = restWidth - plateThickness;

                    componentDictionary[LEFT_WEB].SetPropertyValue(webWidth, "IJOAHgrUtility_PLATE", "WIDTH");
                    componentDictionary[LEFT_WEB].SetPropertyValue(restLength, "IJOAHgrUtility_PLATE", "DEPTH");
                    componentDictionary[RIGHT_WEB].SetPropertyValue(webWidth, "IJOAHgrUtility_PLATE", "WIDTH");
                    componentDictionary[RIGHT_WEB].SetPropertyValue(restLength, "IJOAHgrUtility_PLATE", "DEPTH");
                    componentDictionary[LEFT_FLANGE].SetPropertyValue(restWidth, "IJOAHgrUtility_PLATE", "WIDTH");
                    componentDictionary[LEFT_FLANGE].SetPropertyValue(restLength, "IJOAHgrUtility_PLATE", "DEPTH");
                    componentDictionary[RIGHT_FLANGE].SetPropertyValue(restWidth, "IJOAHgrUtility_PLATE", "WIDTH");
                    componentDictionary[RIGHT_FLANGE].SetPropertyValue(restLength, "IJOAHgrUtility_PLATE", "DEPTH");

                    //Add Connection for the end of the Angled Beam
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        JointHelper.CreateRigidJoint("-1", "Structure", LEFT_FLANGE, "TopStructure", Plane.XY, Plane.YZ, Axis.Y, Axis.Y, restLength / 2, -shoeWidth / 2 - gap, 0);
                    else
                        JointHelper.CreateRigidJoint("-1", "Route", LEFT_FLANGE, "TopStructure", Plane.XY, Plane.YZ, Axis.Y, Axis.Z, length - restLength / 2, 0, shoeWidth / 2 + gap);

                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                        JointHelper.CreateRigidJoint("-1", "Structure", RIGHT_FLANGE, "TopStructure", Plane.XY, Plane.YZ, Axis.Y, Axis.NegativeY, restLength / 2, shoeWidth / 2 + gap, 0);
                    else
                        JointHelper.CreateRigidJoint("-1", "Route", RIGHT_FLANGE, "TopStructure", Plane.XY, Plane.YZ, Axis.Y, Axis.NegativeZ, length - restLength / 2, 0, -shoeWidth / 2 - gap);

                    JointHelper.CreateRigidJoint(LEFT_FLANGE, "TopStructure", LEFT_WEB, "TopStructure", Plane.YZ, Plane.YZ, Axis.Y, Axis.NegativeZ, 0, plateThickness + (webWidth / 2), plateThickness / 2);

                    JointHelper.CreateRigidJoint(RIGHT_FLANGE, "TopStructure", RIGHT_WEB, "TopStructure", Plane.YZ, Plane.YZ, Axis.Y, Axis.NegativeZ, 0, plateThickness + (webWidth / 2), plateThickness / 2);
                }
                else
                {
                    PropertyValueCodelist beginMiterCodelist = (PropertyValueCodelist)componentDictionary[LEFT_SECTION].GetPropertyValue("IJUAHgrMiterOcc", "BeginMiter");
                    if (beginMiterCodelist.PropValue == -1)
                        beginMiterCodelist.PropValue = 1;
                    PropertyValueCodelist endMiterCodelist = (PropertyValueCodelist)componentDictionary[LEFT_SECTION].GetPropertyValue("IJUAHgrMiterOcc", "EndMiter");
                    if (endMiterCodelist.PropValue == -1)
                        endMiterCodelist.PropValue = 1;

                    componentDictionary[LEFT_SECTION].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                    componentDictionary[LEFT_SECTION].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                    componentDictionary[LEFT_SECTION].SetPropertyValue(0.0, "IJUAHgrOccStructuralPart", "Orientation");
                    componentDictionary[LEFT_SECTION].SetPropertyValue(beginMiterCodelist.PropValue, "IJUAHgrMiterOcc", "BeginMiter");
                    componentDictionary[LEFT_SECTION].SetPropertyValue(endMiterCodelist.PropValue, "IJUAHgrMiterOcc", "EndMiter");

                    componentDictionary[RIGHT_SECTION].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                    componentDictionary[RIGHT_SECTION].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                    componentDictionary[RIGHT_SECTION].SetPropertyValue(0.0, "IJUAHgrOccStructuralPart", "Orientation");
                    componentDictionary[RIGHT_SECTION].SetPropertyValue(beginMiterCodelist.PropValue, "IJUAHgrMiterOcc", "BeginMiter");
                    componentDictionary[RIGHT_SECTION].SetPropertyValue(endMiterCodelist.PropValue, "IJUAHgrMiterOcc", "EndMiter");

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
                    if (sectionSize == "Plate")
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
                    if (sectionSize == "Plate")
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

