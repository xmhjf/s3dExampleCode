//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Assy_GD_T1.cs
//   PipeHgrAssemblies,Ingr.SP3D.Content.Support.Rules.Assy_GD_T1
//   Author       :Vijaya
//   Creation Date:08.Apr.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  08.Apr.2013     Vijaya   CR-CP-224484-Initial Creation
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

    public class Assy_GD_T1 : CustomSupportDefinition
    {
        //Constants
        private const string LEFT = "LEFT";
        private const string RIGHT = "RIGHT"; 
      
        double length = 0,gap,shoeWidth;
        string lSize;

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

                    length = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssyGD", "L")).PropValue;

                    lSize = (string)((PropertyValueString)support.GetPropertyValue("IJUAHgrWTSize", "WTSize")).PropValue;

                    gap = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssyGD", "GAP")).PropValue;

                    parts.Add(new PartInfo(LEFT, lSize));
                    parts.Add(new PartInfo(RIGHT, lSize));

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

                PropertyValueCodelist beginMiterCodelist = (PropertyValueCodelist)componentDictionary[LEFT].GetPropertyValue("IJUAHgrMiterOcc", "BeginMiter");
                if (beginMiterCodelist.PropValue == -1)
                    beginMiterCodelist.PropValue = 1;
                PropertyValueCodelist endMiterCodelist = (PropertyValueCodelist)componentDictionary[LEFT].GetPropertyValue("IJUAHgrMiterOcc", "EndMiter");
                if (endMiterCodelist.PropValue == -1)
                    endMiterCodelist.PropValue = 1;

                componentDictionary[LEFT].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                componentDictionary[LEFT].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                componentDictionary[LEFT].SetPropertyValue(0.0, "IJUAHgrOccStructuralPart", "Orientation");
                componentDictionary[LEFT].SetPropertyValue(beginMiterCodelist.PropValue, "IJUAHgrMiterOcc", "BeginMiter");
                componentDictionary[LEFT].SetPropertyValue(endMiterCodelist.PropValue, "IJUAHgrMiterOcc", "EndMiter");

                componentDictionary[RIGHT].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "EndOverLength");
                componentDictionary[RIGHT].SetPropertyValue(0.0, "IJUAHgrOccOverLength", "BeginOverLength");
                componentDictionary[RIGHT].SetPropertyValue(0.0, "IJUAHgrOccStructuralPart", "Orientation");
                componentDictionary[RIGHT].SetPropertyValue(beginMiterCodelist.PropValue, "IJUAHgrMiterOcc", "BeginMiter");
                componentDictionary[RIGHT].SetPropertyValue(endMiterCodelist.PropValue, "IJUAHgrMiterOcc", "EndMiter");

                //Create Joints
                double portDistance = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Vertical);

                componentDictionary[LEFT].SetPropertyValue(length, "IJUAHgrOccLength", "Length");
                componentDictionary[RIGHT].SetPropertyValue(length, "IJUAHgrOccLength", "Length");

                BusinessObject leftPart = componentDictionary[LEFT].GetRelationship("madeFrom", "part").TargetObjects[0];
                CrossSection crossSection = (CrossSection)leftPart.GetRelationship("HgrCrossSection", "CrossSection").TargetObjects[0];

                double steelWidth = (double)((PropertyValueDouble)crossSection.GetPropertyValue("IStructCrossSectionDimensions", "Width")).PropValue;
                double steelDepth = (double)((PropertyValueDouble)crossSection.GetPropertyValue("IStructCrossSectionDimensions", "Depth")).PropValue;

                //Add Connection for the end of the Angled Beam
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    JointHelper.CreateRigidJoint("-1", "Structure", LEFT, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, -steelWidth / 2, shoeWidth / 2 + gap + steelDepth);         
                else
                    JointHelper.CreateRigidJoint("-1", "Route", LEFT, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.X, portDistance - length, -shoeWidth / 2 - gap - steelDepth, -steelWidth / 2);

                //Add Connection for the end of the Angled Beam
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    JointHelper.CreateRigidJoint("-1", "Structure", RIGHT, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, steelWidth / 2, -shoeWidth / 2 - gap - steelDepth);           
                else
                    JointHelper.CreateRigidJoint("-1", "Route", RIGHT, "BeginCap", Plane.XY, Plane.XY, Axis.X, Axis.NegativeX, portDistance - length, shoeWidth / 2 + gap + steelDepth, steelWidth / 2);
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
                    routeConnections.Add(new ConnectionInfo(LEFT, 1)); // partindex, routeindex

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
                    structConnections.Add(new ConnectionInfo(LEFT, 1)); // partindex, routeindex

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

