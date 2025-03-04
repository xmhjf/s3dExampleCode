//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Assy_GD_HD.cs
//   PipeHgrAssemblies,Ingr.SP3D.Content.Support.Rules.Assy_GD_HD
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

    public class Assy_GD_HD : CustomSupportDefinition
    {
        //Constants
        private const string LEFT = "LEFT";
        private const string RIGHT = "RIGHT"; 

        double shoeWidth,length = 0, width, height, gap;

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

                    width = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssyGD_HD", "W")).PropValue;

                    height = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssyGD_HD", "H")).PropValue;

                    gap = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssyGD", "GAP")).PropValue;

                    parts.Add(new PartInfo(LEFT, "Utility_USER_FIXED_BOX_1"));
                    parts.Add(new PartInfo(RIGHT, "Utility_USER_FIXED_BOX_1"));

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
                double portDistance = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Vertical);

                componentDictionary[LEFT].SetPropertyValue(height, "IJOAHgrUtility_USER_FIXED_BOX", "L");
                componentDictionary[LEFT].SetPropertyValue(width, "IJOAHgrUtility_USER_FIXED_BOX", "WIDTH");
                componentDictionary[LEFT].SetPropertyValue(length, "IJOAHgrUtility_USER_FIXED_BOX", "DEPTH");

                componentDictionary[RIGHT].SetPropertyValue(height, "IJOAHgrUtility_USER_FIXED_BOX", "L");
                componentDictionary[RIGHT].SetPropertyValue(width, "IJOAHgrUtility_USER_FIXED_BOX", "WIDTH");
                componentDictionary[RIGHT].SetPropertyValue(length, "IJOAHgrUtility_USER_FIXED_BOX", "DEPTH");

                //Create Joints
                //Add Connection for the end of the Angled Beam
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    JointHelper.CreateRigidJoint("-1", "Structure", LEFT, "StartOther", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, -shoeWidth / 2 - gap);        
                else
                    JointHelper.CreateRigidJoint("-1", "Route", LEFT, "StartOther", Plane.XY, Plane.XY, Axis.X, Axis.X, portDistance - height, -shoeWidth / 2 - gap, 0);

                //Add Connection for the end of the Angled Beam
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    JointHelper.CreateRigidJoint("-1", "Structure", RIGHT, "StartOther", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, shoeWidth / 2 + gap);    
                else
                    JointHelper.CreateRigidJoint("-1", "Route", RIGHT, "StartOther", Plane.XY, Plane.XY, Axis.X, Axis.X, portDistance - height, shoeWidth / 2 + gap, 0);
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

