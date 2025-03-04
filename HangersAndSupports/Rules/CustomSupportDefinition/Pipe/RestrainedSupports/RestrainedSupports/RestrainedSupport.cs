//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   RestrainedSupport.cs
//   RestrainedSupports,Ingr.SP3D.Content.Support.Rules.RestrainedSupport
//   Author       :  BS
//   Creation Date:  13.Dec.2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   27.Dec.2012     BS      CR224473 .Net RestrainedSupport Projected Creation                 
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
    //It is recommended that customers specify namespace of their Rules to be
    //<CompanyName>.SP3D.Content.<Specialization>.
    //It is also recommended that if customers want to change this Rule to suit their
    //requirements, they should change namespace/Rule name so the identity of the modified
    //Rule will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------
    public class RestrainedSupport : CustomSupportDefinition
    {
        //Constants
        private const string SymbolicRestraint = "SymbolicRestraint";


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
                    BoundingBoxHelper.CreateStandardBoundingBoxes(false);
                    parts.Add(new PartInfo(SymbolicRestraint, "SymbolicRestraint"));
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
                return 0;
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
                SupportComponent symbolicRestraint = componentDictionary[SymbolicRestraint];


                BoundingBox boundingBox;
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    boundingBox= BoundingBoxHelper.GetBoundingBox(BoundingBoxType.SupportedAndSupporting);
                }
                else
                {
                    boundingBox = BoundingBoxHelper.GetBoundingBox(BoundingBoxType.Supported);
                }
                Double width,height,routeAngle;
                width = boundingBox.Width;
                height = boundingBox.Height;
                routeAngle = 0;
                String BBox_Low = boundingBox.LowReferencePortName;
                
                symbolicRestraint.SetPropertyValue(width, "IJUAHgrOccGeometryForSymbolic", "Width1");
                symbolicRestraint.SetPropertyValue(height, "IJUAHgrOccGeometryForSymbolic", "Height1");
                symbolicRestraint.SetPropertyValue(routeAngle, "IJUAHgrOccGeometryForSymbolic", "RouteAngle");

                Matrix4X4 lcsBBox_Low = RefPortHelper.PortLCS(BBox_Low);
                symbolicRestraint.SetPropertyValue(lcsBBox_Low.GetIndexValue(12), "IJUAHgrOccGeometryForSymbolic", "RoutePortX");
                symbolicRestraint.SetPropertyValue(lcsBBox_Low.GetIndexValue(13), "IJUAHgrOccGeometryForSymbolic", "RoutePortY");
                //Create the Joint between the Bounding Box Low and Struct Reference Ports
                if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                {
                    JointHelper.CreateRigidJoint("-1", "BBSR_Low", SymbolicRestraint, "Structure", Plane.YZ, Plane.XY, Axis.Y, Axis.X,0,0,0);
                }
                else
                {
                    JointHelper.CreateRigidJoint("-1", "BBR_Low", SymbolicRestraint, "Structure", Plane.YZ, Plane.XY, Axis.Y, Axis.X, 0, 0, 0);
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
                    Collection<ConnectionInfo> routeConnections = new Collection<ConnectionInfo>();
                    for (int iIndex = 1; iIndex <= SupportHelper.SupportedObjects.Count; iIndex++)
                    {
                        routeConnections.Add(new ConnectionInfo(SymbolicRestraint, iIndex));
                    }
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
                    Collection<ConnectionInfo> structConnections = new Collection<ConnectionInfo>();
                    if (SupportHelper.PlacementType  == PlacementType.PlaceByStruct)
                    {                        
                        structConnections.Add(new ConnectionInfo(SymbolicRestraint, 1));                        
                    }
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



