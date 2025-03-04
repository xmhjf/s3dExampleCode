//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Conn_Obj_Wrapper.cs
//   FINL_Assy,Ingr.SP3D.Content.Support.Rules.Conn_Obj_Wrapper
//   Author       :  BS
//   Creation Date:  14/06/2013
//   Description:    

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  14/06/2013    BS  CR-CP-224472-Initial Creation
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Support.Middle;

namespace Ingr.SP3D.Content.Support.Rules
{
    public class Conn_Obj_Wrapper : CustomSupportDefinition
    {
        private const string CONNOBJECT = "Conn_Obj";
        private const string ROUTECONNOBJECT = "RouteConnObj";        
        public Boolean Override { get; set; }
        public int Index { get; set; }
        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();
                    parts.Add(new PartInfo(CONNOBJECT + "_" + Index, "Log_Conn_Part_1"));
                    parts.Add(new PartInfo(ROUTECONNOBJECT + "_" + Index, "Log_Conn_Part_1"));

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
        //Get Route Connections
        //-----------------------------------------------------------------------------------
        public override Collection<ConnectionInfo> SupportedConnections
        {
            get
            {
                try
                {
                    Collection<ConnectionInfo> routeConnections = new Collection<ConnectionInfo>();
                    routeConnections.Add(new ConnectionInfo(ROUTECONNOBJECT + "_" + Index, 1)); // partindex, routeindex

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
                    structConnections.Add(new ConnectionInfo(ROUTECONNOBJECT + "_" + Index, 1)); // partindex, structindex

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

        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                if (Override == false)
                    JointHelper.CreateRigidJoint("-1", "Route", ROUTECONNOBJECT + "_" + Index, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                JointHelper.CreateRigidJoint(ROUTECONNOBJECT + "_" + Index, "Connection", CONNOBJECT + "_" + Index, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in Get Assembly Joints Method of Conn_Obj_Wrapper.cs." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }

    }
}
