//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Assy_GN_VR_CYL.cs
//   PipeHgrAssemblies,Ingr.SP3D.Content.Support.Rules.Assy_GN_VR_CYL
//   Author       :Vijaya
//   Creation Date:08.Apr.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  08.Apr.2013     Vijaya   CR-CP-224484-Initial Creation
//  22-02-2015       PVK    TR-CP-264951  Resolve coverity issues found in November 2014 report
//  28-04-2015       PR      TR 258572  Generic Variable Cylinder support places incorrectly on Metric pipe 
//  07-09-2016       PVK     TR-CP-301088	Improper Support Information Rule definition prevents placement
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.Generic;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Content.Support.Symbols;

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

    public class Assy_GN_VR_CYL : CustomSupportDefinition
    {
        //Constants
        private const string VARIABLE_CYL = "VARIABLE_CYL"; // 1
        double radius;

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

                    radius = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAHgrAssyGN", "RADIUS")).PropValue;

                    parts.Add(new PartInfo(VARIABLE_CYL, "Utility_USER_VARIABLE_CYL_1"));

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
                BoundingBoxHelper.CreateStandardBoundingBoxes(false);
                PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                NominalDiameter pipeDiameter = new NominalDiameter();
                pipeDiameter.Size = pipeInfo.NominalDiameter.Size;

                UnitName unitName = MiddleServiceProvider.UOMMgr.GetUnit(UnitType.Distance, pipeInfo.NominalDiameter.Units);

                pipeDiameter.Size = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Distance, pipeDiameter.Size, unitName, UnitName.DISTANCE_METER);    
                double offset = pipeDiameter.Size / 2;

                componentDictionary[VARIABLE_CYL].SetPropertyValue(radius, "IJOAHgrUtility_VARIABLE_CYL", "RADIUS");
                const double CONST_1 = 1.571;

                //Create Joints
                double angle = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, OrientationAlong.Global_Z);
                if ((HgrCompareDoubleService.cmpdbl(Math.Round(angle, 3), 0) == true ||HgrCompareDoubleService.cmpdbl(Math.Round(angle, 3), CONST_1) == true))  // check if pipes are vertical
                    JointHelper.CreateRigidJoint(VARIABLE_CYL, "StartOther", "-1", "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, -offset, 0, 0);   
                else
                    JointHelper.CreateRigidJoint(VARIABLE_CYL, "StartOther", "-1", "BBRV_High", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, pipeDiameter.Size / 2, 0);

                //Add Connection for the end of the Angled Beam
                JointHelper.CreatePrismaticJoint(VARIABLE_CYL, "StartOther", VARIABLE_CYL, "EndOther", Plane.ZX, Plane.ZX, Axis.Z, Axis.Z, 0, 0);

                //Connection to Structure                 
                JointHelper.CreatePointOnPlaneJoint(VARIABLE_CYL, "EndOther", "-1", "Structure", Plane.XY);
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
                    routeConnections.Add(new ConnectionInfo(VARIABLE_CYL, 1)); // partindex, routeindex

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
                    structConnections.Add(new ConnectionInfo(VARIABLE_CYL, 1)); // partindex, routeindex

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

