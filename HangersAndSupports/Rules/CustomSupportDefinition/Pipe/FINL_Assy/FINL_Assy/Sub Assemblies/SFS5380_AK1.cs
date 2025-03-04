//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SFS5380_AK1.cs
//   FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5380_AK1
//   Author       :  BS
//   Creation Date:  20.Jun.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   20-Jun-2013     BS   CR-CP-224491- Convert HS_FINL_SubAssy to C# .Net 
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Support.Middle;

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

    public class SFS5380_AK1 : CustomSupportDefinition
    {
        private const string PIPECLAMPAK1 = "PipeClamp_AK1";
        private const string EYENUTAK1 = "EyeNut_AK1";
        private const string CONNOBJECTAK1 = "ConnObj_AK1";
        public int Index { get; set; }

        int loadClass;
        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();

                    loadClass = (int)((PropertyValueInt)support.GetPropertyValue("IJOAFINL_LoadClass", "LoadClass")).PropValue;

                    parts.Add(new PartInfo(EYENUTAK1 + Index, "FINLCmp_SFS5389_" + loadClass));
                    parts.Add(new PartInfo(PIPECLAMPAK1 + Index, "FINLCmp_SFS5371"));
                    parts.Add(new PartInfo(CONNOBJECTAK1 + Index, "Log_Conn_Part_1"));

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
        //Get Assembly Catalog Parts
        //-----------------------------------------------------------------------------------
        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;

                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;

                string sAssmDesc = support.SupportDefinition.PartDescription;
                BoundingBoxHelper.CreateStandardBoundingBoxes(false);

                loadClass = (int)((PropertyValueInt)support.GetPropertyValue("IJOAFINL_LoadClass", "LoadClass")).PropValue;

                PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                double clampNominalDiameter = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Npd, pipeInfo.NominalDiameter.Size, MiddleServiceProvider.UOMMgr.GetUnit(UnitType.Npd, pipeInfo.NominalDiameter.Units), UnitName.NPD_MILLIMETER) / 1000;
                double A = 0, E = 0, toClamp, B = 0, C = 0, X = 0, toEyeNut;

                A = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5371", "IJUAFINL_CtoC", "CtoC", "IJUAFINL_PipeND_mm", "PipeND", (clampNominalDiameter - 0.00001), (clampNominalDiameter + 0.00001));
                A = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5371", "IJUAFINL_E", "E", "IJUAFINL_PipeND_mm", "PipeND", (clampNominalDiameter - 0.00001), (clampNominalDiameter + 0.00001));
                toClamp = A / 2 + E;

                B = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5389", "IJUAFINL_B", "B", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString());
                C = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5389", "IJUAFINL_C", "C", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString());
                X = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5389", "IJUAFINL_X", "X", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString());
                toEyeNut = B + C - X;

                if (SupportHelper.SupportedObjects.Count > 1)
                {
                    if (RefPortHelper.DistanceBetweenPorts("Route", "Route_2", PortDistanceType.Direct) < toClamp + toEyeNut + 0.1)
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Assembly may be too short.", "", "SFS5380_AK1.cs", 112);
                        return;
                    }
                }
                else
                {
                    if (RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Direct) < toClamp + toEyeNut + 0.1)
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Assembly may be too short.", "", "SFS5380_AK1.cs", 120);
                        return;
                    }
                }

                double dNDFrom = FINLAssemblyServices.GetDataByCondition("FINLSubAsm_SFS5380_AK1", "IJHgrSupportDefinition", "NDFrom", "IJOAFINL_LoadClass", "LoadClass", loadClass);
                double dNDTo = FINLAssemblyServices.GetDataByCondition("FINLSubAsm_SFS5380_AK1", "IJHgrSupportDefinition", "NDTo", "IJOAFINL_LoadClass", "LoadClass", loadClass);
                double diameter = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Npd, pipeInfo.NominalDiameter.Size, MiddleServiceProvider.UOMMgr.GetUnit(UnitType.Npd, pipeInfo.NominalDiameter.Units), UnitName.NPD_INCH);

                if (dNDFrom > diameter || diameter > dNDTo)
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Load Class of Supported Pipe does not match Load Class of supporting Pipe", "", "SFS5380_AK1.cs", 130);

                if (sAssmDesc.Contains("YK3"))
                {
                    if (SupportHelper.PlacementType == PlacementType.PlaceByPoint || SupportHelper.PlacementType == PlacementType.PlaceByReference)
                    {
                        double pipeDiameter = pipeInfo.OutsideDiameter - 2 * pipeInfo.InsulationThickness;
                        JointHelper.CreateRigidJoint("-1", "BBR_High", PIPECLAMPAK1 + Index, "Route", Plane.XY, Plane.XY, Axis.X, Axis.Y, -pipeDiameter / 2, -pipeDiameter / 2, 0);
                        JointHelper.CreateRevoluteJoint(PIPECLAMPAK1 + Index, "Pin", EYENUTAK1 + Index, "Eye", Axis.X, Axis.X);
                    }
                    else
                    {
                        JointHelper.CreateRigidJoint("-1", "Route", CONNOBJECTAK1 + Index, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, 0, 0);
                        JointHelper.CreatePrismaticJoint(CONNOBJECTAK1 + Index, "Connection", PIPECLAMPAK1 + Index, "Route", Plane.XY, Plane.XY, Axis.Y, Axis.NegativeY, 0, 0);
                        JointHelper.CreateRevoluteJoint(PIPECLAMPAK1 + Index, "Pin", EYENUTAK1 + Index, "Eye", Axis.X, Axis.X);
                    }
                }
                else
                {
                    if (Configuration == 1)
                        JointHelper.CreatePrismaticJoint("-1", "Route", PIPECLAMPAK1 + Index, "Route", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, 0);
                    else
                        JointHelper.CreatePrismaticJoint("-1", "Route", PIPECLAMPAK1 + Index, "Route", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0);

                    JointHelper.CreateRevoluteJoint(PIPECLAMPAK1 + Index, "Pin", EYENUTAK1 + Index, "Eye", Axis.X, Axis.X);
                }
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in Get Assembly Joints Method of FINL_ASSY." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }

        public override Collection<ConnectionInfo> SupportedConnections
        {
            get
            {
                try
                {
                    //Create a collection to hold the ALL Route connection information
                    Collection<ConnectionInfo> routeConnections = new Collection<ConnectionInfo>();
                    routeConnections.Add(new ConnectionInfo(PIPECLAMPAK1 + Index, 1)); // partindex, routeindex

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
                    //Create a collection to hold ALL the Structure Connection information
                    Collection<ConnectionInfo> structConnections = new Collection<ConnectionInfo>();

                    structConnections.Add(new ConnectionInfo(EYENUTAK1 + Index, 1)); // partindex, routeindex

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