//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SFS5380_AR6.cs
//   FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5380_AR6
//   Author       : Rajeswari
//   Creation Date: 17-April-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
// 05-July-2013  Rajeswari CR-CP-224491- Convert HS_FINL_SubAssy to C# .Net
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;

namespace Ingr.SP3D.Content.Support.Rules
{
    //-----------------------------------------------------------------------------------
    //Namespace of this class is Ingr.SP3D.Content.Support.Rules
    //It is recommended that customers specify namespace of their Rules to be
    //<CompanyName>.SP3D.Content.<Specialization>.
    //It is also recommended that if customers want to change this Rule to suit their
    //requirements, they should change namespace/Rule name so the identity of the modified
    //Rule will be different from the one delivered by Intergraph.
    //--------------------------------------------------------------------------------

    public class SFS5380_AR6 : CustomSupportDefinition
    {
        private const string TURNBUCKLE = "TURNBUCKLE_AR6";
        private const string EYENUT = "EYENUT_AR6";
        private const string PIPECLAMP = "PIPECLAMP_AR6";
        private const string LOGOBJ = "LOGOBJ_AR6";
        public int Index { get; set; }

        int loadClass;
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
                  
                    // GET variable values from XLS files
                    loadClass = (int)((PropertyValueInt)support.GetPropertyValue("IJOAFINL_LoadClass", "LoadClass")).PropValue;

                    parts.Add(new PartInfo(TURNBUCKLE + Index, "FINLCmp_SFS5391_" + loadClass));
                    parts.Add(new PartInfo(EYENUT + Index, "FINLCmp_SFS5389_" + loadClass));
                    parts.Add(new PartInfo(PIPECLAMP + Index, "FINLCmp_SFS5857"));
                    parts.Add(new PartInfo(LOGOBJ + Index, "Log_Conn_Part_1"));

                    // Return the collection of Catalog Parts
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
        //Get Assembly Joints
        //-----------------------------------------------------------------------------------
        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;

                // Get interface for accessing items on the collection of Part Occurences
                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
                Part supportPart = support.SupportDefinition;
                string assmDescription = supportPart.PartDescription;

                // load standard bounding box definition
                BoundingBoxHelper.CreateStandardBoundingBoxes(false);

                // Check whether the dis between pipe and structure is too short.
                double A = 0, B = 0, C = 0, E = 0, L = 0, X = 0, toClamp = 0, toTurn = 0, toEyenut = 0, ndFrom = 0, ndTo = 0;
                PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                double nominalPipeDiameter = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Npd, pipeInfo.NominalDiameter.Size, MiddleServiceProvider.UOMMgr.GetUnit(UnitType.Npd, pipeInfo.NominalDiameter.Units), UnitName.NPD_MILLIMETER) / 1000;

                A = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5857", "IJUAFINL_A", "A", "IJUAFINL_PipeND_mm", "PipeND", (nominalPipeDiameter - 0.00001), (nominalPipeDiameter + 0.00001));
                E = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5371", "IJUAFINL_E", "E", "IJUAFINL_PipeND_mm", "PipeND", (nominalPipeDiameter - 0.00001), (nominalPipeDiameter + 0.00001));
                toClamp = A / 2 + E;
                A = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5391", "IJUAFINL_A", "A", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString());
                X = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5391", "IJUAFINL_X", "X", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString());
                L = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5391", "IJUAFINL_L", "L", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString());
                toTurn = A + L - X - X;
                B = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5389", "IJUAFINL_B", "B", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString());
                C = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5389", "IJUAFINL_C", "C", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString());
                X = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5389", "IJUAFINL_X", "X", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString());
                toEyenut = B + C - X;

                // Get NDFrom and NDTo, if invalid pipesize based on PipeDia, pop error
                double nominalDiameter = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Npd, pipeInfo.NominalDiameter.Size, MiddleServiceProvider.UOMMgr.GetUnit(UnitType.Npd, pipeInfo.NominalDiameter.Units), UnitName.NPD_INCH); 
                if (support.SupportedObjects.Count > 1)
                {
                    if (RefPortHelper.DistanceBetweenPorts("Route", "Route_2", PortDistanceType.Direct) < toClamp + toTurn + toEyenut + 0.1)
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Assembly may be too short.", "", "SFS5380_AR6.cs", 115);
                        return;
                    }
                }
                else
                {
                    if (RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Direct) < toClamp + toTurn + toEyenut + 0.1)
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Assembly may be too short.", "", "SFS5380_AR6.cs", 123);
                        return;
                    }
                }

                ndFrom = FINLAssemblyServices.GetDataByCondition("FINLSubAsm_SFS5380_AR6", "IJHgrSupportDefinition", "NDFrom", "IJOAFINL_LoadClass", "LoadClass", loadClass);
                ndTo = FINLAssemblyServices.GetDataByCondition("FINLSubAsm_SFS5380_AR6", "IJHgrSupportDefinition", "NDTo", "IJOAFINL_LoadClass", "LoadClass", loadClass);

                if (ndFrom > nominalDiameter || nominalDiameter > ndTo)
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Load Class of Supported Pipe does not match Load Class of supporting Pipe.", "", "SFS5380_AR6.cs", 132);

                double pipeDiameter = pipeInfo.OutsideDiameter;
                if (assmDescription.IndexOf("YK3") > 0)
                {
                    if (SupportHelper.PlacementType == PlacementType.PlaceByPoint || SupportHelper.PlacementType == PlacementType.PlaceByReference)
                    {
                        JointHelper.CreateRigidJoint("-1", "BBR_High", PIPECLAMP + Index, "Route", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, -pipeDiameter / 2, -pipeDiameter / 2, 0);

                        JointHelper.CreateRigidJoint(EYENUT + Index, "InThdRH", TURNBUCKLE + Index, "BotExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    }
                    else
                    {
                        JointHelper.CreatePrismaticJoint("-1", "Route", PIPECLAMP + Index, "Route", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, 0);

                        JointHelper.CreateRigidJoint(EYENUT + Index, "InThdRH", TURNBUCKLE + Index, "BotExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    }
                    JointHelper.CreateRevoluteJoint(PIPECLAMP + Index, "Pin", EYENUT + Index, "Eye", Axis.X, Axis.X);
                }
                else
                {
                    if (Configuration == 1)
                        JointHelper.CreatePrismaticJoint("-1", "Route", PIPECLAMP + Index, "Route", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, 0);
                    else
                        JointHelper.CreatePrismaticJoint("-1", "Route", PIPECLAMP + Index, "Route", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0);

                    JointHelper.CreateRevoluteJoint(PIPECLAMP + Index, "Pin", EYENUT + Index, "Eye", Axis.X, Axis.X);

                    JointHelper.CreateRevoluteJoint(EYENUT + Index, "InThdRH", TURNBUCKLE + Index, "BotExThdRH", Axis.Z, Axis.Z);

                    JointHelper.CreateGlobalAxesAlignedJoint(TURNBUCKLE + Index, "BotExThdRH", Axis.Z, Axis.Z);
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
                    routeConnections.Add(new ConnectionInfo(PIPECLAMP + Index, 1)); // partindex, routeindex

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
                    structConnections.Add(new ConnectionInfo(TURNBUCKLE + Index, 1)); // partindex, routeindex

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
