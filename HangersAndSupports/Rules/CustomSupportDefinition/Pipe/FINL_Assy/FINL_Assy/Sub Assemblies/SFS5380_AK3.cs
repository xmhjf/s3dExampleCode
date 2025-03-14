﻿//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SFS5380_AK3.cs
//   FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5380_AK3
//   Author       : Rajeswari
//   Creation Date: 27-June-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
// 27-June-2013  Rajeswari CR-CP-224491- Convert HS_FINL_SubAssy to C# .Net 
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.Generic;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle.Services;

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
   public class SFS5380_AK3 : CustomSupportDefinition
    {
        private const string PIPECLAMP1 = "PIPECLAMP1_AK3";
        private const string PIPECLAMP2 = "PIPECLAMP2_AK3";
        private const string TRIANGLEPLATE = "TRIANGLEPLATE_AK3";
        private const string SCREWLUG = "SCREWLUG_AK3";
        private const string LOGOBJ = "LOGOBJ_AK3";
        private const string LOGOBJ2 = "LOGOBJ2_AK3";
        private const string LOGOBJ3 = "LOGOBJ3_AK3";
        private const string LOGOBJ4 = "LOGOBJ4_AK3";
        public int Index { get; set; }

        int loadClass;
        double nominalDiameter;
        PipeObjectInfo pipeInfo;
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

                    pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                    nominalDiameter = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Npd, pipeInfo.NominalDiameter.Size, MiddleServiceProvider.UOMMgr.GetUnit(UnitType.Npd, pipeInfo.NominalDiameter.Units), UnitName.NPD_INCH); 

                    PropertyValue clampsize;
                    Dictionary<String, Object> parameter = new Dictionary<String, Object>();
                    parameter.Add("ClampSizeIn", nominalDiameter.ToString());
                    GenericHelper.GetDataByRule("FINLSubSrv_SFS5371_Clamp", "IJUAFINLSrv_5371_Clamp", "ClampPart", parameter, out clampsize);

                    loadClass = (int)((PropertyValueInt)support.GetPropertyValue("IJOAFINL_LoadClass", "LoadClass")).PropValue;

                    parts.Add(new PartInfo(PIPECLAMP1 + Index, "FINLCmp_SFS5370_" + clampsize));
                    parts.Add(new PartInfo(PIPECLAMP2 + Index, "FINLCmp_SFS5370_" + clampsize));
                    parts.Add(new PartInfo(TRIANGLEPLATE + Index, "FINLCmp_SFS5384"));
                    parts.Add(new PartInfo(SCREWLUG + Index, "FINLCmp_SFS5390_" + (loadClass.ToString()).Trim()));
                    parts.Add(new PartInfo(LOGOBJ + Index, "Log_Conn_Part_1"));
                    parts.Add(new PartInfo(LOGOBJ2 + Index, "Log_Conn_Part_1"));
                    parts.Add(new PartInfo(LOGOBJ3 + Index, "Log_Conn_Part_1"));
                    parts.Add(new PartInfo(LOGOBJ4 + Index, "Log_Conn_Part_1"));

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

                string assmDescription = support.SupportDefinition.PartDescription;
                double holeToHole = 0, C = 0, A = 0, X = 0, toClamp = 0, toPlate = 0, toClevis = 0, ndFrom = 0, ndTo = 0;

                double nominalPipeDiameter = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Npd, pipeInfo.NominalDiameter.Size, MiddleServiceProvider.UOMMgr.GetUnit(UnitType.Npd, pipeInfo.NominalDiameter.Units), UnitName.NPD_MILLIMETER) / 1000;

                C = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5384", "IJUAFINL_C", "C", "IJUAFINL_PipeND_mm", "PipeND", (nominalPipeDiameter - 0.001), (nominalPipeDiameter + 0.001));
                holeToHole = Math.Sqrt(2) * C;

                if (nominalDiameter < 2)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Cannot place AK3 for Pipesize < 2in. Becasue SFS5384 does not available for pipesize < 2in.", "", "SFS5380_AK3.cs", 112);
                    return;
                }

                // load standard bounding box definition
                BoundingBoxHelper.CreateStandardBoundingBoxes(false);

                A = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5370", "IJUAFINL_CtoC", "CtoC", "IJUAFINL_PipeND_mm", "PipeND", (nominalPipeDiameter - 0.001), (nominalPipeDiameter + 0.001));
                toClamp = A / 2;
                toPlate = C / Math.Sqrt(2);

                A = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5390", "IJUAFINL_A", "A", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString());
                X = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5390", "IJUAFINL_X", "X", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString());
                toClevis = A - X;

                if (SupportHelper.SupportedObjects.Count > 1)
                {
                    if (RefPortHelper.DistanceBetweenPorts("Route", "Route_2", PortDistanceType.Direct) < toClamp + toClevis + toPlate + 0.1)
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Assembly may be too short.", "", "SFS5380_AK3.cs", 131);
                        return;
                    }
                }
                else
                {
                    if (RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Direct) < toClamp + toClevis + toPlate + 0.1)
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Assembly may be too short.", "", "SFS5380_AK3.cs", 139);
                        return;
                    }
                }

                // Get NDFrom and NDTo, if invalid pipesize based on PipeDia, pop error
                ndFrom = FINLAssemblyServices.GetDataByCondition("FINLSubAsm_SFS5380_AK2", "IJHgrSupportDefinition", "NDFrom", "IJOAFINL_LoadClass", "LoadClass", loadClass);
                ndTo = FINLAssemblyServices.GetDataByCondition("FINLSubAsm_SFS5380_AK2", "IJHgrSupportDefinition", "NDTo", "IJOAFINL_LoadClass", "LoadClass", loadClass);
                if (ndFrom > nominalDiameter || nominalDiameter > ndTo)
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Load Class of Supported Pipe does not match Load Class of supporting Pipe.", "", "SFS5380_AK3.cs", 148);

                if (assmDescription.IndexOf("YK3") > 0)
                {
                    if (SupportHelper.PlacementType == PlacementType.PlaceByPoint || SupportHelper.PlacementType == PlacementType.PlaceByReference)
                    {
                        double pipeDiameter = pipeInfo.OutsideDiameter;
                        JointHelper.CreateRigidJoint("-1", "BBR_High", PIPECLAMP1 + Index, "Route", Plane.XY, Plane.XY, Axis.X, Axis.Y, -pipeDiameter / 2, -pipeDiameter / 2, 0);
                    }
                    else
                    {
                        JointHelper.CreateRigidJoint("-1", "Route", LOGOBJ2 + Index, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, 0, 0);

                        JointHelper.CreatePrismaticJoint(LOGOBJ2 + Index, "Connection", PIPECLAMP1 + Index, "Route", Plane.XY, Plane.XY, Axis.Y, Axis.NegativeY, 0, 0);
                    }
                }
                else
                    JointHelper.CreatePrismaticJoint(PIPECLAMP1 + Index, "Route", "-1", "Route", Plane.XY, Plane.XY, Axis.Y, Axis.X, 0, 0);

                JointHelper.CreateRigidJoint(PIPECLAMP1 + Index, "Pin", TRIANGLEPLATE + Index, "Hole33", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, 0);

                JointHelper.CreateRigidJoint(TRIANGLEPLATE + Index, "Hole22", PIPECLAMP2 + Index, "Pin", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, 0);

                JointHelper.CreateRigidJoint(TRIANGLEPLATE + Index, "Hole1", LOGOBJ + Index, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, 0);

                JointHelper.CreateRevoluteJoint(SCREWLUG + Index, "Pin", LOGOBJ + Index, "Connection", Axis.X, Axis.NegativeX);
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
                    //Create a collection to hold the ALL Route connection information
                    Collection<ConnectionInfo> routeConnections = new Collection<ConnectionInfo>();
                    routeConnections.Add(new ConnectionInfo(PIPECLAMP1 + Index, 1)); // partindex, routeindex
                    routeConnections.Add(new ConnectionInfo(PIPECLAMP2 + Index, 1)); // partindex, routeindex

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
                    structConnections.Add(new ConnectionInfo(SCREWLUG + Index, 1)); // partindex, routeindex

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
