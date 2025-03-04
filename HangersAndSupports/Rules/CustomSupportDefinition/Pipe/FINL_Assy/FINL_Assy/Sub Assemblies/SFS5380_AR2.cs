//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SFS5380_AK2.cs
//   FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5380_AR2
//   Author       : Rajeswari
//   Creation Date: 17-April-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
// 05-July-2013  Rajeswari CR-CP-224491- Convert HS_FINL_SubAssy to C# .Net
//   17/12/2015     Ramya   TR 284319	Multiple Record exception dumps are created on copy pasting supports
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
    public class SFS5380_AR2 : CustomSupportDefinition
    {
        private const string PIPECLAMP1 = "PIPECLAMP1_AR2";
        private const string PIPECLAMP2 = "PIPECLAMP2_AR2";
        private const string TRIANGLEPLATE1 = "TRIANGLEPLATE1_AR2";
        private const string TRIANGLEPLATE2 = "TRIANGLEPLATE2_AR2";
        private const string SCREWLUG1 = "SCREWLUG1_AR2";
        private const string SCREWLUG2 = "SCREWLUG2_AR2";
        private const string STOPPER1 = "STOPPER1_AR2";
        private const string STOPPER2 = "STOPPER2_AR2";
        private const string LOGOBJ = "LOGOBJ_AR2";
        public int Index { get; set; }

        int loadClass;
        double nominalPipeDiameter;
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

                    // Get Pipe Dia to get the partnumber
                    pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                    nominalPipeDiameter = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Npd, pipeInfo.NominalDiameter.Size, MiddleServiceProvider.UOMMgr.GetUnit(UnitType.Npd, pipeInfo.NominalDiameter.Units), UnitName.NPD_MILLIMETER) / 1000;

                    // check Nom Pipe Dia
                    if (nominalPipeDiameter < 50.0 / 1000 && nominalPipeDiameter > 500.0 / 1000)
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "Parts" + ": " + "ERROR: " + "Pipe size not available.", "", "SFS5380_AR2.cs", 66);
                        return null;
                    }

                    loadClass = (int)((PropertyValueInt)support.GetPropertyValue("IJOAFINL_LoadClass", "LoadClass")).PropValue;

                    parts.Add(new PartInfo(PIPECLAMP1 + Index, "FINLCmp_SFS5856"));
                    parts.Add(new PartInfo(PIPECLAMP2 + Index, "FINLCmp_SFS5856"));
                    parts.Add(new PartInfo(TRIANGLEPLATE1 + Index, "FINLCmp_SFS5862"));
                    parts.Add(new PartInfo(TRIANGLEPLATE2 + Index, "FINLCmp_SFS5862"));
                    parts.Add(new PartInfo(SCREWLUG1 + Index, "FINLCmp_SFS5390_" + (loadClass.ToString()).Trim()));
                    parts.Add(new PartInfo(SCREWLUG2 + Index, "FINLCmp_SFS5390_" + (loadClass.ToString()).Trim()));
                    parts.Add(new PartInfo(STOPPER1 + Index, "FINLCmp_SFS5368"));
                    parts.Add(new PartInfo(STOPPER2 + Index, "FINLCmp_SFS5368"));
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
                string assmDescription = support.SupportDefinition.PartDescription;

                double holeToHole = 0, clampWidth = 0, stopperL = 0, A = 0, X = 0, toClevis, ndFrom = 0, ndTo = 0;

                holeToHole = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5862", "IJUAFINL_C", "C", "IJUAFINL_PipeND_mm", "PipeND", (nominalPipeDiameter - 0.001), (nominalPipeDiameter + 0.001));
                clampWidth = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5856", "IJUAFINL_B", "B", "IJUAFINL_PipeND_mm", "PipeND", (nominalPipeDiameter - 0.001), (nominalPipeDiameter + 0.001));
                stopperL = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5368", "IJUAFINL_L", "L", "IJUAFINL_PipeND_mm", "PipeND", (nominalPipeDiameter - 0.001), (nominalPipeDiameter + 0.001));

                PropertyValueCodelist typecodelist = (PropertyValueCodelist)componentDictionary[STOPPER1 + Index].GetPropertyValue("IJOAFINL_StopperType", "StopperType");
                int type = (int)((PropertyValueCodelist)support.GetPropertyValue("IJUAFINLStopperType", "Type")).PropValue;
                if (type == 1)
                {
                    typecodelist.PropValue = 1;
                    componentDictionary[STOPPER1 + Index].SetPropertyValue(typecodelist.PropValue, "IJOAFINL_StopperType", "StopperType");
                    componentDictionary[STOPPER2 + Index].SetPropertyValue(typecodelist.PropValue, "IJOAFINL_StopperType", "StopperType");
                }
                else
                {
                    typecodelist.PropValue = 2;
                    componentDictionary[STOPPER1 + Index].SetPropertyValue(typecodelist.PropValue, "IJOAFINL_StopperType", "StopperType");
                    componentDictionary[STOPPER2 + Index].SetPropertyValue(typecodelist.PropValue, "IJOAFINL_StopperType", "StopperType");
                }

                // GET variable values from XLS files
                A = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5390", "IJUAFINL_A", "A", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString());
                X = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5390", "IJUAFINL_X", "X", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString());
                toClevis = A - X;

                if (support.SupportedObjects.Count > 1)
                {
                    if (RefPortHelper.DistanceBetweenPorts("Route", "Route_2", PortDistanceType.Direct) < toClevis + 0.1)
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Assembly may be too short.", "", "SFS5380_AR2.cs", 136);
                        return;
                    }
                }
                else
                {
                    if (RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Direct) < toClevis + 0.1)
                    {
                        MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Assembly may be too short.", "", "SFS5380_AR2.cs", 145);
                        return;
                    }
                }

                // Get NDFrom and NDTo, if invalid pipesize based on PipeDia, pop error
                ndFrom = FINLAssemblyServices.GetDataByCondition("FINLSubAsm_SFS5380_AR2", "IJHgrSupportDefinition", "NDFrom", "IJOAFINL_LoadClass", "LoadClass", loadClass);
                ndTo = FINLAssemblyServices.GetDataByCondition("FINLSubAsm_SFS5380_AR2", "IJHgrSupportDefinition", "NDTo", "IJOAFINL_LoadClass", "LoadClass", loadClass);
                double nominalDiameter = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Npd, pipeInfo.NominalDiameter.Size, MiddleServiceProvider.UOMMgr.GetUnit(UnitType.Npd, pipeInfo.NominalDiameter.Units), UnitName.NPD_INCH);

                if (ndFrom > nominalDiameter || nominalDiameter > ndTo)
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Load Class of Supported Pipe does not match Load Class of supporting Pipe.", "", "SFS5380_AR2.cs", 155);

                double routeXStructAngle = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, "Structure", PortAxisType.X, OrientationAlong.Direct);
                double routeXStructAngle2 = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, "Structure", PortAxisType.Y, OrientationAlong.Global_Z);

                string supportingType = String.Empty;

                if ((SupportHelper.SupportingObjects.Count != 0))
                {
                    if ((SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.Member) || (SupportingHelper.SupportingObjectInfo(1).SupportingObjectType == SupportingObjectType.HangerBeam))
                        supportingType = "Steel";    //Steel                      
                    else
                        supportingType = "Slab"; //Slab
                }
                else
                    supportingType = "Slab"; //For PlaceByReference

                if (support.SupportedObjects.Count > 1)
                {
                    JointHelper.CreateRigidJoint("-1", "Route", LOGOBJ + Index, "Connection", Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeX, 0, 0, 0);

                    if ((RefPortHelper.AngleBetweenPorts("Route", PortAxisType.X, "world", PortAxisType.Z, OrientationAlong.Global_Z)) * 180 / Math.PI < 1)
                        JointHelper.CreateRevoluteJoint(LOGOBJ + Index, "Connection", PIPECLAMP1 + Index, "Route", Axis.X, Axis.NegativeY);
                    else
                        JointHelper.CreateRevoluteJoint(LOGOBJ + Index, "Connection", PIPECLAMP1 + Index, "Route", Axis.X, Axis.Y);
                }
                else
                {
                    if (supportingType == "Slab")
                    {
                        if (Configuration == 1 || Configuration == 3)
                            JointHelper.CreateRigidJoint("-1", "Route", PIPECLAMP1 + Index, "Route", Plane.XY, Plane.ZX, Axis.X, Axis.Z, 0, 0, 0);
                        else
                            JointHelper.CreateRigidJoint("-1", "Route", PIPECLAMP1 + Index, "Route", Plane.XY, Plane.ZX, Axis.X, Axis.NegativeX, 0, 0, 0);

                        JointHelper.CreateRevoluteJoint(LOGOBJ + Index, "Connection", PIPECLAMP1 + Index, "Route", Axis.X, Axis.Y);
                    }
                    else
                    {
                        if (Math.Abs((routeXStructAngle) * 180 / Math.PI - 90) < 1)
                        {
                            if (Math.Abs((routeXStructAngle2) * 180 / Math.PI - 180) < 1)
                                JointHelper.CreateRigidJoint("-1", "Route", LOGOBJ + Index, "Connection", Plane.XY, Plane.YZ, Axis.X, Axis.NegativeY, 0, 0, 0);
                            else if (Math.Abs((routeXStructAngle2) * 180 / Math.PI) < 1)
                                JointHelper.CreateRigidJoint("-1", "Route", LOGOBJ + Index, "Connection", Plane.XY, Plane.ZX, Axis.X, Axis.X, 0, 0, 0);
                            else
                                JointHelper.CreateRigidJoint("-1", "Route", LOGOBJ + Index, "Connection", Plane.XY, Plane.ZX, Axis.X, Axis.X, 0, 0, 0);
                        }
                        else
                            JointHelper.CreateRigidJoint("-1", "Route", LOGOBJ + Index, "Connection", Plane.XY, Plane.ZX, Axis.X, Axis.Z, 0, 0, 0);

                        JointHelper.CreateRigidJoint(LOGOBJ + Index, "Connection", PIPECLAMP1 + Index, "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                    }
                }
                JointHelper.CreateRigidJoint(PIPECLAMP1 + Index, "Route", PIPECLAMP2 + Index, "Route", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, holeToHole, 0);

                JointHelper.CreateRigidJoint(PIPECLAMP2 + Index, "Pin1", TRIANGLEPLATE1 + Index, "Hole1", Plane.ZX, Plane.YZ, Axis.Z, Axis.NegativeZ, 0, 0, 0);

                JointHelper.CreateRigidJoint(SCREWLUG1 + Index, "Pin", TRIANGLEPLATE1 + Index, "Hole2", Plane.XY, Plane.NegativeYZ, Axis.X, Axis.Y, 0, 0, 0);

                JointHelper.CreateRigidJoint(PIPECLAMP2 + Index, "Pin2", TRIANGLEPLATE2 + Index, "Hole1", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, 0, 0);

                JointHelper.CreateRigidJoint(SCREWLUG2 + Index, "Pin", TRIANGLEPLATE2 + Index, "Hole2", Plane.XY, Plane.NegativeYZ, Axis.X, Axis.Y, 0, 0, 0);

                JointHelper.CreateRigidJoint(PIPECLAMP2 + Index, "Route", STOPPER1 + Index, "Route", Plane.XY, Plane.NegativeZX, Axis.Y, Axis.NegativeX, 0, 0, clampWidth / 2 + stopperL / 2);

                JointHelper.CreateRigidJoint(PIPECLAMP1 + Index, "Route", STOPPER2 + Index, "Route", Plane.XY, Plane.NegativeZX, Axis.Y, Axis.NegativeX, 0, 0, clampWidth / 2 + stopperL / 2);

                if (support.SupportedObjects.Count == 2)
                    JointHelper.CreatePlanarJoint(SCREWLUG2 + Index, "InThdRH", "-1", "Route_2", Plane.YZ, Plane.ZX, 0);
                else if (support.SupportedObjects.Count == 3)
                {
                }
                else
                {
                    if (supportingType == "Steel")
                    {
                        if (support.SupportingObjects.Count > 1)
                            JointHelper.CreatePlanarJoint(SCREWLUG2 + Index, "InThdRH", "-1", "Route", Plane.YZ, Plane.ZX, 0);
                    }
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
                return 4;
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
                    structConnections.Add(new ConnectionInfo(SCREWLUG1 + Index, 1)); // partindex, routeindex
                    structConnections.Add(new ConnectionInfo(SCREWLUG2 + Index, 1)); // partindex, routeindex

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
