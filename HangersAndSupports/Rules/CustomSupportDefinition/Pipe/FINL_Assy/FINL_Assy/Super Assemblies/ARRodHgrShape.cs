﻿//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   ARRodHgrShape.cs
//   FINL_Assy,Ingr.SP3D.Content.Support.Rules.ARRodHgrShape
//   Author       :  Hema
//   Creation Date:  28/06/2013
//   Description:    

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  28/06/2013       Hema    CR-CP-224492 Convert HS_FINL_SupAssy to C# .Net
//   17/12/2015     Ramya   TR 284319	Multiple Record exception dumps are created on copy pasting supports
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Support.Middle;
using System.Linq;
namespace Ingr.SP3D.Content.Support.Rules
{
    public class ARRodHgrShape : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        CustomSupportDefinition subTopAssembly,subMiddleAssembly,subBottomAssembly; //default value is IsEmpty

        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    
                    string subTop = (string)((PropertyValueString)support.GetPropertyValue("IJUAFINLSrv_RodHgrSteel", "Sub_Top")).PropValue;
                    string subBottom = (string)((PropertyValueString)support.GetPropertyValue("IJUAFINLSrv_RodHgrSteel", "Sub_Bottom")).PropValue;
                    string subMiddle = (string)((PropertyValueString)support.GetPropertyValue("IJUAFINLSrv_RodHgrSteel", "Sub_Middle")).PropValue;
                    if (!string.IsNullOrEmpty(subTop))
                        subTopAssembly = FINLAssemblyServices.GetAssembly(subTop, support);
                    if (!string.IsNullOrEmpty(subBottom))
                        subBottomAssembly = FINLAssemblyServices.GetAssembly(subBottom, support);
                    if (!string.IsNullOrEmpty(subMiddle))
                        subMiddleAssembly = FINLAssemblyServices.GetAssembly(subMiddle, support);

                    if (!string.IsNullOrEmpty(subMiddle))
                        return new Collection<PartInfo>(subTopAssembly.Parts.Concat(subMiddleAssembly.Parts).Concat(subBottomAssembly.Parts).ToList());
                    else
                        return new Collection<PartInfo>(subTopAssembly.Parts.Concat(subBottomAssembly.Parts).ToList());    //Return the collection of Catalog Parts

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
                    //Get the part index in Super-Assy, which part in Bottom sub-assembly will connect to the Route
                    Collection<ConnectionInfo> routeConnections = new Collection<ConnectionInfo>();
                    routeConnections = subBottomAssembly.SupportedConnections;
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
                    //Get the part index in Super-Assy, which part in top sub-assembly will connect to the structure
                    Collection<ConnectionInfo> structConnections = new Collection<ConnectionInfo>();
                    structConnections = subTopAssembly.SupportingConnections;
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
                ReadOnlyCollection<object> topJoints = FINLAssemblyServices.GetSubAssemblyJoints(this, subTopAssembly, oSupCompColl);
                ReadOnlyCollection<object> bottomJoints = FINLAssemblyServices.GetSubAssemblyJoints(this, subBottomAssembly, oSupCompColl);

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

                if (subMiddleAssembly != null)
                {
                    ReadOnlyCollection<object> middleJoints = FINLAssemblyServices.GetSubAssemblyJoints(this, subMiddleAssembly, oSupCompColl);
                    JointHelper.m_oCollOfJoints = new Collection<object>(topJoints.Concat(middleJoints).Concat(bottomJoints).ToList());
                }
                else
                    JointHelper.m_oCollOfJoints = new Collection<object>(topJoints.Concat(bottomJoints).ToList());

                Collection<ConnectionInfo> subAssyRouteConnections = subTopAssembly.SupportedConnections;
                Collection<ConnectionInfo> subAssyStructConnections = subBottomAssembly.SupportingConnections;

                if (subBottomAssembly.GetType().FullName.Equals("Ingr.SP3D.Content.Support.Rules.SFS5380_AR6"))
                {
                    if (subTopAssembly.GetType().FullName.Equals("Ingr.SP3D.Content.Support.Rules.SFS5380_YK3") && (SupportHelper.PlacementType == PlacementType.PlaceByPoint|| SupportHelper.PlacementType == PlacementType.PlaceByReference))
                    {
                        if (supportingType == "Slab")
                        {
                            if (Configuration == 1)
                                JointHelper.CreateRigidJoint(subAssyRouteConnections[0].PartKey, "TopExThdRH", subAssyStructConnections[0].PartKey, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                            else
                                JointHelper.CreateRigidJoint(subAssyRouteConnections[0].PartKey, "TopExThdRH", subAssyStructConnections[0].PartKey, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.NegativeX, 0, 0, 0);
                        }
                        else
                        {
                            if (Configuration == 1)
                                JointHelper.CreateRigidJoint(subAssyRouteConnections[0].PartKey, "TopExThdRH", subAssyStructConnections[0].PartKey, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, 0, 0);
                            else
                                JointHelper.CreateRigidJoint(subAssyRouteConnections[0].PartKey, "TopExThdRH", subAssyStructConnections[0].PartKey, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, 0);
                        }
                    }
                    else
                        JointHelper.CreateRevoluteJoint(subAssyRouteConnections[0].PartKey, "TopExThdRH", subAssyStructConnections[0].PartKey, "InThdRH", Axis.Z, Axis.Z);
                }
                else if (subBottomAssembly.GetType().FullName.Equals("Ingr.SP3D.Content.Support.Rules.SFS5380_AR3"))
                {
                    if (subTopAssembly.GetType().FullName.Equals("Ingr.SP3D.Content.Support.Rules.SFS5380_YK3"))
                    {
                        if (SupportHelper.PlacementType == PlacementType.PlaceByPoint || SupportHelper.PlacementType == PlacementType.PlaceByReference)
                        {
                            if (supportingType == "Slab")
                                if (Configuration == 1)
                                    JointHelper.CreateRigidJoint(subAssyRouteConnections[0].PartKey, "TopExThdRH", subAssyStructConnections[0].PartKey, "InThdRH", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, 0, 0);
                                else
                                    JointHelper.CreateRigidJoint(subAssyRouteConnections[0].PartKey, "TopExThdRH", subAssyStructConnections[0].PartKey, "InThdRH", Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeX, 0, 0, 0);
                            else
                                if (Configuration == 1)
                                    JointHelper.CreateRigidJoint(subAssyRouteConnections[0].PartKey, "TopExThdRH", subAssyStructConnections[0].PartKey, "InThdRH", Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeY, 0, 0, 0);
                                else
                                    JointHelper.CreateRigidJoint(subAssyRouteConnections[0].PartKey, "TopExThdRH", subAssyStructConnections[0].PartKey, "InThdRH", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, 0, 0, 0);
                        }
                        else
                            JointHelper.CreateRevoluteJoint(subAssyRouteConnections[0].PartKey, "TopExThdRH", subAssyStructConnections[0].PartKey, "InThdRH", Axis.Z, Axis.NegativeZ);
                    }
                    else
                    {
                        JointHelper.CreateRevoluteJoint(subAssyRouteConnections[0].PartKey, "TopExThdRH", subAssyStructConnections[0].PartKey, "InThdRH", Axis.Z, Axis.NegativeZ);
                    }
                }
                else if (subBottomAssembly.GetType().FullName.Equals("Ingr.SP3D.Content.Support.Rules.SFS5380_AR5"))
                {
                    if (subTopAssembly.GetType().FullName.Equals("Ingr.SP3D.Content.Support.Rules.SFS5380_YK3"))
                    {
                        if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                            JointHelper.CreateRevoluteJoint(subAssyRouteConnections[0].PartKey, "TopExThdRH", subAssyStructConnections[0].PartKey, "InThdRH", Axis.Z, Axis.NegativeZ);
                        else
                        {
                            if (supportingType == "Slab")
                            {
                                if (Configuration == 1)
                                    JointHelper.CreateRigidJoint(subAssyRouteConnections[0].PartKey, "TopExThdRH", subAssyStructConnections[0].PartKey, "InThdRH", Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeY, 0, 0, 0);
                                else
                                    JointHelper.CreateRigidJoint(subAssyRouteConnections[0].PartKey, "TopExThdRH", subAssyStructConnections[0].PartKey, "InThdRH", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, 0, 0, 0);
                            }
                            else
                            {
                                if (Configuration == 1)
                                    JointHelper.CreateRigidJoint(subAssyRouteConnections[0].PartKey, "TopExThdRH", subAssyStructConnections[0].PartKey, "InThdRH", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, 0, 0, 0);
                                else
                                    JointHelper.CreateRigidJoint(subAssyRouteConnections[0].PartKey, "TopExThdRH", subAssyStructConnections[0].PartKey, "InThdRH", Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeX, 0, 0, 0);
                            }
                        }
                    }
                    else
                        JointHelper.CreateRevoluteJoint(subAssyRouteConnections[0].PartKey, "TopExThdRH", subAssyStructConnections[0].PartKey, "InThdRH", Axis.Z, Axis.NegativeZ);
                }
                else if (subBottomAssembly.GetType().FullName.Equals("Ingr.SP3D.Content.Support.Rules.SFS5380_AR1"))
                {
                    if (subTopAssembly.GetType().FullName.Equals("Ingr.SP3D.Content.Support.Rules.SFS5380_YK3") && (supportingType == "Slab"))
                    {
                        if (Configuration == 1)
                            JointHelper.CreateRigidJoint(subAssyRouteConnections[0].PartKey, "TopExThdRH", subAssyStructConnections[0].PartKey, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        else
                            JointHelper.CreateRigidJoint(subAssyRouteConnections[0].PartKey, "TopExThdRH", subAssyStructConnections[0].PartKey, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.NegativeX, 0, 0, 0);
                    }
                    else if (subTopAssembly.GetType().FullName.Equals("Ingr.SP3D.Content.Support.Rules.SFS5380_YK3") && (SupportHelper.PlacementType == PlacementType.PlaceByPoint || SupportHelper.PlacementType == PlacementType.PlaceByReference))
                    {
                        if (Configuration == 1)
                            JointHelper.CreateRigidJoint(subAssyRouteConnections[0].PartKey, "TopExThdRH", subAssyStructConnections[0].PartKey, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, 0, 0, 0);
                        else
                            JointHelper.CreateRigidJoint(subAssyRouteConnections[0].PartKey, "TopExThdRH", subAssyStructConnections[0].PartKey, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.Y, 0, 0, 0);
                    }
                    else
                        JointHelper.CreateRevoluteJoint(subAssyRouteConnections[0].PartKey, "TopExThdRH", subAssyStructConnections[0].PartKey, "InThdRH", Axis.Z, Axis.Z);
                }
                else
                    JointHelper.CreateRevoluteJoint(subAssyRouteConnections[0].PartKey, "TopExThdRH", subAssyStructConnections[0].PartKey, "InThdRH", Axis.Z, Axis.Z);

                int loadClass = (int)((PropertyValueInt)support.GetPropertyValue("IJOAFINL_LoadClass", "LoadClass")).PropValue;
                PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                double maxTemperature = pipeInfo.MaxDesignTemperature, maxLoad = 0, bottomPipeNominalDiameter;
                string subBottomAssemblyName = subBottomAssembly.GetType().FullName;
                if (subBottomAssemblyName.Equals("Ingr.SP3D.Content.Support.Rules.SFS5380_AR1") || subBottomAssemblyName.Equals("Ingr.SP3D.Content.Support.Rules.SFS5380_AR5") || subBottomAssemblyName.Equals("Ingr.SP3D.Content.Support.Rules.SFS5380_AR6"))
                {
                    bottomPipeNominalDiameter = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Npd, pipeInfo.NominalDiameter.Size, MiddleServiceProvider.UOMMgr.GetUnit(UnitType.Npd, pipeInfo.NominalDiameter.Units), UnitName.NPD_INCH);
                   
                    if (maxTemperature <= 20)
                        maxLoad = FINLAssemblyServices.GetDataByCondition("FINLSubSrv_SFS5857_Clamp", "IJUAFINLSrv_5857_Clamp", "MAX_LOAD_20", "IJUAFINLSrv_5857_Clamp", "ClampSizeIn", Convert.ToString(bottomPipeNominalDiameter));
                    else if (maxTemperature > 20 && maxTemperature <= 200)
                        maxLoad = FINLAssemblyServices.GetDataByCondition("FINLSubSrv_SFS5857_Clamp", "IJUAFINLSrv_5857_Clamp", "MAX_LOAD_300", "IJUAFINLSrv_5857_Clamp", "ClampSizeIn", Convert.ToString(bottomPipeNominalDiameter));
                    else
                        maxLoad = FINLAssemblyServices.GetDataByCondition("FINLSubSrv_SFS5857_Clamp", "IJUAFINLSrv_5857_Clamp", "MAX_LOAD_480", "IJUAFINLSrv_5857_Clamp", "ClampSizeIn", Convert.ToString(bottomPipeNominalDiameter));
                }
                else if (subBottomAssembly.GetType().FullName.Equals("Ingr.SP3D.Content.Support.Rules.SFS5380_AR3"))
                {
                    bottomPipeNominalDiameter = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Npd, pipeInfo.NominalDiameter.Size, MiddleServiceProvider.UOMMgr.GetUnit(UnitType.Npd, pipeInfo.NominalDiameter.Units), UnitName.NPD_MILLIMETER) / 1000;
                    maxLoad = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5862", "IJOAFINL_MaxLoad", "MaxLoad", "IJUAFINL_PipeND_mm", "PipeND", bottomPipeNominalDiameter - 0.00001, bottomPipeNominalDiameter + 0.00001);
                }
                support.SetPropertyValue(maxLoad, "IJUAFINL_MaxLoad", "MaxLoad");
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in Get Assembly Joints Method." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        #region ICustomHgrBOMDescription Members
        //-----------------------------------------------------------------------------------
        //BOM Description
        //-----------------------------------------------------------------------------------
        public string BOMDescription(BusinessObject SupportOrComponent)
        {
            string bomDesciption = "";
            try
            {
                double heightAdjustment = (double)((PropertyValueDouble)SupportOrComponent.GetPropertyValue("IJUAFINL_HeightAdj", "HeightAdjustment")).PropValue;
                double length = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Vertical) - heightAdjustment;
                string subBottom = (string)((PropertyValueString)SupportOrComponent.GetPropertyValue("IJUAFINLSrv_RodHgrSteel", "Sub_Bottom")).PropValue;
                string subTop = (string)((PropertyValueString)SupportOrComponent.GetPropertyValue("IJUAFINLSrv_RodHgrSteel", "Sub_Top")).PropValue;

                string top = string.Empty, bottom = string.Empty;

                if (subBottom.Equals("FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5380_AR1"))
                    bottom = "AR1";
                else if (subBottom.Equals("FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5380_AR3"))
                    bottom = "AR3";
                else if (subBottom.Equals("FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5380_AR5"))
                    bottom = "AR5";
                else if (subBottom.Equals("FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5380_AR6"))
                    bottom = "AR6";

                PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                double bottomNominalDiameter = pipeInfo.NominalDiameter.Size;
                if (subTop.Equals("FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5380_YK3"))
                    top = "YK3";

                int loadClass = (int)((PropertyValueInt)SupportOrComponent.GetPropertyValue("IJOAFINL_LoadClass", "LoadClass")).PropValue;

                return bomDesciption = "Pipe Hanger SFS 5380 - " + Convert.ToString(loadClass) + " " + top + " - " + Math.Round(length * 1000, 0) + " " + bottom + " DN " + bottomNominalDiameter;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in BOM of Assembly" + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        #endregion
    }
}
