﻿//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   AKDblRodHgrShape.cs
//   FINL_Assy,Ingr.SP3D.Content.Support.Rules.AKDblRodHgrShape
//   Author       :  Hema
//   Creation Date:  14/06/2013
//   Description:    

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   25/06/2013       Hema    CR-CP-224492 Convert HS_FINL_SupAssy to C# .Net
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
    public class AKDblRodHgrShape : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        CustomSupportDefinition subTopAssembly1,subTopAssembly2,subBottomAssembly; 

        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    
                    string topSub = (string)((PropertyValueString)support.GetPropertyValue("IJUAFINLSrv_RodHgrSteel", "Sub_Top")).PropValue;
                    string bottomSub = (string)((PropertyValueString)support.GetPropertyValue("IJUAFINLSrv_RodHgrSteel", "Sub_Bottom")).PropValue;

                    if (!string.IsNullOrEmpty(topSub))
                    {
                        subTopAssembly1 = FINLAssemblyServices.GetAssembly(topSub, support);
                        subTopAssembly2 = FINLAssemblyServices.GetAssembly(topSub, support);
                    }
                    if (!string.IsNullOrEmpty(bottomSub))
                        subBottomAssembly = FINLAssemblyServices.GetAssembly(bottomSub, support);

                    FINLAssemblyServices.SetValueOnPropertyType(subTopAssembly1, "Index", 1);
                    FINLAssemblyServices.SetValueOnPropertyType(subTopAssembly2, "Index", 2);

                    return new Collection<PartInfo>(subTopAssembly1.Parts.Concat(subTopAssembly2.Parts).Concat(subBottomAssembly.Parts).ToList());    //Return the collection of Catalog Parts
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
                    structConnections = subTopAssembly1.SupportingConnections;
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

                ReadOnlyCollection<object> top1Joints = FINLAssemblyServices.GetSubAssemblyJoints(this, subTopAssembly1, oSupCompColl);
                ReadOnlyCollection<object> bottomJoints = FINLAssemblyServices.GetSubAssemblyJoints(this, subBottomAssembly, oSupCompColl);
                ReadOnlyCollection<object> top2Joints = FINLAssemblyServices.GetSubAssemblyJoints(this, subTopAssembly2, oSupCompColl);

                JointHelper.m_oCollOfJoints = new Collection<object>(top1Joints.Concat(top2Joints).Concat(bottomJoints).ToList());

                Collection<ConnectionInfo> subAssyRouteConnections1 = subTopAssembly1.SupportedConnections;
                Collection<ConnectionInfo> subAssyRouteConnections2 = subTopAssembly2.SupportedConnections;
                Collection<ConnectionInfo> subAssyStructConnections = subBottomAssembly.SupportingConnections;
                
                if (subBottomAssembly.GetType().FullName.Equals("Ingr.SP3D.Content.Support.Rules.SFS5380_AK4"))
                {
                    JointHelper.CreateRevoluteJoint(subAssyRouteConnections1[0].PartKey, "TopExThdRH", subAssyStructConnections[0].PartKey, "Hole2", Axis.Z, Axis.Z);
                    JointHelper.CreateRevoluteJoint(subAssyRouteConnections2[0].PartKey, "TopExThdRH", subAssyStructConnections[0].PartKey, "Hole1", Axis.Z, Axis.Z);
                }
                else
                {
                    if (SupportHelper.PlacementType == PlacementType.PlaceByPoint || SupportHelper.PlacementType == PlacementType.PlaceByReference)
                    {
                        if (Configuration == 1 || Configuration == 2)
                        {
                            JointHelper.CreateRigidJoint(subAssyStructConnections[0].PartKey, "InThdRH", subAssyRouteConnections1[0].PartKey, "TopExThdRH", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, 0, 0, 0);
                            JointHelper.CreateRigidJoint(subAssyStructConnections[1].PartKey, "InThdRH", subAssyRouteConnections2[0].PartKey, "TopExThdRH", Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeY, 0, 0, 0);
                        }
                        else
                        {
                            JointHelper.CreateRigidJoint(subAssyStructConnections[0].PartKey, "InThdRH", subAssyRouteConnections1[0].PartKey, "TopExThdRH", Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeY, 0, 0, 0);
                            JointHelper.CreateRigidJoint(subAssyStructConnections[1].PartKey, "InThdRH", subAssyRouteConnections2[0].PartKey, "TopExThdRH", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, 0, 0, 0);
                        }
                    }
                }
                int loadClass = (int)((PropertyValueInt)support.GetPropertyValue("IJOAFINL_LoadClass", "LoadClass")).PropValue;
                PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                double maxTemperature = pipeInfo.MaxDesignTemperature, maxLoad = 0;
                string subBottomAssemblyName = subBottomAssembly.GetType().FullName;

                if (subBottomAssemblyName.Equals("Ingr.SP3D.Content.Support.Rules.SFS5380_AK2"))
                {
                    if (maxTemperature <= 20)
                        maxLoad = FINLAssemblyServices.GetDataByCondition("FINLSubSrv_Hgr_Load", "IJUAFINLSrv_HangerLoad", "MAX_LOAD_20_1", "IJUAFINLSrv_HangerLoad", "LoadClass", Convert.ToString(loadClass)) * 1000;
                    else if (maxTemperature > 20 && maxTemperature <= 200)
                        maxLoad = FINLAssemblyServices.GetDataByCondition("FINLSubSrv_Hgr_Load", "IJUAFINLSrv_HangerLoad", "MAX_LOAD_200_1", "IJUAFINLSrv_HangerLoad", "LoadClass", Convert.ToString(loadClass)) * 1000;
                    else
                        maxLoad = FINLAssemblyServices.GetDataByCondition("FINLSubSrv_Hgr_Load", "IJUAFINLSrv_HangerLoad", "MAX_LOAD_300_1", "IJUAFINLSrv_HangerLoad", "LoadClass", Convert.ToString(loadClass)) * 1000;
                }
                else if (subBottomAssemblyName.Equals("Ingr.SP3D.Content.Support.Rules.SFS5380_AK4"))
                    maxLoad = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5388", "IJUAFINL_MaxLoad", "MaxLoad", "IJUAFINL_LoadClass", "LoadClass", Convert.ToString(loadClass));

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
                PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                double bottomPipeDiameter = pipeInfo.OutsideDiameter, d5384_CE = 0, E = 0, length, rodTakeOut = 0, AK4Th = 0, bottomPipeNominalDiameter;

                string subBottom = (string)((PropertyValueString)SupportOrComponent.GetPropertyValue("IJUAFINLSrv_RodHgrSteel", "Sub_Bottom")).PropValue;
                string subTop1 = (string)((PropertyValueString)SupportOrComponent.GetPropertyValue("IJUAFINLSrv_RodHgrSteel", "Sub_Top")).PropValue;
                double shoeHeight = (double)((PropertyValueDouble)SupportOrComponent.GetPropertyValue("IJUAFINL_ShoeH", "ShoeH")).PropValue;
                double heightAdjustment = (double)((PropertyValueDouble)SupportOrComponent.GetPropertyValue("IJUAFINL_HeightAdj", "HeightAdjustment")).PropValue;
                int loadClass = (int)((PropertyValueInt)SupportOrComponent.GetPropertyValue("IJOAFINL_LoadClass", "LoadClass")).PropValue;

                bottomPipeNominalDiameter = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Npd, pipeInfo.NominalDiameter.Size, MiddleServiceProvider.UOMMgr.GetUnit(UnitType.Npd, pipeInfo.NominalDiameter.Units), UnitName.NPD_MILLIMETER);

                d5384_CE = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5384", "IJUAFINL_C", "C", "IJUAFINL_PipeND_mm", "PipeND", bottomPipeNominalDiameter / 1000 - 0.001, bottomPipeNominalDiameter / 1000 + 0.001);
                E = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5384", "IJUAFINL_E", "E", "IJUAFINL_PipeND_mm", "PipeND", bottomPipeNominalDiameter / 1000 - 0.001, bottomPipeNominalDiameter / 1000 + 0.001);
                d5384_CE = d5384_CE + E;

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

                if (subBottom.Equals("FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5380_AK4") && (SupportHelper.PlacementType == PlacementType.PlaceByReference || (supportingType == "Slab")))
                {
                    rodTakeOut = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5381", "IJUAFINL_TakeOut", "TakeOut", "IJUAFINL_LoadClass", "LoadClass", Convert.ToString(loadClass));
                    AK4Th = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5388", "IJUAFINL_Thickness", "Thickness", "IJUAFINL_LoadClass", "LoadClass", Convert.ToString(loadClass));
                    length = 0.6 + heightAdjustment - rodTakeOut - AK4Th;
                }
                else if (subBottom.Equals("FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5380_AK4"))
                    length = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Vertical) + bottomPipeDiameter / 2 + shoeHeight - heightAdjustment;
                else
                    length = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Vertical) - d5384_CE - heightAdjustment;

                string top = string.Empty, bottom = string.Empty;

                if (subBottom.Equals("FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5380_AK2"))
                    bottom = "AK2";
                else if (subBottom.Equals("FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5380_AK4"))
                    bottom = "AK4";

                if (subTop1.Equals("FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5380_YK3"))
                    top = "YK3";
                return bomDesciption = "Pipe Hanger SFS 5380 - " + Convert.ToString(loadClass) + " 2x" + top + " - " + Math.Round(length * 1000, 0) + " " + bottom + " DN " + bottomPipeNominalDiameter;
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