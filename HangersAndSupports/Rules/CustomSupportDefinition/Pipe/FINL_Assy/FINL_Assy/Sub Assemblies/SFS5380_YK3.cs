//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SFS5380_YK3.cs
//   FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5380_YK3
//   Author       : Rajeswari
//   Creation Date: 27-June-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
// 01-July-2013  Rajeswari CR-CP-224491- Convert HS_FINL_SubAssy to C# .Net
// 20-05-2014    PVK       TR-CP-237153  Problems in placement of .Net HS_FINL_Assy, HS_FINL_SupAssy.
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
    //--------------------------------------------------------------------------------

    public class SFS5380_YK3 : CustomSupportDefinition
    {
        private const string ROD = "ROD_YK3";
        private const string ATTACHMENTA = "ATTACHMENTA_YK3";
        private const string HEXNUT1 = "HEXNUT1_YK3";
        private const string HEXNUT2 = "HEXNUT2_YK3";
        private const string LOGOBJ = "LOGOBJ_YK3";
        private const string LOGOBJ2 = "LOGOBJ2_YK3";
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

                    parts.Add(new PartInfo(ROD + Index, "FINLCmp_SFS5381_" + loadClass));
                    parts.Add(new PartInfo(ATTACHMENTA + Index, "FINLCmp_SFS5385_" + loadClass));
                    parts.Add(new PartInfo(HEXNUT1 + Index, "FINLCmp_SFS4032_" + loadClass));
                    parts.Add(new PartInfo(HEXNUT2 + Index, "FINLCmp_SFS4032_" + loadClass));
                    parts.Add(new PartInfo(LOGOBJ + Index, "Log_Conn_Part_1"));
                    parts.Add(new PartInfo(LOGOBJ2 + Index, "Log_Conn_Part_1"));

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

                double takeOut = 0, boltHeight = 0, dis = 0, pipeDiameter = 0;
                double lengthV = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Vertical);

                takeOut = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5381", "IJUAFINL_TakeOut", "TakeOut", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString());
                boltHeight = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS4032", "IJUAFINL_Thickness", "Thickness", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString());
                double length = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINL_HeightAdj", "HeightAdjustment")).PropValue;

                PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
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

                if (assmDescription.IndexOf("AK2") > 0 || assmDescription.IndexOf("AK4") > 0 || assmDescription.IndexOf("AR2") > 0 || assmDescription.IndexOf("AR4") > 0)
                {
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    {
                        if (Configuration == 1)
                            JointHelper.CreateRigidJoint("-1", "Structure", LOGOBJ + Index, "Connection", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, length, 0, 0);
                        else
                            JointHelper.CreateRigidJoint("-1", "Structure", LOGOBJ + Index, "Connection", Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeX, length, 0, 0);

                        JointHelper.CreatePrismaticJoint(ATTACHMENTA + Index, "Structure", LOGOBJ + Index, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0);

                        JointHelper.CreateRigidJoint(ATTACHMENTA + Index, "Hole", ROD + Index, "BotExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                        JointHelper.CreatePrismaticJoint(ROD + Index, "BotExThdRH", ROD + Index, "TopExThdRH", Plane.YZ, Plane.YZ, Axis.Z, Axis.Z, 0, 0);
                    }
                    else if ((assmDescription.IndexOf("AK4") > 0 || assmDescription.IndexOf("AR4") > 0) && SupportHelper.PlacementType == PlacementType.PlaceByReference)
                    {
                        if (Configuration == 1)
                            JointHelper.CreateRigidJoint("-1", "Structure", LOGOBJ + Index, "Connection", Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeY, length, 0, 0);
                        else
                            JointHelper.CreateRigidJoint("-1", "Structure", LOGOBJ + Index, "Connection", Plane.XY, Plane.NegativeXY, Axis.X, Axis.Y, length, 0, 0);

                        JointHelper.CreatePrismaticJoint(ATTACHMENTA + Index, "Structure", LOGOBJ + Index, "Connection", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0);

                        JointHelper.CreateRigidJoint(ATTACHMENTA + Index, "Hole", ROD + Index, "BotExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                        JointHelper.CreatePrismaticJoint(ROD + Index, "BotExThdRH", ROD + Index, "TopExThdRH", Plane.YZ, Plane.YZ, Axis.Z, Axis.Z, 0, 0);

                    }

                    else
                    {
                        if ((assmDescription.IndexOf("AK4") > 0 || assmDescription.IndexOf("AR4") > 0) && (supportingType == "Slab"))
                        {

                            pipeDiameter = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Npd, pipeInfo.NominalDiameter.Size, MiddleServiceProvider.UOMMgr.GetUnit(UnitType.Npd, pipeInfo.NominalDiameter.Units), UnitName.NPD_MILLIMETER) / 1000;
                            dis = 0.6;

                            componentDictionary[ROD + Index].SetPropertyValue(length + dis, "IJUAHgrOccLength", "Length");

                            if (Configuration == 1 || Configuration == 3)
                                JointHelper.CreatePlanarJoint(ATTACHMENTA + Index, "Structure", "-1", "Structure", Plane.ZX, Plane.XY, 0);
                            else
                                JointHelper.CreatePlanarJoint(ATTACHMENTA + Index, "Structure", "-1", "Structure", Plane.ZX, Plane.NegativeXY, 0);

                            JointHelper.CreateGlobalAxesAlignedJoint(ATTACHMENTA + Index, "Structure", Axis.Z, Axis.Z);
                            JointHelper.CreateRigidJoint(ATTACHMENTA + Index, "Hole", ROD + Index, "BotExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        }
                        else
                        {
                            if (Configuration == 1)
                                JointHelper.CreateRigidJoint("-1", "Structure", LOGOBJ + Index, "Connection", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, length, 0, 0);
                            else
                                JointHelper.CreateRigidJoint("-1", "Structure", LOGOBJ + Index, "Connection", Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeX, length, 0, 0);
                            pipeDiameter = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Npd, pipeInfo.NominalDiameter.Size, MiddleServiceProvider.UOMMgr.GetUnit(UnitType.Npd, pipeInfo.NominalDiameter.Units), UnitName.NPD_MILLIMETER) / 1000;
                            if (assmDescription.IndexOf("AK2") > 0)
                            {
                                dis = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5384", "IJUAFINL_C", "C", "IJUAFINL_PipeND_mm", "PipeND", pipeDiameter);
                                dis = dis + FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5390", "IJUAFINL_A", "A", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString()) - FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5390", "IJUAFINL_X", "X", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString());
                                dis = dis + FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5385", "IJUAFINL_Width", "Width", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString()) / 2 - (FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5385", "IJUAFINL_Thickness", "Thickness", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString()) + FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5385", "IJUAFINL_TakeOut", "TakeOut", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString()));
                            }
                            else if (assmDescription.IndexOf("AR2") > 0)
                            {
                                dis = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5862", "IJUAFINL_C", "C", "IJUAFINL_PipeND_mm", "PipeND", pipeDiameter);
                                dis = dis + FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5390", "IJUAFINL_A", "A", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString()) - FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5390", "IJUAFINL_X", "X", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString());
                                dis = dis + FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5385", "IJUAFINL_Width", "Width", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString()) / 2 - (FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5385", "IJUAFINL_Thickness", "Thickness", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString()) + FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5385", "IJUAFINL_TakeOut", "TakeOut", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString()));
                            }
                            else
                                dis = 0.6;
                            JointHelper.CreateRigidJoint(ROD + Index, "BotExThdRH", ATTACHMENTA + Index, "Hole", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                            componentDictionary[ROD + Index].SetPropertyValue(lengthV - dis - length, "IJUAHgrOccLength", "Length");

                            JointHelper.CreateGlobalAxesAlignedJoint(ROD + Index, "BotExThdRH", Axis.Z, Axis.Z);
                        }
                    }
                }
                else
                {
                    if (SupportHelper.PlacementType == PlacementType.PlaceByStruct)
                    {
                        if (Configuration == 1)
                            JointHelper.CreateRigidJoint("-1", "Structure", ATTACHMENTA + Index, "Structure", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, length, 0, 0);
                       else
                            JointHelper.CreateRigidJoint("-1", "Structure", ATTACHMENTA + Index, "Structure", Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeX, length, 0, 0);

                        JointHelper.CreateRigidJoint(ATTACHMENTA + Index, "Hole", ROD + Index, "BotExThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);

                        JointHelper.CreatePrismaticJoint(ROD + Index, "BotExThdRH", ROD + Index, "TopExThdRH", Plane.YZ, Plane.YZ, Axis.Z, Axis.Z, 0, 0);
                    }
                    else
                    {
                        JointHelper.CreateRigidJoint(ROD + Index, "BotExThdRH", ATTACHMENTA + Index, "Hole", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, 0);
                        pipeDiameter = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Npd, pipeInfo.NominalDiameter.Size, MiddleServiceProvider.UOMMgr.GetUnit(UnitType.Npd, pipeInfo.NominalDiameter.Units), UnitName.NPD_MILLIMETER) / 1000;
                        if (assmDescription.IndexOf("AK1") > 0)
                        {
                            dis = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5371", "IJUAFINL_E", "E", "IJUAFINL_PipeND_mm", "PipeND", pipeDiameter) + FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5371", "IJUAFINL_CtoC", "CtoC", "IJUAFINL_PipeND_mm", "PipeND", pipeDiameter) / 2;
                            dis = dis + FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5389", "IJUAFINL_B", "B", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString()) + FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5389", "IJUAFINL_C", "C", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString()) - FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5389", "IJUAFINL_X", "X", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString());
                            dis = dis + FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5385", "IJUAFINL_Width", "Width", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString()) / 2 - (FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5385", "IJUAFINL_Thickness", "Thickness", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString()) + FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5385", "IJUAFINL_TakeOut", "TakeOut", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString()));
                        }
                        else if (assmDescription.IndexOf("AR1") > 0)
                        {
                            dis = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5857", "IJUAFINL_E", "E", "IJUAFINL_PipeND_mm", "PipeND", pipeDiameter) + FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5857", "IJUAFINL_A", "A", "IJUAFINL_PipeND_mm", "PipeND", pipeDiameter) / 2;
                            dis = dis + FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5389", "IJUAFINL_B", "B", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString()) + FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5389", "IJUAFINL_C", "C", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString()) - FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5389", "IJUAFINL_X", "X", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString());
                            dis = dis + FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5385", "IJUAFINL_Width", "Width", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString()) / 2 - (FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5385", "IJUAFINL_Thickness", "Thickness", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString()) + FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5385", "IJUAFINL_TakeOut", "TakeOut", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString()));
                        }
                        else if (assmDescription.IndexOf("AK3") > 0)
                        {
                            dis = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5370", "IJUAFINL_CtoC", "CtoC", "IJUAFINL_PipeND_mm", "PipeND", pipeDiameter) / 2;
                            dis = dis + FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5384", "IJUAFINL_C", "C", "IJUAFINL_PipeND_mm", "PipeND", pipeDiameter) * Math.Sqrt(2) / 2;
                            dis = dis + FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5390", "IJUAFINL_A", "A", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString()) - FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5390", "IJUAFINL_X", "X", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString());
                            dis = dis + FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5385", "IJUAFINL_Width", "Width", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString()) / 2 - (FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5385", "IJUAFINL_Thickness", "Thickness", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString()) + FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5385", "IJUAFINL_TakeOut", "TakeOut", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString()));
                        }
                        else if (assmDescription.IndexOf("AR3") > 0)
                        {
                            dis = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5856", "IJUAFINL_A", "A", "IJUAFINL_PipeND_mm", "PipeND", pipeDiameter) / 2;
                            dis = dis + FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5862", "IJUAFINL_C", "C", "IJUAFINL_PipeND_mm", "PipeND", pipeDiameter) * Math.Sqrt(2) / 2;
                            dis = dis + FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5390", "IJUAFINL_A", "A", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString()) - FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5390", "IJUAFINL_X", "X", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString());
                            dis = dis + FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5385", "IJUAFINL_Width", "Width", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString()) / 2 - (FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5385", "IJUAFINL_Thickness", "Thickness", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString()) + FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5385", "IJUAFINL_TakeOut", "TakeOut", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString()));
                        }
                        else if (assmDescription.IndexOf("AK5") > 0)
                        {
                            dis = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5371", "IJUAFINL_CtoC", "CtoC", "IJUAFINL_PipeND_mm", "PipeND", pipeDiameter) / 2 + FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5857", "IJUAFINL_E", "E", "IJUAFINL_PipeND_mm", "PipeND", pipeDiameter);
                            dis = dis + FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5393B", "IJUAFINL_L", "L", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString());
                            dis = dis + FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5390", "IJUAFINL_A", "A", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString()) - FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5390", "IJUAFINL_X", "X", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString());
                            dis = dis + FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5385", "IJUAFINL_Width", "Width", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString()) / 2 - (FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5385", "IJUAFINL_Thickness", "Thickness", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString()) + FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5385", "IJUAFINL_TakeOut", "TakeOut", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString()));
                        }
                        else if (assmDescription.IndexOf("AR5") > 0)
                        {
                            dis = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5857", "IJUAFINL_A", "A", "IJUAFINL_PipeND_mm", "PipeND", pipeDiameter) / 2 + FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5857", "IJUAFINL_E", "E", "IJUAFINL_PipeND_mm", "PipeND", pipeDiameter);
                            dis = dis + FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5393B", "IJUAFINL_L", "L", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString());
                            dis = dis + FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5390", "IJUAFINL_A", "A", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString()) - FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5390", "IJUAFINL_X", "X", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString());
                            dis = dis + FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5385", "IJUAFINL_Width", "Width", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString()) / 2 - (FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5385", "IJUAFINL_Thickness", "Thickness", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString()) + FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5385", "IJUAFINL_TakeOut", "TakeOut", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString()));
                        }
                        else if (assmDescription.IndexOf("AK6") > 0)
                        {
                            dis = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5371", "IJUAFINL_CtoC", "CtoC", "IJUAFINL_PipeND_mm", "PipeND", pipeDiameter) / 2 + FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5857", "IJUAFINL_E", "E", "IJUAFINL_PipeND_mm", "PipeND", pipeDiameter);
                            dis = dis + FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5389", "IJUAFINL_B", "B", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString()) + FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5389", "IJUAFINL_C", "C", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString()) - FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5389", "IJUAFINL_X", "X", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString()); ;
                            dis = dis + FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5391", "IJUAFINL_L", "L", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString()) + FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5391", "IJUAFINL_A", "A", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString()) - 2 * FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5391", "IJUAFINL_X", "X", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString());
                            dis = dis + FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5385", "IJUAFINL_Width", "Width", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString()) / 2 - (FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5385", "IJUAFINL_Thickness", "Thickness", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString()) + FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5385", "IJUAFINL_TakeOut", "TakeOut", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString()));
                        }
                        else if (assmDescription.IndexOf("AR6") > 0)
                        {
                            dis = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5857", "IJUAFINL_A", "A", "IJUAFINL_PipeND_mm", "PipeND", pipeDiameter) / 2 + FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5857", "IJUAFINL_E", "E", "IJUAFINL_PipeND_mm", "PipeND", pipeDiameter);
                            dis = dis + FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5389", "IJUAFINL_B", "B", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString()) + FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5389", "IJUAFINL_C", "C", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString()) - FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5389", "IJUAFINL_X", "X", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString()); ;
                            dis = dis + FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5391", "IJUAFINL_L", "L", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString()) + FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5391", "IJUAFINL_A", "A", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString()) - 2 * FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5391", "IJUAFINL_X", "X", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString());
                            dis = dis + FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5385", "IJUAFINL_Width", "Width", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString()) / 2 - (FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5385", "IJUAFINL_Thickness", "Thickness", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString()) + FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5385", "IJUAFINL_TakeOut", "TakeOut", "IJUAFINL_LoadClass", "LoadClass", loadClass.ToString()));
                        }
                        else
                            dis = 0.6;
                        componentDictionary[ROD + Index].SetPropertyValue(lengthV - dis - length, "IJUAHgrOccLength", "Length");

                        JointHelper.CreateGlobalAxesAlignedJoint(ROD + Index, "BotExThdRH", Axis.Z, Axis.Z);
                    }
                }
                JointHelper.CreateRigidJoint(ROD + Index, "BotExThdRH", HEXNUT1 + Index, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -takeOut, 0, 0);

                JointHelper.CreateRigidJoint(ROD + Index, "BotExThdRH", HEXNUT2 + Index, "InThdRH", Plane.XY, Plane.XY, Axis.X, Axis.X, -takeOut + boltHeight, 0, 0);
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
                    routeConnections.Add(new ConnectionInfo(ROD + Index, 1)); // partindex, routeindex

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
                    structConnections.Add(new ConnectionInfo(ATTACHMENTA + Index, 1)); // partindex, routeindex

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

