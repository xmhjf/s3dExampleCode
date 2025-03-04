//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   ARRodHgrTopStr.cs
//   FINL_Assy,Ingr.SP3D.Content.Support.Rules.ARRodHgrTopStr
//   Author       :  BS
//   Creation Date:  14/06/2013
//   Description:    

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  14/06/2013    BS  CR-CP-224472-Convert HS_FINL_SupAssy to C# .Net
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Support.Middle;
using System.Linq;
namespace Ingr.SP3D.Content.Support.Rules
{
    public class ARRodHgrTopStr : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        CustomSupportDefinition subTopAssembly,subMiddleAssembly,subBottomAssembly;

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

                    Collection<PartInfo> topParts = subTopAssembly.Parts;
                    if (!string.IsNullOrEmpty(subMiddle))
                        return new Collection<PartInfo>(subTopAssembly.Parts.Concat(subMiddleAssembly.Parts).Concat(subBottomAssembly.Parts).ToList());
                    else
                        return new Collection<PartInfo>(subTopAssembly.Parts.Concat(subBottomAssembly.Parts).ToList());
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
        //Get Route Connections
        //-----------------------------------------------------------------------------------
        public override Collection<ConnectionInfo> SupportedConnections
        {
            get
            {
                try
                {
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
                if (subMiddleAssembly != null)
                {
                    ReadOnlyCollection<object> middleJoints = FINLAssemblyServices.GetSubAssemblyJoints(this, subMiddleAssembly, oSupCompColl);
                    JointHelper.m_oCollOfJoints = new Collection<object>(topJoints.Concat(middleJoints).Concat(bottomJoints).ToList());
                }
                else
                    JointHelper.m_oCollOfJoints = new Collection<object>(topJoints.Concat(bottomJoints).ToList());

                Collection<ConnectionInfo> iSubAssyRouteCon = subTopAssembly.SupportedConnections;
                Collection<ConnectionInfo> iSubAssyStructCon = subBottomAssembly.SupportingConnections;
                if (subBottomAssembly.GetType().FullName == "Ingr.SP3D.Content.Support.Rules.SFS5380_AR6")
                    JointHelper.CreateRevoluteJoint(iSubAssyRouteCon[0].PartKey, "BotExThdRH", iSubAssyStructCon[0].PartKey, "InThdRH", Axis.Z, Axis.NegativeZ);
                else if (subBottomAssembly.GetType().FullName == "Ingr.SP3D.Content.Support.Rules.SFS5380_AR1" && subTopAssembly.GetType().FullName == "Ingr.SP3D.Content.Support.Rules.SFS5380_YK4")
                    JointHelper.CreateRevoluteJoint(iSubAssyRouteCon[0].PartKey, "BotExThdRH", iSubAssyStructCon[0].PartKey, "InThdRH", Axis.Z, Axis.NegativeZ);
                else
                    JointHelper.CreateRevoluteJoint(iSubAssyRouteCon[0].PartKey, "BotExThdRH", iSubAssyStructCon[0].PartKey, "InThdRH", Axis.Z, Axis.Z);


                int loadClass = (int)((PropertyValueInt)support.GetPropertyValue("IJOAFINL_LoadClass", "LoadClass")).PropValue;
                PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                double maxTemperature = pipeInfo.MaxDesignTemperature;

                double maxLoad = 0;
                string subBottomAssemblyName = subBottomAssembly.GetType().FullName;
                if (subBottomAssemblyName == "Ingr.SP3D.Content.Support.Rules.SFS5380_AR1" || subBottomAssemblyName == "Ingr.SP3D.Content.Support.Rules.SFS5380_AR5" || subBottomAssemblyName == "Ingr.SP3D.Content.Support.Rules.SFS5380_AR6")
                {
                    double pipeDiameter = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Npd, pipeInfo.NominalDiameter.Size, MiddleServiceProvider.UOMMgr.GetUnit(UnitType.Npd, pipeInfo.NominalDiameter.Units), UnitName.NPD_INCH);
                        if (maxTemperature <= 20)
                            maxLoad = FINLAssemblyServices.GetDataByCondition("FINLSubSrv_SFS5857_Clamp", "IJUAFINLSrv_5857_Clamp", "MAX_LOAD_20", "IJUAFINLSrv_5857_Clamp", "ClampSizeIn", pipeDiameter.ToString());
                        else if (maxTemperature > 20 && maxTemperature <= 200)
                            maxLoad = FINLAssemblyServices.GetDataByCondition("FINLSubSrv_SFS5857_Clamp", "IJUAFINLSrv_5857_Clamp", "MAX_LOAD_300", "IJUAFINLSrv_5857_Clamp", "ClampSizeIn", pipeDiameter.ToString());
                        else
                            maxLoad = FINLAssemblyServices.GetDataByCondition("FINLSubSrv_SFS5857_Clamp", "IJUAFINLSrv_5857_Clamp", "MAX_LOAD_480", "IJUAFINLSrv_5857_Clamp", "ClampSizeIn", pipeDiameter.ToString());
                }
                else if(subBottomAssemblyName == "Ingr.SP3D.Content.Support.Rules.SFS5380_AR3")
                {
                    double pipeDiameter = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Npd, pipeInfo.NominalDiameter.Size, MiddleServiceProvider.UOMMgr.GetUnit(UnitType.Npd, pipeInfo.NominalDiameter.Units), UnitName.NPD_MILLIMETER) / 1000;
                    maxLoad = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5862", "IJOAFINL_MaxLoad", "MaxLoad", "IJUAFINL_PipeND_mm", "PipeND", (pipeDiameter - 0.0001), (pipeDiameter + 0.0001));
                }
                support.SetPropertyValue(maxLoad, "IJUAFINL_MaxLoad", "MaxLoad");
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in Get Assembly Joints Method of ARRodHgrTopStr." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
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
                double length = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Vertical);
                string subBottom = (string)((PropertyValueString)SupportOrComponent.GetPropertyValue("IJUAFINLSrv_RodHgrSteel", "Sub_Bottom")).PropValue;
                string subTop = (string)((PropertyValueString)SupportOrComponent.GetPropertyValue("IJUAFINLSrv_RodHgrSteel", "Sub_Top")).PropValue;

                string top = string.Empty, bottom = string.Empty;

                if (subBottom == "FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5380_AR1")
                    bottom = "AR1";
                else if (subBottom == "FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5380_AR3")
                    bottom = "AR3";
                else if (subBottom == "FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5380_AR5")
                    bottom = "AR5";
                else if (subBottom == "FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5380_AR6")
                    bottom = "AR6";

                PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                double bottomNominalDiameter = pipeInfo.NominalDiameter.Size;
                if (subTop == "FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5380_YK4")
                    top = "YK4";

                int loadClass = (int)((PropertyValueInt)SupportOrComponent.GetPropertyValue("IJOAFINL_LoadClass", "LoadClass")).PropValue;

                return bomDesciption = "Pipe Hanger SFS 5380 - " + Convert.ToString(loadClass) + " " + top + " - " + Math.Round(length * 1000, 0) + " " + bottom + " DN " + bottomNominalDiameter;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in BOM of Assembly - ARRodHgrTopStr" + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        #endregion
    }
}