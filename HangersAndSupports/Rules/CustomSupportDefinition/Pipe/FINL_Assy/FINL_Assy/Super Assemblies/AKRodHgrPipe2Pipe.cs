//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   AKRodHgrPipe2Pipe.cs
//   FINL_Assy,Ingr.SP3D.Content.Support.Rules.AKRodHgrPipe2Pipe
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
using Ingr.SP3D.Support.Middle;
using System.Linq;
namespace Ingr.SP3D.Content.Support.Rules
{
    public class AKRodHgrPipe2Pipe : CustomSupportDefinition, ICustomHgrBOMDescription
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
                if (subMiddleAssembly != null)
                {
                    ReadOnlyCollection<object> middleJoints = FINLAssemblyServices.GetSubAssemblyJoints(this, subMiddleAssembly, oSupCompColl);
                    JointHelper.m_oCollOfJoints = new Collection<object>(topJoints.Concat(middleJoints).Concat(bottomJoints).ToList());
                }
                else
                    JointHelper.m_oCollOfJoints = new Collection<object>(topJoints.Concat(bottomJoints).ToList());
                
                Collection<ConnectionInfo> subAssyRouteConnections = subTopAssembly.SupportingConnections;
                Collection<ConnectionInfo> subAssyStructConnections = subBottomAssembly.SupportingConnections;
              
                if (subBottomAssembly.GetType().FullName == "Ingr.SP3D.Content.Support.Rules.SFS5380_AK6")
                    JointHelper.CreateRevoluteJoint(subAssyRouteConnections[0].PartKey, "TopExThdRH", subAssyStructConnections[0].PartKey, "InThdRH", Axis.Z, Axis.Z);
                else if (subBottomAssembly.GetType().FullName == "Ingr.SP3D.Content.Support.Rules.SFS5380_AK1" && subTopAssembly.GetType().FullName == "Ingr.SP3D.Content.Support.Rules.SFS5380_YK5")
                    JointHelper.CreateRevoluteJoint(subAssyRouteConnections[0].PartKey, "TopExThdRH", subAssyStructConnections[0].PartKey, "InThdRH", Axis.Z, Axis.Z);
                else
                    JointHelper.CreateRevoluteJoint(subAssyRouteConnections[0].PartKey, "TopExThdRH", subAssyStructConnections[0].PartKey, "InThdRH", Axis.Z, Axis.NegativeZ);

                int loadClass = (int)((PropertyValueInt)support.GetPropertyValue("IJOAFINL_LoadClass", "LoadClass")).PropValue;
                PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                double maxTemperature = pipeInfo.MaxDesignTemperature, maxLoad=0;
                string subBottomAssemblyName = subBottomAssembly.GetType().FullName;

                if (subBottomAssemblyName.Equals("Ingr.SP3D.Content.Support.Rules.SFS5380_AK1") || subBottomAssemblyName.Equals("Ingr.SP3D.Content.Support.Rules.SFS5380_AK5") || subBottomAssemblyName.Equals("Ingr.SP3D.Content.Support.Rules.SFS5380_AK6"))
                {
                    if (maxTemperature <= 20)
                        maxLoad = FINLAssemblyServices.GetDataByCondition("FINLSubSrv_Hgr_Load", "IJUAFINLSrv_HangerLoad", "MAX_LOAD_20_1", "IJUAFINLSrv_HangerLoad", "LoadClass", Convert.ToString(loadClass)) * 1000;
                    else if (maxTemperature > 20 && maxTemperature <= 200)
                        maxLoad = FINLAssemblyServices.GetDataByCondition("FINLSubSrv_Hgr_Load", "IJUAFINLSrv_HangerLoad", "MAX_LOAD_200_1", "IJUAFINLSrv_HangerLoad", "LoadClass", Convert.ToString(loadClass)) * 1000;
                    else
                        maxLoad = FINLAssemblyServices.GetDataByCondition("FINLSubSrv_Hgr_Load", "IJUAFINLSrv_HangerLoad", "MAX_LOAD_300_1", "IJUAFINLSrv_HangerLoad", "LoadClass", Convert.ToString(loadClass)) * 1000;
                }
                else if (subBottomAssembly.GetType().FullName.Equals("Ingr.SP3D.Content.Support.Rules.SFS5380_AK3"))
                {
                    if (maxTemperature <= 20)
                        maxLoad = FINLAssemblyServices.GetDataByCondition("FINLSubSrv_Hgr_Load", "IJUAFINLSrv_HangerLoad", "MAX_LOAD_20_2", "IJUAFINLSrv_HangerLoad", "LoadClass", Convert.ToString(loadClass)) * 1000;
                    else if (maxTemperature > 20 && maxTemperature <= 200)
                        maxLoad = FINLAssemblyServices.GetDataByCondition("FINLSubSrv_Hgr_Load", "IJUAFINLSrv_HangerLoad", "MAX_LOAD_200_2", "IJUAFINLSrv_HangerLoad", "LoadClass", Convert.ToString(loadClass)) * 1000;
                    else
                        maxLoad = FINLAssemblyServices.GetDataByCondition("FINLSubSrv_Hgr_Load", "IJUAFINLSrv_HangerLoad", "MAX_LOAD_300_2", "IJUAFINLSrv_HangerLoad", "LoadClass", Convert.ToString(loadClass)) * 1000;
                }

                support.SetPropertyValue(maxLoad,"IJUAFINL_MaxLoad","MaxLoad");
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
                double length = RefPortHelper.DistanceBetweenPorts("Route", "Route_2", PortDistanceType.Vertical);
                string subBottom = (string)((PropertyValueString)SupportOrComponent.GetPropertyValue("IJUAFINLSrv_RodHgrSteel", "Sub_Bottom")).PropValue;
                string subTop = (string)((PropertyValueString)SupportOrComponent.GetPropertyValue("IJUAFINLSrv_RodHgrSteel", "Sub_Top")).PropValue;

                string top = string.Empty, bottom = string.Empty;
                double topNominalDiameter = 0, bottomNominalDiameter;

                if (subBottom.Equals("FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5380_AK1"))
                    bottom = "AK1";
                else if (subBottom.Equals("FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5380_AK3"))
                    bottom = "AK3";
                else if (subBottom.Equals("FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5380_AK5"))
                    bottom = "AK5";
                else if (subBottom == "FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5380_AK6")
                    bottom = "AK6";

                PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                bottomNominalDiameter = pipeInfo.NominalDiameter.Size;
                if (subTop.Equals("FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5380_YK5"))
                {
                    top = "YK5";
                    pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(2);
                    topNominalDiameter = pipeInfo.NominalDiameter.Size;
                }
                    
                int loadClass = (int)((PropertyValueInt)SupportOrComponent.GetPropertyValue("IJOAFINL_LoadClass", "LoadClass")).PropValue;

                return bomDesciption = "Pipe Hanger SFS 5380 - " + Convert.ToString(loadClass) + " " + top + " - DN " + topNominalDiameter + " - " + Math.Round(length * 1000, 0) + " " + bottom + " DN " + bottomNominalDiameter;
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