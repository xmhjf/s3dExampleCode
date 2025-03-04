//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SFS5365_Anchor_C.cs
//   FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5365_Anchor_C
//   Author       :  BS
//   Creation Date:  04/07/2013
//   Description:    

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//  04/07/2013        BS    CR-CP-224485- Converted HS_FINL_Assy to C# .Net
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;
using Ingr.SP3D.Support.Middle;


namespace Ingr.SP3D.Content.Support.Rules
{
    public class SFS5365_Anchor_C : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        private const string UBolt = "UBolt";
        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();
                    CatalogStructHelper catalogHelper = new CatalogStructHelper();
                    PartClass finlCmp_SFS5369PartClass = (PartClass)catalogHelper.GetPartClass("FINLCmp_SFS5369"); ;

                    //Use the default selection rule to get a catalog part for each part class
                    string partselection = finlCmp_SFS5369PartClass.GetPropertyValue("IJHgrPartClass", "PartSelectionRule").ToString();
                    parts.Add(new PartInfo(UBolt, "FINLCmp_SFS5369", partselection));

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
        //Get Route Connections
        //-----------------------------------------------------------------------------------
        public override Collection<ConnectionInfo> SupportedConnections
        {
            get
            {
                try
                {
                    Collection<ConnectionInfo> routeConnections = new Collection<ConnectionInfo>();
                    routeConnections.Add(new ConnectionInfo(UBolt, 1));      //partindex, routeindex
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
                    structConnections.Add(new ConnectionInfo(UBolt, 1));      //partindex, routeindex
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
                double npdMetric;
                //to change NomPipeDia to metric unit
                PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                double pipeDiameter = pipeInfo.OutsideDiameter - 2 * pipeInfo.InsulationThickness;
                if (!pipeInfo.NominalDiameter.Units.Equals("mm"))
                    npdMetric = MiddleServiceProvider.UOMMgr.ConvertUnitToUnit(UnitType.Npd, pipeInfo.NominalDiameter.Size, MiddleServiceProvider.UOMMgr.GetUnit(UnitType.Npd, pipeInfo.NominalDiameter.Units), UnitName.NPD_MILLIMETER) / 1000;
                else
                    npdMetric = pipeInfo.NominalDiameter.Size / 1000;
                //check valid pipe size
                NominalDiameter minNominalDiameter = new NominalDiameter();
                minNominalDiameter.Size = 10;
                minNominalDiameter.Units = "mm";
                NominalDiameter maxNominalDiameter = new NominalDiameter();
                maxNominalDiameter.Size = 500;
                maxNominalDiameter.Units = "mm";
                NominalDiameter[] diameterArray = GetNominalDiameterArray(new double[] { 17, 90, 175, 225, 375, 450 }, "mm");

                //check valid pipe size
                if (IsPipeSizeValid(pipeInfo.NominalDiameter, minNominalDiameter, maxNominalDiameter, diameterArray) == false)
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Pipe size not valid.", "", "SFS5365_Anchor_C.cs", 134);

                double E = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5369", "IJUAFINL_E", "E", "IJUAFINL_PipeND_mm", "PipeND", npdMetric - 0.003, npdMetric + 0.003);
                double D = FINLAssemblyServices.GetDataByCondition("FINLCmp_SFS5369", "IJUAFINL_D", "D", "IJUAFINL_PipeND_mm", "PipeND", npdMetric - 0.003, npdMetric + 0.003);
                double gap = E / 2 - D / 2 - pipeDiameter / 2;

                JointHelper.CreateRigidJoint("-1", "Route", UBolt, "Route", Plane.XY, Plane.XY, Axis.X, Axis.Y, gap, 0, 0);

            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in Get Assembly Joints Method of SFS5365_Anchor_C." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        #region ICustomHgrBOMDescription Members
        //-----------------------------------------------------------------------------------
        //BOM Description
        //-----------------------------------------------------------------------------------
        public string BOMDescription(BusinessObject SupportOrComponent)
        {
            string BOMString = "";
            try
            {
                PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                return BOMString = "Anchor C SFS 5365 DN " + pipeInfo.NominalDiameter.Size;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in BOM of Assembly - SFS5365_Anchor_C" + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        #endregion


    }

}