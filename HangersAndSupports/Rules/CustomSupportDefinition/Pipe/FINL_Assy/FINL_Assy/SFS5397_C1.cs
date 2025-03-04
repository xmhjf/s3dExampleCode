//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SFS5397_C1.cs
//   FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5397_C1
//   Author       :  Vijaya
//   Creation Date:  27.Jun.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   27-Jun-2013     Vijaya   CR-CP-224485- Convert HS_FINL_Assy to C# .Net 
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
    [SymbolVersion("1.0.0.0")]
    public class SFS5397_C1 : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        private const string HORSTEEL5397C1 = "HORSTEEL5397_C1";
        private const string VERTSTEEL5397C1 = "VERTSTEEL5397_C1";
        private const string PLATE15397C1 = "PLATE15397_C1";
        private const string PLATE25397C1 = "PLATE25397_C1";
        private const int STEELDENSITYKGPERM = 7900;
        private string size2;
        double pipeDiameter = 0.0, shoeHeight = 0.0, overHang = 0.0, horizontalLength, cDim;

        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();

                    PropertyValueCodelist sizeCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJUAFINLSize2", "Size2");
                    size2 = sizeCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(sizeCodeList.PropValue).DisplayName;

                    shoeHeight = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINLShoeH", "ShoeH")).PropValue;
                    overHang = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINLOverHang", "OverHang")).PropValue;
                    cDim = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINLC_Dim", "C_Dim")).PropValue;

                    parts.Add(new PartInfo(HORSTEEL5397C1, "Util_Fixed_Box_Metric_1"));
                    parts.Add(new PartInfo(VERTSTEEL5397C1, "Util_Fixed_Box_Metric_1"));
                    parts.Add(new PartInfo(PLATE15397C1, "Util_Plate_Metric_1"));
                    parts.Add(new PartInfo(PLATE25397C1, "Util_Plate_Metric_1"));

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
        //Get Assembly Joints
        //-----------------------------------------------------------------------------------
        public override void ConfigureSupport(Collection<SupportComponent> oSupCompColl)
        {
            try
            {
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;

                Dictionary<string, SupportComponent> componentDictionary = SupportHelper.SupportComponentDictionary;
                // To get Pipe Nominal Diameter
                PipeObjectInfo pipeInfo = (PipeObjectInfo)SupportedHelper.SupportedObjectInfo(1);
                pipeDiameter = pipeInfo.OutsideDiameter + pipeInfo.InsulationThickness * 2;

                NominalDiameter[] diameterArray = GetNominalDiameterArray(new double[] { 17, 90, 175, 225, 375, 450, 550, 650, 750, 850, 950, 1050, 1100, 1150 }, "mm");

                NominalDiameter minNominalDiameter = new NominalDiameter(), maxNominalDiameter = new NominalDiameter();
                minNominalDiameter.Size = 10;
                minNominalDiameter.Units = "mm";
                maxNominalDiameter.Size = 1200;
                maxNominalDiameter.Units = "mm";

                if (IsPipeSizeValid(pipeInfo.NominalDiameter, minNominalDiameter, maxNominalDiameter, diameterArray) == false)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Pipe size not valid.", "", "SFS5397_C1.cs", 111);
                    return;
                }

                double maxLength = 0.0, MA = 0.0, MP = 0.0, FQ = 0.0, FU = 0.0, plateThickness = 0.0, plateWidth = 0.0, plateDepth = 0.0;
                PropertyValue TSSize;
                Dictionary<String, Object> parameter = new Dictionary<String, Object>();
                parameter.Add("Size2", size2);
                GenericHelper.GetDataByRule("FINLSrv_SFS5397_C1_Dim", "IJUAFINLSrv_SFS5397_C1_Dim", "TS_Size", parameter, out TSSize);
                string tsSize = TSSize.ToString();


                maxLength = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5397_C1_Dim", "IJUAFINLSrv_SFS5397_C1_Dim", "Max_Len", "IJUAFINLSrv_SFS5397_C1_Dim", "Size2", size2);
                MA = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5397_C1_Dim", "IJUAFINLSrv_SFS5397_C1_Dim", "MA", "IJUAFINLSrv_SFS5397_C1_Dim", "Size2", size2);
                MP = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5397_C1_Dim", "IJUAFINLSrv_SFS5397_C1_Dim", "MP", "IJUAFINLSrv_SFS5397_C1_Dim", "Size2", size2);
                FQ = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5397_C1_Dim", "IJUAFINLSrv_SFS5397_C1_Dim", "FQ", "IJUAFINLSrv_SFS5397_C1_Dim", "Size2", size2);
                FU = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5397_C1_Dim", "IJUAFINLSrv_SFS5397_C1_Dim", "FU", "IJUAFINLSrv_SFS5397_C1_Dim", "Size2", size2);
                plateThickness = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5397_C1_Dim", "IJUAFINLSrv_SFS5397_C1_Dim", "Plate_T", "IJUAFINLSrv_SFS5397_C1_Dim", "Size2", size2);
                plateWidth = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5397_C1_Dim", "IJUAFINLSrv_SFS5397_C1_Dim", "Plate_W", "IJUAFINLSrv_SFS5397_C1_Dim", "Size2", size2);
                plateDepth = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5397_C1_Dim", "IJUAFINLSrv_SFS5397_C1_Dim", "Plate_D", "IJUAFINLSrv_SFS5397_C1_Dim", "Size2", size2);


                support.SetPropertyValue(MA, "IJUAFINLMPMA", "MA");
                support.SetPropertyValue(MP, "IJUAFINLMPMA", "MP");
                support.SetPropertyValue(FQ, "IJUAFINLFQFU", "FQ");
                support.SetPropertyValue(FU, "IJUAFINLFQFU", "FU");


                double tWidth, tFlangeThickness, tWebThickness, tDepth, verticalLength = 350.00 / 1000.00;
                FINLAssemblyServices.GetCrossSectionDimensions("Euro", "HSSR", tsSize, out  tWidth, out   tFlangeThickness, out  tWebThickness, out  tDepth);


                if (overHang < 0)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Overhang must a positive value.", "", "SFS5397_C1.cs", 145);
                    return;
                }
                double structDistance = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Horizontal);
                horizontalLength = structDistance + overHang + cDim - plateThickness;

                if (overHang < pipeDiameter / 2)
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Overhang must exceed pipe outside dia:" + (pipeDiameter / 2 * 1000).ToString() + "mm.", "", "SFS5397_C1.cs", 152);


                if (structDistance + overHang > maxLength)
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Maximum allowable length of " + (maxLength * 1000).ToString() + "mm exceeded.", "", "SFS5397_C1.cs", 156);


                if (cDim < 150.00 / 1000.00)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Dimension C must exceed 100 mm.", "", "SFS5397_C1.cs", 161);
                    return;
                }

                if (structDistance - tWidth < pipeDiameter / 2)
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Distance between Pipe CL and structure must exceed " + (tWidth * 1000 + pipeDiameter / 2 * 1000).ToString() + "mm.", "", "SFS5397_C1.cs", 166);

                double horSteelWeight = (tWidth * tDepth * horizontalLength - ((tWidth - 2.0 * 4.0 / 1000.00) * (tDepth - 2.0 * 4.0 / 1000.00) * horizontalLength)) * STEELDENSITYKGPERM;
                double vertSteelWeight = (tWidth * tDepth * verticalLength - ((tWidth - 2.0 * 4.0 / 1000.00) * (tDepth - 2.0 * 4.0 / 1000.00) * verticalLength)) * STEELDENSITYKGPERM;

                string inputBOM = tsSize + " Length:" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, horizontalLength, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);
                componentDictionary[HORSTEEL5397C1].SetPropertyValue(tWidth, "IJOAHgrUtilMetricWidth", "Width");
                componentDictionary[HORSTEEL5397C1].SetPropertyValue(tDepth, "IJOAHgrUtilMetricDepth", "Depth");
                componentDictionary[HORSTEEL5397C1].SetPropertyValue(horizontalLength, "IJOAHgrUtilMetricL", "L");
                componentDictionary[HORSTEEL5397C1].SetPropertyValue(inputBOM, "IJOAHgrUtilMetricBomDesc", "InputBomDesc");
                componentDictionary[HORSTEEL5397C1].SetPropertyValue(horSteelWeight, "IJOAHgrUtilMetricInWt", "InputWeight");

                inputBOM = tsSize + ", Length: " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, verticalLength, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);

                componentDictionary[VERTSTEEL5397C1].SetPropertyValue(tWidth, "IJOAHgrUtilMetricWidth", "Width");
                componentDictionary[VERTSTEEL5397C1].SetPropertyValue(tDepth, "IJOAHgrUtilMetricDepth", "Depth");
                componentDictionary[VERTSTEEL5397C1].SetPropertyValue(verticalLength, "IJOAHgrUtilMetricL", "L");
                componentDictionary[VERTSTEEL5397C1].SetPropertyValue(inputBOM, "IJOAHgrUtilMetricBomDesc", "InputBomDesc");
                componentDictionary[VERTSTEEL5397C1].SetPropertyValue(vertSteelWeight, "IJOAHgrUtilMetricInWt", "InputWeight");


                string plateBOM = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, plateThickness, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + " Steel Plate " +
                   MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, plateWidth, UnitName.DISTANCE_MILLIMETER) + "x" +
                   MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, plateThickness, UnitName.DISTANCE_MILLIMETER);

                componentDictionary[PLATE15397C1].SetPropertyValue(plateWidth, "IJOAHgrUtilMetricWidth", "Width");
                componentDictionary[PLATE15397C1].SetPropertyValue(plateDepth, "IJOAHgrUtilMetricDepth", "Depth");
                componentDictionary[PLATE15397C1].SetPropertyValue(plateThickness, "IJOAHgrUtilMetricT", "T");
                componentDictionary[PLATE15397C1].SetPropertyValue(plateBOM, "IJOAHgrUtilMetricBomDesc", "InputBomDesc");

                componentDictionary[PLATE25397C1].SetPropertyValue(plateWidth, "IJOAHgrUtilMetricWidth", "Width");
                componentDictionary[PLATE25397C1].SetPropertyValue(plateDepth, "IJOAHgrUtilMetricDepth", "Depth");
                componentDictionary[PLATE25397C1].SetPropertyValue(plateThickness, "IJOAHgrUtilMetricT", "T");
                componentDictionary[PLATE25397C1].SetPropertyValue(plateBOM, "IJOAHgrUtilMetricBomDesc", "InputBomDesc");

                double theta1 = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.Y, "Structure", PortAxisType.X, OrientationAlong.Global_Z);

                if (Math.Abs(theta1) < Math.PI / 2)
                {
                    //create graphics
                    if (Configuration == 1)
                    {
                        JointHelper.CreateRigidJoint("-1", "Route", HORSTEEL5397C1, "StartOther", Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeX, structDistance + cDim - plateThickness, -pipeDiameter / 2 - shoeHeight - tWidth / 2, 0);
                        JointHelper.CreateRigidJoint("-1", "Route", PLATE15397C1, "BotStructure", Plane.XY, Plane.ZX, Axis.X, Axis.X, structDistance + cDim, -pipeDiameter / 2 - shoeHeight - tDepth - plateThickness, 0);
                    }
                    else
                    {
                        JointHelper.CreateRigidJoint("-1", "Route", HORSTEEL5397C1, "StartOther", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, structDistance + cDim - plateThickness, -pipeDiameter / 2 - shoeHeight - tWidth / 2, 0);
                        JointHelper.CreateRigidJoint("-1", "Route", PLATE15397C1, "BotStructure", Plane.XY, Plane.ZX, Axis.X, Axis.X, structDistance + cDim, -pipeDiameter / 2 - shoeHeight, 0);
                    }
                    JointHelper.CreateRigidJoint(HORSTEEL5397C1, "StartOther", VERTSTEEL5397C1, "StartOther", Plane.XY, Plane.ZX, Axis.X, Axis.X, cDim + tWidth / 2, -tWidth / 2, 0);
                    JointHelper.CreateRigidJoint(VERTSTEEL5397C1, "StartOther", PLATE25397C1, "TopStructure", Plane.XY, Plane.ZX, Axis.X, Axis.X, 250.00 / 1000.00 + plateThickness, -tWidth / 2, 0);
                }
                else
                {
                    if (Configuration == 1)
                    {
                        JointHelper.CreateRigidJoint("-1", "Route", HORSTEEL5397C1, "StartOther", Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeX, structDistance + cDim - plateThickness, pipeDiameter / 2 + shoeHeight + tWidth / 2, 0);
                        JointHelper.CreateRigidJoint("-1", "Route", PLATE15397C1, "TopStructure", Plane.XY, Plane.ZX, Axis.X, Axis.X, structDistance + cDim, pipeDiameter / 2 + shoeHeight, 0);
                    }
                    else
                    {
                        JointHelper.CreateRigidJoint("-1", "Route", HORSTEEL5397C1, "StartOther", Plane.XY, Plane.NegativeXY, Axis.X, Axis.X, structDistance + cDim - plateThickness, pipeDiameter / 2 + shoeHeight + tWidth / 2, 0);
                        JointHelper.CreateRigidJoint("-1", "Route", PLATE15397C1, "TopStructure", Plane.XY, Plane.ZX, Axis.X, Axis.X, structDistance + cDim, pipeDiameter / 2 + shoeHeight + tDepth + plateThickness, 0);
                    }
                    JointHelper.CreateRigidJoint(HORSTEEL5397C1, "StartOther", VERTSTEEL5397C1, "StartOther", Plane.XY, Plane.ZX, Axis.X, Axis.X, cDim + tWidth / 2, -tWidth / 2, 0);
                    JointHelper.CreateRigidJoint(VERTSTEEL5397C1, "StartOther", PLATE25397C1, "TopStructure", Plane.XY, Plane.ZX, Axis.X, Axis.X, 250.00 / 1000.00 + plateThickness, -tWidth / 2, 0);
                }

            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in Get Assembly Joints Method of FINL_Assy." + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
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
                    routeConnections.Add(new ConnectionInfo(HORSTEEL5397C1, 1));
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
                    structConnections.Add(new ConnectionInfo(VERTSTEEL5397C1, 1));
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
        #region ICustomHgrBOMDescription Members
        //-----------------------------------------------------------------------------------
        //BOM Description
        //-----------------------------------------------------------------------------------
        public string BOMDescription(BusinessObject SupportOrComponent)
        {
            string bomString = "";
            try
            {
                Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                PropertyValueCodelist sizeCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJUAFINLSize2", "Size2");
                bomString = "Console support C1 -" + sizeCodeList.PropValue + " SFS 5397";

                return bomString;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in BOM of Assembly - RectHorTypeB" + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        #endregion
    }
}

