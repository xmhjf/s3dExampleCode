//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SFS5397_A1.cs
//   FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5397_A1
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
    public class SFS5397_A1 : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        private const string HORSTEEL5397A1 = "HORSTEEL5397_A1";
        private const string L5397A1 = "L5397_A1";
        private const string PLATE5397A1 = "PLATE5397_A1";
        private const string VERTSTEEL5397A1 = "VERTSTEEL5397_A1";
        private const int STEELDENSITYKGPERM = 7900;
        private string size2;
        double pipeDiameter = 0.0, shoeHeight = 0.0, overHang = 0.0, horizontalLength, structDistance;

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
                    structDistance = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Horizontal);
                    horizontalLength = structDistance + overHang - 10.00 / 1000.00;

                    parts.Add(new PartInfo(HORSTEEL5397A1, "Util_Fixed_Box_Metric_1"));
                    parts.Add(new PartInfo(L5397A1, "Utility_GENERIC_L_1"));
                    parts.Add(new PartInfo(PLATE5397A1, "Util_Plate_Metric_1"));

                    if (horizontalLength >= 500.00 / 1000.00)// with clamps                    
                        parts.Add(new PartInfo(VERTSTEEL5397A1, "Utility_CUTBACK_TS1_1"));

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
                //To get Pipe Nominal Diameter
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
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Pipe size not valid.", "", "SFS5397_A1.cs", 114);
                    return;
                }

                double maxLength = 0.0, K = 0.0, MA = 0.0, MP = 0.0, FQ = 0.0, FU = 0.0, tWidth, tFlangeThickness, tWebThickness, tDepth, lWidth, lFlangeThickness, lWebThickness, lDepth;
                PropertyValue TSSize, LSize;
                Dictionary<String, Object> parameter = new Dictionary<String, Object>();
                parameter.Add("Size2", size2);
                GenericHelper.GetDataByRule("FINLSrv_SFS5397_A1_Dim", "IJUAFINLSrv_SFS5397_A1_Dim", "Arm_Size", parameter, out TSSize);
                GenericHelper.GetDataByRule("FINLSrv_SFS5397_A1_Dim", "IJUAFINLSrv_SFS5397_A1_Dim", "L_Size", parameter, out LSize);
                string tsSize = TSSize.ToString(), lSize = LSize.ToString();

                maxLength = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5397_A1_Dim", "IJUAFINLSrv_SFS5397_A1_Dim", "Max_Len", "IJUAFINLSrv_SFS5397_A1_Dim", "Size2", size2);
                K = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5397_A1_Dim", "IJUAFINLSrv_SFS5397_A1_Dim", "K", "IJUAFINLSrv_SFS5397_A1_Dim", "Size2", size2);
                MA = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5397_A1_Dim", "IJUAFINLSrv_SFS5397_A1_Dim", "MA", "IJUAFINLSrv_SFS5397_A1_Dim", "Size2", size2);
                MP = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5397_A1_Dim", "IJUAFINLSrv_SFS5397_A1_Dim", "MP", "IJUAFINLSrv_SFS5397_A1_Dim", "Size2", size2);
                FQ = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5397_A1_Dim", "IJUAFINLSrv_SFS5397_A1_Dim", "FQ", "IJUAFINLSrv_SFS5397_A1_Dim", "Size2", size2);
                FU = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5397_A1_Dim", "IJUAFINLSrv_SFS5397_A1_Dim", "FU", "IJUAFINLSrv_SFS5397_A1_Dim", "Size2", size2);

                support.SetPropertyValue(MA, "IJUAFINLMPMA", "MA");
                support.SetPropertyValue(MP, "IJUAFINLMPMA", "MP");
                support.SetPropertyValue(FQ, "IJUAFINLFQFU", "FQ");
                support.SetPropertyValue(FU, "IJUAFINLFQFU", "FU");

                //Get the t Section Data
                FINLAssemblyServices.GetCrossSectionDimensions("Euro", "HSSR", tsSize, out  tWidth, out   tFlangeThickness, out  tWebThickness, out  tDepth);
                //Get the C Section Data
                FINLAssemblyServices.GetCrossSectionDimensions("Euro", "L", lSize, out  lWidth, out   lFlangeThickness, out  lWebThickness, out  lDepth);


                if (overHang < 0)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Overhang must a positive value.", "", "SFS5397_A1.cs", 146);
                    return;
                }

                horizontalLength = structDistance + overHang - lFlangeThickness;

                if (structDistance + overHang + lFlangeThickness > maxLength)
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Maximum allowable length of " + (maxLength * 1000).ToString() + "mm exceeded.", "", "SFS5397_A1.cs", 153);

                if (overHang < pipeDiameter / 2)
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Overhang must exceed pipe outside dia:" + (pipeDiameter / 2 * 1000).ToString() + "mm.", "", "SFS5397_A1.cs", 156);

                double verticalLength = ((400.00 / 1000.00) - (1.0 / 2.0 * Math.Round((tDepth / Math.Tan(Math.PI / 4) + (tDepth / Math.Tan(Math.PI / 4))), 2)));
                double VertSteelOffset = 400.00 / 1000.00 * Math.Cos(Math.PI / 4) - horizontalLength;
                double horSteelWeight = (tWidth * tDepth * horizontalLength - ((tWidth - 2.0 * 4.0 / 1000.0) * (tDepth - 2.0 * 4.0 / 1000.0) * horizontalLength)) * STEELDENSITYKGPERM;
                string inputBOM = tsSize + " Length:" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, horizontalLength, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);
                componentDictionary[HORSTEEL5397A1].SetPropertyValue(tWidth, "IJOAHgrUtilMetricWidth", "Width");
                componentDictionary[HORSTEEL5397A1].SetPropertyValue(tDepth, "IJOAHgrUtilMetricDepth", "Depth");
                componentDictionary[HORSTEEL5397A1].SetPropertyValue(horizontalLength, "IJOAHgrUtilMetricL", "L");
                componentDictionary[HORSTEEL5397A1].SetPropertyValue(inputBOM, "IJOAHgrUtilMetricBomDesc", "InputBomDesc");
                componentDictionary[HORSTEEL5397A1].SetPropertyValue(horSteelWeight, "IJOAHgrUtilMetricInWt", "InputWeight");

                if (horizontalLength >= 500.00 / 1000.00)
                {
                    string bom = tsSize + ", Length: " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, verticalLength, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);
                    componentDictionary[VERTSTEEL5397A1].SetPropertyValue(tWidth, "IJOAHgrUtility_GENERIC_L", "WIDTH");
                    componentDictionary[VERTSTEEL5397A1].SetPropertyValue(tDepth, "IJOAHgrUtility_GENERIC_L", "DEPTH");
                    componentDictionary[VERTSTEEL5397A1].SetPropertyValue(verticalLength, "IJOAHgrUtility_GENERIC_L", "L");
                    componentDictionary[VERTSTEEL5397A1].SetPropertyValue(tFlangeThickness, "IJOAHgrUtility_GENERIC_L", "THICKNESS");
                    componentDictionary[VERTSTEEL5397A1].SetPropertyValue(Math.PI / 4, "IJOAHgrUtility_CUTBACK", "ANGLE");//rotae about y-axis
                    componentDictionary[VERTSTEEL5397A1].SetPropertyValue(Math.PI / 4, "IJOAHgrUtility_CUTBACK", "ANGLE2");
                    componentDictionary[VERTSTEEL5397A1].SetPropertyValue(bom, "IJOAHgrUtility_GENERIC_L", "BOM_DESC");
                }

                double lLENGTH = 300.00 / 1000.00, plateD = Math.Round((K + 30.00 / 1000.00 - lDepth / 2), 3);
                inputBOM = lSize + " Length:" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, horizontalLength, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);
                componentDictionary[L5397A1].SetPropertyValue(lWidth, "IJOAHgrUtility_GENERIC_L", "WIDTH");
                componentDictionary[L5397A1].SetPropertyValue(lDepth, "IJOAHgrUtility_GENERIC_L", "DEPTH");
                componentDictionary[L5397A1].SetPropertyValue(lLENGTH, "IJOAHgrUtility_GENERIC_L", "L");
                componentDictionary[L5397A1].SetPropertyValue(lFlangeThickness, "IJOAHgrUtility_GENERIC_L", "THICKNESS");
                componentDictionary[L5397A1].SetPropertyValue(inputBOM, "IJOAHgrUtility_GENERIC_L", "BOM_DESC");


                //Check the Structure is beside
                String PlateBOM = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, lFlangeThickness, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + " Steel Plate " +
                  MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, lWidth, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0) + "x" +
                  MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, plateD, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);
                //four hole plate
                componentDictionary[PLATE5397A1].SetPropertyValue(lWidth, "IJOAHgrUtilMetricWidth", "Width");
                componentDictionary[PLATE5397A1].SetPropertyValue(plateD, "IJOAHgrUtilMetricDepth", "Depth");
                componentDictionary[PLATE5397A1].SetPropertyValue(lFlangeThickness, "IJOAHgrUtilMetricT", "T");
                componentDictionary[PLATE5397A1].SetPropertyValue(PlateBOM, "IJOAHgrUtilMetricBomDesc", "InputBomDesc");

                double theta1 = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.Y, "Structure", PortAxisType.X, OrientationAlong.Global_Z);

                if (Math.Abs(theta1) < Math.PI / 2)
                {
                    JointHelper.CreateRigidJoint("-1", "Route", HORSTEEL5397A1, "StartOther", Plane.XY, Plane.XY, Axis.X, Axis.Y, -overHang, -(tDepth / 2 + pipeDiameter / 2 + shoeHeight), 0);
                    if (horizontalLength >= 500.00 / 1000.00)
                        JointHelper.CreateRigidJoint(HORSTEEL5397A1, "StartOther", VERTSTEEL5397A1, "EndStructure", Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeX, -VertSteelOffset, 0, tDepth / 2);

                    JointHelper.CreateRigidJoint(HORSTEEL5397A1, "EndOther", L5397A1, "Structure", Plane.XY, Plane.ZX, Axis.X, Axis.Z, lFlangeThickness, -lLENGTH / 2, -tDepth / 2 - lFlangeThickness);
                    JointHelper.CreateRigidJoint(HORSTEEL5397A1, "EndOther", PLATE5397A1, "TopStructure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, plateD / 2 + tDepth / 2 + (lWidth - tDepth - lFlangeThickness));
                }
                else
                {
                    JointHelper.CreateRigidJoint("-1", "Route", HORSTEEL5397A1, "StartOther", Plane.XY, Plane.XY, Axis.X, Axis.NegativeY, -overHang, tDepth / 2 + pipeDiameter / 2 + shoeHeight, 0);
                    if (horizontalLength >= 500.00 / 1000.00)
                        JointHelper.CreateRigidJoint(HORSTEEL5397A1, "StartOther", VERTSTEEL5397A1, "EndStructure", Plane.XY, Plane.NegativeXY, Axis.X, Axis.NegativeX, -VertSteelOffset, 0, tDepth / 2);

                    JointHelper.CreateRigidJoint(HORSTEEL5397A1, "EndOther", L5397A1, "Structure", Plane.XY, Plane.ZX, Axis.X, Axis.Z, lFlangeThickness, -lLENGTH / 2, -tDepth / 2 - lFlangeThickness);
                    JointHelper.CreateRigidJoint(HORSTEEL5397A1, "EndOther", PLATE5397A1, "TopStructure", Plane.XY, Plane.XY, Axis.X, Axis.X, 0, 0, plateD / 2 + tDepth / 2 + (lWidth - tDepth - lFlangeThickness));
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
                    routeConnections.Add(new ConnectionInfo(HORSTEEL5397A1, 1));
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
                    structConnections.Add(new ConnectionInfo(HORSTEEL5397A1, 1));
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
                bomString = "Console support A1 -" + sizeCodeList.PropValue + " SFS 5397";

                return bomString;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in BOM of Assembly - SFS5397_A1" + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        #endregion
    }
}

