//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SFS5397_A3.cs
//   FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5397_A3
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
    public class SFS5397_A3 : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        private const string HORSTEEL5397A3 = "HORSTEEL5397_A3";
        private const string L5397A3 = "L5397_A3";
        private const string CSECTION5397A3 = "CSECTION5397_A3";
        private const string VERTSTEEL5397A3 = "VERTSTEEL5397_A3";
        private string size2, size;
        double pipeDiameter = 0.0, shoeHeight = 0.0, overHang = 0.0, horizontalLength;

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

                    PropertyValue CSize;
                    Dictionary<String, Object> parameter = new Dictionary<String, Object>();
                    parameter.Add("Size2", size2);
                    GenericHelper.GetDataByRule("FINLSrv_SFS5397_A3_Dim", "IJUAFINLSrv_SFS5397_A3_Dim", "C_Size", parameter, out CSize);
                    size = CSize.ToString();


                    double cWidth, cFlangeThickness, cWebThickness, cDepth;
                    //Get the C Section Data
                    FINLAssemblyServices.GetCrossSectionDimensions("EURO-OTUA-2002", "C", size, out  cWidth, out   cFlangeThickness, out  cWebThickness, out  cDepth);

                    horizontalLength = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Horizontal) + overHang - cWidth;

                    parts.Add(new PartInfo(HORSTEEL5397A3, "Utility_GENERIC_C_1"));
                    parts.Add(new PartInfo(L5397A3, "Utility_GENERIC_L_1"));
                    parts.Add(new PartInfo(CSECTION5397A3, "Utility_GENERIC_C_1"));

                    if (horizontalLength >= 500.00 / 1000.00)// with clamps                    
                        parts.Add(new PartInfo(VERTSTEEL5397A3, "Utility_CUTBACK_C1_1"));

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
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Pipe size not valid.", "", "SFS5397_A3.cs", 1);
                    return;
                }
                double maxLength = 0.0, K = 0.0, MA = 0.0, MP = 0.0, FQ = 0.0, FU = 0.0;
                PropertyValue CSize, LSize;
                Dictionary<String, Object> parameter = new Dictionary<String, Object>();
                parameter.Add("Size2", size2);
                GenericHelper.GetDataByRule("FINLSrv_SFS5397_A3_Dim", "IJUAFINLSrv_SFS5397_A3_Dim", "C_Size", parameter, out CSize);
                GenericHelper.GetDataByRule("FINLSrv_SFS5397_A3_Dim", "IJUAFINLSrv_SFS5397_A3_Dim", "L_Size", parameter, out LSize);
                string cSize = CSize.ToString(), lSize = LSize.ToString();


                maxLength = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5397_A3_Dim", "IJUAFINLSrv_SFS5397_A3_Dim", "Max_Len", "IJUAFINLSrv_SFS5397_A3_Dim", "Size2", size2);
                K = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5397_A3_Dim", "IJUAFINLSrv_SFS5397_A3_Dim", "K", "IJUAFINLSrv_SFS5397_A3_Dim", "Size2", size2);
                MA = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5397_A3_Dim", "IJUAFINLSrv_SFS5397_A3_Dim", "MA", "IJUAFINLSrv_SFS5397_A3_Dim", "Size2", size2);
                MP = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5397_A3_Dim", "IJUAFINLSrv_SFS5397_A3_Dim", "MP", "IJUAFINLSrv_SFS5397_A3_Dim", "Size2", size2);
                FQ = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5397_A3_Dim", "IJUAFINLSrv_SFS5397_A3_Dim", "FQ", "IJUAFINLSrv_SFS5397_A3_Dim", "Size2", size2);
                FU = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5397_A3_Dim", "IJUAFINLSrv_SFS5397_A3_Dim", "FU", "IJUAFINLSrv_SFS5397_A3_Dim", "Size2", size2);


                support.SetPropertyValue(MA, "IJUAFINLMPMA", "MA");
                support.SetPropertyValue(MP, "IJUAFINLMPMA", "MP");
                support.SetPropertyValue(FQ, "IJUAFINLFQFU", "FQ");
                support.SetPropertyValue(FU, "IJUAFINLFQFU", "FU");

                double cWidth, cFlangeThickness, cWebThickness, cDepth, lWidth, lFlangeThickness, lWebThickness, lDepth;
                //Get the C Section Data
                FINLAssemblyServices.GetCrossSectionDimensions("EURO-OTUA-2002", "C", cSize, out  cWidth, out   cFlangeThickness, out  cWebThickness, out  cDepth);
                FINLAssemblyServices.GetCrossSectionDimensions("Euro", "L", lSize, out  lWidth, out   lFlangeThickness, out  lWebThickness, out  lDepth);


                if (overHang < 0)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Overhang must a positive value.", "", "SFS5397_A3.cs", 157);
                    return;
                }
                double structDistance = RefPortHelper.DistanceBetweenPorts("Route", "Structure", PortDistanceType.Horizontal);
                horizontalLength = structDistance + overHang - cWidth;

                if (structDistance + overHang + lFlangeThickness > maxLength)
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Maximum allowable length of " + (maxLength * 1000).ToString() + "mm exceeded.", "", "SFS5397_A3.cs", 164);

                if (overHang < pipeDiameter / 2)
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Overhang must exceed pipe outside dia:" + (pipeDiameter / 2 * 1000).ToString() + "mm.", "", "SFS5397_A3.cs", 167);

                double verticalLength = ((400.00 / 1000.00) - (1.0 / 2.0 * Math.Round((cWidth / Math.Tan(Math.PI / 4) + (cWidth / Math.Tan(Math.PI / 4))), 3)));
                double cCut = cWidth / Math.Sin(Math.PI / 4);


                string inputBOM = cSize + " Length:" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, horizontalLength, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);
                componentDictionary[HORSTEEL5397A3].SetPropertyValue(cWidth, "IJOAHgrUtility_Width", "Width");
                componentDictionary[HORSTEEL5397A3].SetPropertyValue(cDepth, "IJOAHgrUtility_Depth", "Depth");
                componentDictionary[HORSTEEL5397A3].SetPropertyValue(horizontalLength, "IJOAHgrUtility_L", "L");

                componentDictionary[HORSTEEL5397A3].SetPropertyValue(cFlangeThickness, "IJOAHgrUtility_FlangeTh", "FlangeTh");
                componentDictionary[HORSTEEL5397A3].SetPropertyValue(cWebThickness, "IJOAHgrUtility_WebTh", "WebTh");
                componentDictionary[HORSTEEL5397A3].SetPropertyValue(inputBOM, "IJOAHgrUtility_BomDesc", "InputBomDesc");

                if (horizontalLength >= 500.00 / 1000.00)
                {
                    string bom = cSize + ", Length: " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, verticalLength, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);

                    componentDictionary[VERTSTEEL5397A3].SetPropertyValue(cWidth, "IJOAHgrUtility_GENERIC_W", "WIDTH");
                    componentDictionary[VERTSTEEL5397A3].SetPropertyValue(cDepth, "IJOAHgrUtility_GENERIC_W", "DEPTH");
                    componentDictionary[VERTSTEEL5397A3].SetPropertyValue(verticalLength, "IJOAHgrUtility_GENERIC_W", "L");
                    componentDictionary[VERTSTEEL5397A3].SetPropertyValue(cFlangeThickness, "IJOAHgrUtility_GENERIC_W", "T_FLANGE");
                    componentDictionary[VERTSTEEL5397A3].SetPropertyValue(cWebThickness, "IJOAHgrUtility_GENERIC_W", "T_WEB");
                    componentDictionary[VERTSTEEL5397A3].SetPropertyValue(-Math.PI / 4, "IJOAHgrUtility_CUTBACK", "ANGLE");//rotae about y-axis
                    componentDictionary[VERTSTEEL5397A3].SetPropertyValue(-Math.PI / 4, "IJOAHgrUtility_CUTBACK", "ANGLE2");
                    componentDictionary[VERTSTEEL5397A3].SetPropertyValue(bom, "IJOAHgrUtility_GENERIC_W", "BOM_DESC");

                }

                double lLength = 300.00 / 1000.00;
                inputBOM = lSize + " Length:" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, horizontalLength, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);
                componentDictionary[L5397A3].SetPropertyValue(lWidth, "IJOAHgrUtility_GENERIC_L", "WIDTH");
                componentDictionary[L5397A3].SetPropertyValue(lDepth, "IJOAHgrUtility_GENERIC_L", "DEPTH");
                componentDictionary[L5397A3].SetPropertyValue(lLength, "IJOAHgrUtility_GENERIC_L", "L");
                componentDictionary[L5397A3].SetPropertyValue(lFlangeThickness, "IJOAHgrUtility_GENERIC_L", "THICKNESS");
                componentDictionary[L5397A3].SetPropertyValue(inputBOM, "IJOAHgrUtility_GENERIC_L", "BOM_DESC");

                //additional C - section
                Double lC = Math.Round((K + 50.00 / 1000.00 - lDepth / 2), 4);
                String plateBOM = cSize + ", Length: " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, Math.Round((K + 50.00 / 1000.00 - lDepth / 2), 4), UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);

                //four hole plate
                componentDictionary[CSECTION5397A3].SetPropertyValue(cWidth, "IJOAHgrUtility_Width", "Width");
                componentDictionary[CSECTION5397A3].SetPropertyValue(cDepth, "IJOAHgrUtility_Depth", "Depth");
                componentDictionary[CSECTION5397A3].SetPropertyValue(lC, "IJOAHgrUtility_L", "L");
                componentDictionary[CSECTION5397A3].SetPropertyValue(cFlangeThickness, "IJOAHgrUtility_FlangeTh", "FlangeTh");
                componentDictionary[CSECTION5397A3].SetPropertyValue(cWebThickness, "IJOAHgrUtility_WebTh", "WebTh");
                componentDictionary[CSECTION5397A3].SetPropertyValue(plateBOM, "IJOAHgrUtility_BomDesc", "InputBomDesc");

                double theta1 = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.Y, "Structure", PortAxisType.X, OrientationAlong.Global_Z);

                if (Math.Abs(theta1) < Math.PI / 2)
                {
                    JointHelper.CreateRigidJoint("-1", "Route", HORSTEEL5397A3, "BeginCap", Plane.XY, Plane.YZ, Axis.X, Axis.NegativeY, -overHang, cDepth / 2 - (cDepth / 2 + pipeDiameter / 2 + shoeHeight), 0);
                    if (horizontalLength >= 500.00 / 1000.00)
                        JointHelper.CreateRigidJoint(HORSTEEL5397A3, "EndCap", VERTSTEEL5397A3, "EndStructure", Plane.XY, Plane.YZ, Axis.X, Axis.NegativeZ, cWidth, cDepth / 2, cCut - 400.00 / 1000.00 * Math.Cos(Math.PI / 4));

                    JointHelper.CreateRigidJoint(HORSTEEL5397A3, "EndCap", CSECTION5397A3, "BeginCap", Plane.XY, Plane.YZ, Axis.X, Axis.NegativeZ, 0, 0, cWidth);
                    JointHelper.CreateRigidJoint(HORSTEEL5397A3, "EndCap", L5397A3, "Structure", Plane.XY, Plane.ZX, Axis.X, Axis.NegativeZ, 0, cDepth / 2 + lLength / 2, cWidth);
                }
                else
                {
                    JointHelper.CreateRigidJoint("-1", "Route", HORSTEEL5397A3, "BeginCap", Plane.XY, Plane.YZ, Axis.X, Axis.Y, -overHang, -cDepth / 2 + (cDepth / 2 + pipeDiameter / 2 + shoeHeight), 0);
                    if (horizontalLength >= 500.00 / 1000.00)
                        JointHelper.CreateRigidJoint(HORSTEEL5397A3, "EndCap", VERTSTEEL5397A3, "EndStructure", Plane.XY, Plane.YZ, Axis.X, Axis.NegativeZ, cWidth, cDepth / 2, cCut - 400.00 / 1000.00 * Math.Cos(Math.PI / 4));

                    JointHelper.CreateRigidJoint(HORSTEEL5397A3, "EndCap", CSECTION5397A3, "BeginCap", Plane.XY, Plane.YZ, Axis.X, Axis.NegativeZ, 0, 0, cWidth);
                    JointHelper.CreateRigidJoint(HORSTEEL5397A3, "EndCap", L5397A3, "Structure", Plane.XY, Plane.ZX, Axis.X, Axis.NegativeZ, 0, cDepth / 2 + lLength / 2, cWidth);
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
                    routeConnections.Add(new ConnectionInfo(HORSTEEL5397A3, 1));
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
                    structConnections.Add(new ConnectionInfo(HORSTEEL5397A3, 1));
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
                bomString = "Console support A3 -" + sizeCodeList.PropValue + " SFS 5397";

                return bomString;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in BOM of Assembly - SFS5397_A3" + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        #endregion
    }
}

