//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   SFS5397_B3.cs
//   FINL_Assy,Ingr.SP3D.Content.Support.Rules.SFS5397_B3
//   Author       :  Vijaya
//   Creation Date:  27.Jun.2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   28-Jun-2013     Vijaya   CR-CP-224485- Convert HS_FINL_Assy to C# .Net 
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
    public class SFS5397_B3 : CustomSupportDefinition, ICustomHgrBOMDescription
    {
        private const string HORSTEEL5397B3 = "HORSTEEL5397_B3";
        private const string STEEL5397B3 = "STEEL5397_B3";
        private const string VERTSTEEL5397B3 = "VERTSTEEL5397_B3";

        private string size3;
        double pipeDiameter = 0.0, shoeHeight = 0.0, overHang = 0.0, horizontalLength, cDim, Length;

        public override Collection<PartInfo> Parts
        {
            get
            {
                try
                {
                    Ingr.SP3D.Support.Middle.Support support = SupportHelper.Support;
                    Collection<PartInfo> parts = new Collection<PartInfo>();

                    PropertyValueCodelist sizeCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJUAFINLSize3", "Size3");
                    size3 = sizeCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(sizeCodeList.PropValue).DisplayName;

                    shoeHeight = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINLShoeH", "ShoeH")).PropValue;
                    overHang = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINLOverHang", "OverHang")).PropValue;
                    cDim = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINLC_Dim", "C_Dim")).PropValue;
                    Length = (double)((PropertyValueDouble)support.GetPropertyValue("IJUAFINLL", "L")).PropValue;

                    parts.Add(new PartInfo(HORSTEEL5397B3, "Utility_GENERIC_C_1"));
                    parts.Add(new PartInfo(STEEL5397B3, "Utility_GENERIC_C_1"));

                    if (Length >= 500.00 / 1000.00)// with clamps
                        parts.Add(new PartInfo(VERTSTEEL5397B3, "Utility_CUTBACK_C2_1"));


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
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Pipe size not valid.", "", "SFS5397_B3.cs", 113);
                    return;
                }
                double maxLength = 0.0, MA = 0.0, MP = 0.0, FQ = 0.0, FU = 0.0;
                PropertyValue CSize1, CSize2;
                Dictionary<String, Object> parameter = new Dictionary<String, Object>();
                parameter.Add("Size3", size3);
                GenericHelper.GetDataByRule("FINLSrv_SFS5397_B3_Dim", "IJUAFINLSrv_SFS5397_B3_Dim", "C_Size", parameter, out CSize1);
                GenericHelper.GetDataByRule("FINLSrv_SFS5397_B3_Dim", "IJUAFINLSrv_SFS5397_B3_Dim", "Brace_Size", parameter, out CSize2);
                string cSize1 = CSize1.ToString(), cSize2 = CSize2.ToString();


                maxLength = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5397_B3_Dim", "IJUAFINLSrv_SFS5397_B3_Dim", "Max_Len", "IJUAFINLSrv_SFS5397_B3_Dim", "Size3", size3);
                MA = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5397_B3_Dim", "IJUAFINLSrv_SFS5397_B3_Dim", "MA", "IJUAFINLSrv_SFS5397_B3_Dim", "Size3", size3);
                MP = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5397_B3_Dim", "IJUAFINLSrv_SFS5397_B3_Dim", "MP", "IJUAFINLSrv_SFS5397_B3_Dim", "Size3", size3);
                FQ = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5397_B3_Dim", "IJUAFINLSrv_SFS5397_B3_Dim", "FQ", "IJUAFINLSrv_SFS5397_B3_Dim", "Size3", size3);
                FU = FINLAssemblyServices.GetDataByCondition("FINLSrv_SFS5397_B3_Dim", "IJUAFINLSrv_SFS5397_B3_Dim", "FU", "IJUAFINLSrv_SFS5397_B3_Dim", "Size3", size3);

                support.SetPropertyValue(MA, "IJUAFINLMPMA", "MA");
                support.SetPropertyValue(MP, "IJUAFINLMPMA", "MP");
                support.SetPropertyValue(FQ, "IJUAFINLFQFU", "FQ");
                support.SetPropertyValue(FU, "IJUAFINLFQFU", "FU");

                double c1Width, c1FlangeThickness, c1Depth, c1WebThickness, c2Width, c2FlangeThickness, c2Depth, c2WebThickness;

                //Get the C Section Data
                FINLAssemblyServices.GetCrossSectionDimensions("EURO-OTUA-2002", "L", cSize1, out  c1Width, out   c1FlangeThickness, out  c1WebThickness, out  c1Depth);
                FINLAssemblyServices.GetCrossSectionDimensions("EURO-OTUA-2002", "L", cSize2, out  c2Width, out   c2FlangeThickness, out  c2WebThickness, out  c2Depth);

                if (overHang < 0)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Overhang must a positive value.", "", "SFS5397_B3.cs", 144);
                    return;
                }

                horizontalLength = 680.00 / 1000.00 + cDim + cDim + overHang;

                if (overHang < pipeDiameter / 2)
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Overhang must exceed pipe outside dia:" + (pipeDiameter / 2 * 1000).ToString() + "mm.", "", "SFS5397_B3.cs", 151);

                if (cDim < 150.00 / 1000.00)
                {
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERERRORMESSAGE", "ConfigureSupport" + ": " + "ERROR: " + "Dimension C must exceed 150mm", "", "SFS5397_B3.cs", 155);
                    return;
                }

                if (Length > maxLength)
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Maximum allowable length of " + (maxLength * 1000).ToString() + "mm exceeded.", "", "SFS5397_B3.cs", 160);
                if (Length < overHang)
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Dimension L cannot be less than overhang : " + (overHang * 1000).ToString() + "mm exceeded.", "", "SFS5397_B3.cs", 162);
                if (Length < overHang + pipeDiameter / 2)
                    MiddleServiceProvider.ErrorLogger.Log(0, "", "", "USERWARNINGMESSAGE", "ConfigureSupport" + ": " + "WARNING: " + "Dimension L cannot be less than overhang + pipe radius: " + (overHang * 1000 + pipeDiameter / 2 * 1000).ToString() + "mm.", "", "SFS5397_B3.cs", 164);


                double verticalLength = ((780.00 / 1000.00 + cDim + cDim) - 1.0 / 2.0 * (c2Depth / Math.Tan(Math.PI / 6) + (c2Depth / Math.Tan(Math.PI / 2))));
                double vertSteelOffset = (780.00 / 1000.00 + cDim) * Math.Cos(Math.PI / 6) - horizontalLength;
                double cCut = c2Depth / Math.Sin(Math.PI / 6);

                string inputBOM = cSize2 + "Length:" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, horizontalLength, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);
                componentDictionary[HORSTEEL5397B3].SetPropertyValue(c1Width, "IJOAHgrUtility_Width", "Width");
                componentDictionary[HORSTEEL5397B3].SetPropertyValue(c1Depth, "IJOAHgrUtility_Depth", "Depth");
                componentDictionary[HORSTEEL5397B3].SetPropertyValue(horizontalLength, "IJOAHgrUtility_L", "L");

                componentDictionary[HORSTEEL5397B3].SetPropertyValue(c1FlangeThickness, "IJOAHgrUtility_FlangeTh", "FlangeTh");
                componentDictionary[HORSTEEL5397B3].SetPropertyValue(c1WebThickness, "IJOAHgrUtility_WebTh", "WebTh");
                componentDictionary[HORSTEEL5397B3].SetPropertyValue(inputBOM, "IJOAHgrUtility_BomDesc", "InputBomDesc");
                string bom;

                if (Length >= 500.00 / 1000.00)
                {
                    bom = cSize2 + ", Length: " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, Math.Round(verticalLength, 3), UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);

                    componentDictionary[VERTSTEEL5397B3].SetPropertyValue(c2Width, "IJOAHgrUtility_GENERIC_W", "WIDTH");
                    componentDictionary[VERTSTEEL5397B3].SetPropertyValue(c2Depth, "IJOAHgrUtility_GENERIC_W", "DEPTH");
                    componentDictionary[VERTSTEEL5397B3].SetPropertyValue(verticalLength, "IJOAHgrUtility_GENERIC_W", "L");
                    componentDictionary[VERTSTEEL5397B3].SetPropertyValue(c2WebThickness, "IJOAHgrUtility_GENERIC_W", "T_WEB");
                    componentDictionary[VERTSTEEL5397B3].SetPropertyValue(c2FlangeThickness, "IJOAHgrUtility_GENERIC_W", "T_FLANGE");
                    componentDictionary[VERTSTEEL5397B3].SetPropertyValue(-Math.PI / 6, "IJOAHgrUtility_CUTBACK", "ANGLE");//rotae about y-axis
                    componentDictionary[VERTSTEEL5397B3].SetPropertyValue(-Math.PI / 2, "IJOAHgrUtility_CUTBACK", "ANGLE2");
                    componentDictionary[VERTSTEEL5397B3].SetPropertyValue(bom, "IJOAHgrUtility_GENERIC_W", "BOM_DESC");

                }

                double additionalLength = 280.00 / 1000.00;
                bom = cSize1 + ", Length: " + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, additionalLength, UnitName.DISTANCE_MILLIMETER).Split('.').GetValue(0);
                componentDictionary[STEEL5397B3].SetPropertyValue(c1Width, "IJOAHgrUtility_Width", "Width");
                componentDictionary[STEEL5397B3].SetPropertyValue(c1Depth, "IJOAHgrUtility_Depth", "Depth");
                componentDictionary[STEEL5397B3].SetPropertyValue(additionalLength, "IJOAHgrUtility_L", "L");
                componentDictionary[STEEL5397B3].SetPropertyValue(c1WebThickness, "IJOAHgrUtility_FlangeTh", "FlangeTh");
                componentDictionary[STEEL5397B3].SetPropertyValue(c1FlangeThickness, "IJOAHgrUtility_WebTh", "WebTh");
                componentDictionary[STEEL5397B3].SetPropertyValue(bom, "IJOAHgrUtility_BomDesc", "InputBomDesc");

                double theta1 = RefPortHelper.AngleBetweenPorts("Route", PortAxisType.Y, "Structure", PortAxisType.X, OrientationAlong.Global_Z);

                if (Math.Abs(theta1) < Math.PI / 2)
                {
                    if (Configuration == 1)
                        JointHelper.CreateRigidJoint("-1", "Route", HORSTEEL5397B3, "BeginCap", Plane.XY, Plane.YZ, Axis.X, Axis.Z, -overHang, c1Depth / 2 - (c1Depth / 2 + pipeDiameter / 2 + shoeHeight), c1Width / 2);
                    else
                        JointHelper.CreateRigidJoint("-1", "Route", HORSTEEL5397B3, "BeginCap", Plane.XY, Plane.NegativeYZ, Axis.X, Axis.NegativeZ, overHang, c1Depth / 2 - (c1Depth / 2 + pipeDiameter / 2 + shoeHeight), c1Width / 2);

                    if (Length >= 500.00 / 1000.00)
                        JointHelper.CreateRigidJoint(HORSTEEL5397B3, "BeginCap", VERTSTEEL5397B3, "StartStructure", Plane.XY, Plane.ZX, Axis.X, Axis.NegativeZ, c1Width / 2, c1Depth, cCut + overHang);

                    JointHelper.CreateRigidJoint(HORSTEEL5397B3, "BeginCap", STEEL5397B3, "EndCap", Plane.XY, Plane.YZ, Axis.X, Axis.NegativeZ, 0, 0, overHang + Length);
                }
                else
                {
                    if (Configuration == 1)
                        JointHelper.CreateRigidJoint("-1", "Route", HORSTEEL5397B3, "BeginCap", Plane.XY, Plane.NegativeYZ, Axis.X, Axis.Z, overHang, -(c1Depth / 2 - (c1Depth / 2 + pipeDiameter / 2 + shoeHeight)), c1Width / 2);
                    else
                        JointHelper.CreateRigidJoint("-1", "Route", HORSTEEL5397B3, "BeginCap", Plane.XY, Plane.YZ, Axis.X, Axis.NegativeZ, -overHang, -(c1Depth / 2 - (c1Depth / 2 + pipeDiameter / 2 + shoeHeight)), c1Width / 2);

                    if (Length >= 500.00 / 1000.00)
                        JointHelper.CreateRigidJoint(HORSTEEL5397B3, "BeginCap", VERTSTEEL5397B3, "StartStructure", Plane.XY, Plane.ZX, Axis.X, Axis.NegativeZ, c1Width / 2, c1Depth, cCut + overHang);

                    JointHelper.CreateRigidJoint(HORSTEEL5397B3, "BeginCap", STEEL5397B3, "EndCap", Plane.XY, Plane.YZ, Axis.X, Axis.NegativeZ, 0, 0, Length);
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
                    routeConnections.Add(new ConnectionInfo(HORSTEEL5397B3, 1));
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
                    structConnections.Add(new ConnectionInfo(HORSTEEL5397B3, 1));
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
                PropertyValueCodelist sizeCodeList = (PropertyValueCodelist)support.GetPropertyValue("IJUAFINLSize3", "Size3");
                bomString = "Console support B3 -" + sizeCodeList.PropValue + " SFS 5397";

                return bomString;
            }
            catch (Exception e)
            {
                Type myType = this.GetType();
                CmnException e1 = new CmnException("Error in BOM of Assembly - SFS5397_B3" + myType.Assembly.FullName + "," + myType.Namespace + "." + myType.Name + ". Error:" + e.Message, e);
                throw e1;
            }
        }
        #endregion
    }
}

