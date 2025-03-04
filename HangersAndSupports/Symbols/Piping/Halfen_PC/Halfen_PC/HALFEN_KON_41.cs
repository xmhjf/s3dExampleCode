//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   HALFEN_KON_41.cs
//    Halfen_PC,Ingr.SP3D.Content.Support.Symbols.HALFEN_KON_41
//   Author       :  
//   Creation Date:  
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//    23-11-2012  Sasidhar    CR-CP-222275 Converted VB HS_HALFEN_PC Project to C#.Net
//    20-03-2013      Vijay    DI-CP-228142  Modify the error handling for delivered H&S symbols
//    22-02-2015      PVK      TR-CP-264951  Resolve coverity issues found in November 2014 report
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using System.Collections.ObjectModel;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;

namespace Ingr.SP3D.Content.Support.Symbols
{
    //-----------------------------------------------------------------------------------
    //Namespace of this class is Ingr.SP3D.Content.Support.Symbols
    //It is recommended that customers specify namespace of their symbols to be
    //CompanyName.SP3D.Content.Specialization.
    //It is also recommended that if customers want to change this symbol to suit their
    //requirements, they should change namespace/symbol name so the identity of the modified
    //symbol will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------

    [CacheOption(CacheOptionType.Cached)]
    [SymbolVersion("1.0.0.0")]
    [VariableOutputs]
    public class HALFEN_KON_41 : HangerComponentSymbolDefinition, ICustomHgrBOMDescription, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Halfen_PC,Ingr.SP3D.Content.Support.Symbols.HALFEN_KON_41"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "L", "L", 1)]
        public InputDouble m_dL;
        [InputDouble(3, "Finish", "Finish", 1)]
        public InputDouble m_dFinish;
        [InputDouble(4, "Bracket_W", "Bracket_W", 0.999999)]
        public InputDouble m_dBracket_W;
        [InputDouble(5, "Bracket_H", "Bracket_H", 0.999999)]
        public InputDouble m_dBracket_H;
        [InputDouble(6, "Bracket_D", "Bracket_D", 0.999999)]
        public InputDouble m_dBracket_D;
        [InputDouble(7, "Bracket_Th", "Bracket_Th", 0.999999)]
        public InputDouble m_dBracket_Th;
        [InputDouble(8, "Bracket_Bolt_CC", "Bracket_Bolt_CC", 0.999999)]
        public InputDouble m_dBracket_Bolt_CC;
        [InputDouble(9, "Bracket_Down", "Bracket_Down", 0.999999)]
        public InputDouble m_dBracket_Down;
        [InputDouble(10, "Angle_Inset", "Angle_Inset", 0.999999)]
        public InputDouble m_dAngle_Inset;
        [InputDouble(11, "Bolt_Inset_From_Top", "Bolt_Inset_From_Top", 0.999999)]
        public InputDouble m_dBolt_Inset_From_Top;
        [InputDouble(12, "Bolt_Inset_From_Bottom", "Bolt_Inset_From_Bottom", 0.999999)]
        public InputDouble m_dBolt_Inset_From_Bottom;
        [InputString(13, "Type", "Type", "No Value")]
        public InputString m_sType;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("BRACKET1", "BRACKET1")]
        [SymbolOutput("BRACKET2", "BRACKET2")]
        [SymbolOutput("BRACKET3", "BRACKET3")]
        public AspectDefinition m_Symbolic;

        #endregion

        #region "Construct Outputs"

        /// <summary>
        /// Construct symbol outputs in aspects.
        /// </summary>
        /// <remarks></remarks>
        protected override void ConstructOutputs()
        {
            try
            {
                Part part = (Part)m_PartInput.Value;
                SP3DConnection connection = default(SP3DConnection);
                connection = OccurrenceConnection;

                double L = m_dL.Value;
                Double bracketW = m_dBracket_W.Value;
                Double bracketH = m_dBracket_H.Value;
                Double bracketD = m_dBracket_D.Value;
                Double bracketTh = m_dBracket_Th.Value;
                Double bracketA = 0.0, lValue = 0.0, lM = 0.0, armL = 0.0, armW = 0.0, angleArmDown = 0.0, angleArmAngle = 0.0, angleArmL = 0.0;
                Double bracketDown = m_dBracket_Down.Value;
                Double boltInsetFromTop = m_dBolt_Inset_From_Top.Value;
                Double boltInsetFromBottom = m_dBolt_Inset_From_Bottom.Value;

                PropertyValueCodelist codeListData = (PropertyValueCodelist)part.GetPropertyValue("IJOAHgrL_List", "L");
                CodelistItem codelist = codeListData.PropertyInfo.CodeListInfo.GetCodelistItem(Convert.ToInt32(L));
                if (codelist != null)
                {
                    if (codelist.Value < 1 || codelist.Value > 5)
                    {
                        ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HalfenPCLocalizer.GetString(HalfenPClResourceIDs.ErrCodeListValue, "The CodeList value should between 1 to 5"));
                        return;
                    }
                }
                if (codelist != null)
                {
                    lValue = Convert.ToDouble(codelist.ShortDisplayName.Trim());
                }
                else
                {
                    lValue = 475;
                }
                lM = lValue / 1000;
                armW = bracketW - 2 * bracketTh;
                armL = lM - bracketTh;

                if (m_sType.Value == "2")
                {
                    if (lValue < 325)
                    {
                        ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, HalfenPCLocalizer.GetString(HalfenPClResourceIDs.ErrNGTLValue, "Minimum length of Type 2 is 325 mm."));
                    }
                    else
                    {
                        CatalogBaseHelper cataloghelper = new CatalogBaseHelper();
                        PartClass auxTable = (PartClass)cataloghelper.GetPartClass("HALFEN_KON_41_AUX");

                        ReadOnlyCollection<BusinessObject> classItems = auxTable.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                        Collection<PropertyValue> attributes = new Collection<PropertyValue>();

                        foreach (BusinessObject classItem in classItems)
                        {
                            ReadOnlyCollection<PropertyValue> properties = classItem.GetAllProperties();

                            if (classItem.GetPropertyValue("IJUAHgrHALFEN_KON_41_AUX", "L_CODE").ToString() == Math.Round(lValue).ToString())
                            {
                                bracketH = (double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrHALFEN_KON_41_AUX", "H")).PropValue;
                                angleArmL = (double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrHALFEN_KON_41_AUX", "LS")).PropValue;
                                bracketA = (double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrHALFEN_KON_41_AUX", "A")).PropValue;
                                double Z = bracketH - 2 * armW - 3 * boltInsetFromTop - 2 * boltInsetFromBottom;
                                angleArmAngle = 17;
                                angleArmDown = Math.Sin(angleArmAngle * Math.PI / 180) * angleArmL;
                                break;
                            }
                        }
                    }
                }
                else
                {
                    if (lValue > 475)
                    {
                        ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, HalfenPCLocalizer.GetString(HalfenPClResourceIDs.ErrNLTLValue, "Maximum length of Type 1 is 475 mm."));
                    }
                }
                Port port1 = new Port(connection, part, "Base", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(connection, part, "End", new Position(armL, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;
                if (bracketD <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HalfenPCLocalizer.GetString(HalfenPClResourceIDs.ErrBracktDArguments, "BracketD can not be zero or negative"));
                    return;
                }
                if (bracketW <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HalfenPCLocalizer.GetString(HalfenPClResourceIDs.ErrBracktWArguments, "BracketW can not be zero or negative"));
                    return;
                }
                if (bracketH <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HalfenPCLocalizer.GetString(HalfenPClResourceIDs.ErrBracktHArguments, "BracketH can not be zero or negative"));
                    return;
                }

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, (armW / 2 + bracketDown - bracketH) + (bracketH / 2));
                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Projection3d bracket2 = (Projection3d)symbolGeometryHelper.CreateBox(null, bracketD, bracketW, bracketH);
                m_Symbolic.Outputs["BRACKET2"] = bracket2;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(bracketTh, 0, 0);
                symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                Projection3d bracket1 = (Projection3d)symbolGeometryHelper.CreateBox(null, armL, armW, armW);
                m_Symbolic.Outputs["BRACKET1"] = bracket1;

                if (m_sType.Value != "1")
                {
                    symbolGeometryHelper = new SymbolGeometryHelper();
                    symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                    symbolGeometryHelper.SetOrientation(new Vector(1, 0, 0), new Vector(0, 1, 0));
                    Projection3d bracket3 = (Projection3d)symbolGeometryHelper.CreateBox(null, armW, angleArmL, armW);

                    Matrix4X4 matrix = new Matrix4X4();
                    matrix = new Matrix4X4();
                    matrix.SetIdentity();
                    matrix.Rotate((Math.PI * (180 - angleArmAngle) / 180), new Vector(1, 0, 0));
                    bracket3.Transform(matrix);

                    matrix = new Matrix4X4();
                    matrix.SetIdentity();
                    matrix.Rotate((Math.PI / 2), new Vector(0, 0, 1));
                    bracket3.Transform(matrix);

                    matrix = new Matrix4X4();
                    matrix.SetIdentity();
                    matrix.Translate(new Vector(angleArmL / 2 + bracketTh, -armW / 2, (angleArmL / 2) * Math.Tan(angleArmAngle * Math.PI / 180) - angleArmDown - armW));
                    bracket3.Transform(matrix);

                    m_Symbolic.Outputs["BRACKET3"] = bracket3;
                }
            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, HalfenPCLocalizer.GetString(HalfenPClResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of HALFEN_KON_41.cs."));
                    return;
                }
            }
        }
        #endregion

        #region "ICustomHgrBOMDescription Members"

        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string bomString = "";
            try
            {
                Part part = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];
                PropertyValueCodelist lengthCodelist = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrL_List", "L");
                PropertyValueCodelist finishCodelist = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrFinish_List", "Finish");

                String length = lengthCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(lengthCodelist.PropValue).DisplayName;
                String finish = finishCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(finishCodelist.PropValue).DisplayName;

                String type = (string)((PropertyValueString)part.GetPropertyValue("IJUAHgrType", "Type")).PropValue;
                String queryCode = type + "-" + length + "-" + finish;

                string stockNumber = "";

                CatalogBaseHelper cataloghelper = new CatalogBaseHelper();
                PartClass auxTable = (PartClass)cataloghelper.GetPartClass("HALFEN_KON_41_AUX");
                ReadOnlyCollection<BusinessObject> classItems = auxTable.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                Collection<PropertyValue> attributes = new Collection<PropertyValue>();
                foreach (BusinessObject classItem in classItems)
                {
                    ReadOnlyCollection<PropertyValue> properties = classItem.GetAllProperties();
                    if (classItem.GetPropertyValue("IJUAHgrHALFEN_KON_41_AUX", "Query_Code").ToString() == queryCode)
                    {
                        stockNumber = ((PropertyValueString)classItem.GetPropertyValue("IJUAHgrHALFEN_KON_41_AUX", "Stock_Number")).PropValue;
                    }
                }
                bomString = part.PartDescription + " - fv - " + length + " - " + stockNumber;

                return bomString;
            }
            catch //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, HalfenPCLocalizer.GetString(HalfenPClResourceIDs.ErrBOMDescription, "Error in BOMDescription of HALFEN_KON_41.cs."));
                }
                return "";
            }
        }
        #endregion

        #region "ICustomWeightCG Members"
        void ICustomWeightCG.EvaluateWeightCG(BusinessObject supportComponentBO)
        {
            try
            {
                Double weight, cogX = 0, cogY = 0, cogZ = 0;
                Part part = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];
                PropertyValueCodelist lengthCodelist = (PropertyValueCodelist)supportComponentBO.GetPropertyValue("IJOAHgrL_List", "L");
                PropertyValueCodelist finishCodelist = (PropertyValueCodelist)supportComponentBO.GetPropertyValue("IJOAHgrFinish_List", "Finish");
                String length = lengthCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(lengthCodelist.PropValue).DisplayName;
                String finish = finishCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(finishCodelist.PropValue).DisplayName;
                String type = (string)((PropertyValueString)part.GetPropertyValue("IJUAHgrType", "Type")).PropValue;
                String queryCode = type + "-" + length + "-" + finish;

                CatalogBaseHelper cataloghelper = new CatalogBaseHelper();
                PartClass auxTable = (PartClass)cataloghelper.GetPartClass("HALFEN_KON_41_AUX");
                ReadOnlyCollection<BusinessObject> classItems = auxTable.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                Collection<PropertyValue> attributes = new Collection<PropertyValue>();
                foreach (BusinessObject classItem in classItems)
                {
                    ReadOnlyCollection<PropertyValue> properties = classItem.GetAllProperties();
                    if (classItem.GetPropertyValue("IJUAHgrHALFEN_KON_41_AUX", "Query_Code").ToString() == queryCode)
                    {

                        weight = (double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrHALFEN_KON_41_AUX", "Weight")).PropValue;
                        SupportComponent supportComponent = (SupportComponent)supportComponentBO;
                        supportComponent.SetWeightAndCOG(weight, cogX, cogY, cogZ);
                    }
                }
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, HalfenPCLocalizer.GetString(HalfenPClResourceIDs.ErrWeightCG, "Error in WeightCG of HALFEN_KON_41.cs."));
                }
            }
        }
        #endregion

    }

}
