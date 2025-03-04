//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   Anvil_FIG257.cs
//    Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG257
//   Author       : Vijaya 
//   Creation Date: 1-May-2013  
//   Description: Initial Creation-CR-CP-222292

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   1-May-2013     Vijaya   CR-CP-222292 Convert HS_Anvil VB Project to C# .Net
//   11-Dec-2014     PVK     TR-CP-264951	Resolve P3 coverity issues found in November 2014 report
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
    [SymbolVersion("1.0.0.0")]
    [CacheOption(CacheOptionType.Cached)]
    public class Anvil_FIG257 : HangerComponentSymbolDefinition, ICustomHgrBOMDescription
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Anvil,Ingr.SP3D.Content.Support.Symbols.Anvil_FIG257"
        //----------------------------------------------------------------------------------
        #region "Definition of Inputs"
        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "TYP", "TYP", 1)]
        public InputDouble m_TYP;
        [InputDouble(3, "BASE_TYP", "BASE_TYP", 1)]
        public InputDouble m_BASE_TYP;
        [InputDouble(4, "BASE_TEMP", "BASE_TEMP", 0.999999)]
        public InputDouble m_dBASE_TEMP;
        [InputDouble(5, "PIPE_DIA", "PIPE_DIA", 0.999999)]
        public InputDouble m_dPIPE_DIA;
        [InputDouble(6, "FINISH", "FINISH", 1)]
        public InputDouble m_oFINISH;
        #endregion

        #region "Definitions of Aspects and their Outputs"
        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("WEB", "WEB")]
        [SymbolOutput("T_BASE", "T_BASE")]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("BASE", "BASE")]
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
                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 rotateMatrix = new Matrix4X4();

                Double pipeDiameter = m_dPIPE_DIA.Value, type = m_TYP.Value, baseType = m_BASE_TYP.Value;
                double TL = 0.3048, TW = 0.104775, thickness = 0.00635, tHeight = 0.1016, BL = 0.0, W = 0.0, H = 0.0;

                //=================================================
                //Construction of Physical Aspect 
                //=================================================

                string actualType = string.Empty, actualBaseType = string.Empty, realFigName = string.Empty;
                MetadataManager metadataManager = MiddleServiceProvider.SiteMgr.ActiveSite.ActivePlant.PlantModel.MetadataMgr;
                if (metadataManager != null)
                {
                    actualType = metadataManager.GetCodelistInfo("Anvil_Shoe_Type", "UDP").GetCodelistItem((int)type).ShortDisplayName;
                    actualBaseType = metadataManager.GetCodelistInfo("Anvil_Shoe_Base", "UDP").GetCodelistItem((int)baseType).ShortDisplayName;
                }
                else
                {
                    actualType = "Type 1";
                    actualBaseType = "Welded";
                }
                realFigName = "FIG257_BOLT";
                if (actualBaseType == "Welded")
                    realFigName = "FIG257_WELD";
                CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                PartClass auxTable = (PartClass)catalogBaseHelper.GetPartClass("Anvil_" + realFigName);
                ReadOnlyCollection<BusinessObject> classItems = auxTable.GetRelationship("PartClassContainsClassItems", "ClassItem").TargetObjects;
                foreach (BusinessObject classItem in classItems)
                {
                    if (((string)((PropertyValueString)classItem.GetPropertyValue("IJUAHgrAnvil_" + realFigName, "TYPE")).PropValue == actualType))
                    {
                        H = (double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_" + realFigName, "H")).PropValue;
                        W = (double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_" + realFigName, "W")).PropValue;
                        BL = (double)((PropertyValueDouble)classItem.GetPropertyValue("IJUAHgrAnvil_" + realFigName, "BL")).PropValue;
                        break;
                    }
                }

                //ports
                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "Structure", new Position(0, 0, -(H + pipeDiameter / 2)), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                //Validating Inputs
                if (W <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidWGTZero, "W value should be greater than zero"));
                    return;
                }
                if (BL <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidBLGTZero, "BL value should be greater than zero"));
                    return;
                }
                if (H <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidHGTZero, "H value should be greater than zero"));
                    return;
                }
                if (thickness <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidThickness, "Thickness should be greater than zero"));
                    return;
                }
                if (TL <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidTLGTZero, "TL value should be greater than zero"));
                    return;
                }
                if (TW <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidTWGTZero, "TW value should be greater than zero"));
                    return;
                }
                if (tHeight <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidT_HEIGHT, "T_HEIGHT  should be greater than zero"));
                    return;
                }

                symbolGeometryHelper.ActivePosition = new Position(-W / 2, -BL / 2, pipeDiameter / 2 + H - (H - tHeight));
                rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                rotateMatrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                Projection3d baseBox = symbolGeometryHelper.CreateBox(null, W, BL, H - tHeight, 9);
                baseBox.Transform(rotateMatrix);
                m_Symbolic.Outputs["BASE"] = baseBox;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-thickness / 2, -TL / 2, pipeDiameter / 2);
                rotateMatrix = new Matrix4X4();
                rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                rotateMatrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                Projection3d web = symbolGeometryHelper.CreateBox(null, thickness, TL, tHeight - thickness, 9);
                web.Transform(rotateMatrix);
                m_Symbolic.Outputs["WEB"] = web;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(-TW / 2, -TL / 2, pipeDiameter / 2 + tHeight - thickness);
                rotateMatrix = new Matrix4X4();
                rotateMatrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                rotateMatrix.Rotate(Math.PI, new Vector(1, 0, 0), new Position(0, 0, 0));
                Projection3d tbase = symbolGeometryHelper.CreateBox(null, TW, TL, thickness, 9);
                tbase.Transform(rotateMatrix);
                m_Symbolic.Outputs["T_BASE"] = tbase;
            }
            catch  //General Unhandled exception 
            {
                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrConstructOutputs, "Error in constructOutputs of Anvil_FIG257"));
                    return;
                }
            }
        }
        #endregion

        #region "ICustomHgrBOMDescription Members"

        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string bomDescription = "";
            try
            {
                Part catalogPart = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];
                int finish = (int)((PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrShoeFinish", "FINISH")).PropValue;
                if (finish == 0)
                    finish = 1;
                PropertyValueCodelist typeCodeList = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG257", "TYP");
                if (typeCodeList.PropValue < 1 && typeCodeList.PropValue > 6)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidTypCodelist, "TYP should be between 1 to 6"));
                    return "";
                }
                string type = typeCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(typeCodeList.PropValue).ShortDisplayName;
                PropertyValueCodelist baseTypeCodeList = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrAnvil_FIG257", "BASE_TYP");
                if (baseTypeCodeList.PropValue < 1 && baseTypeCodeList.PropValue > 2)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrInvalidbaseTypCodelist, "BASE_TYP should be between 1 to 2"));
                    return "";
                }
                string baseType = baseTypeCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(baseTypeCodeList.PropValue).ShortDisplayName;

                bomDescription = catalogPart.PartDescription + ", Base Type: " + type + ", Base Connection: " + baseType;
                return bomDescription;
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, AnvilLocalizer.GetString(AnvilSymbolResourceIDs.ErrBOMDescription, "Error in BOM description of Anvil_FIG257"));
                return "";
            }
        }
        #endregion

    }

}
