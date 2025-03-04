//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Gen_RoundPad.cs
//    Generic,Ingr.SP3D.Content.Support.Symbols.Gen_RoundPad
//   Author       :  Hema
//   Creation Date:  14-11-2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   14-11-2012      Hema   CR-CP-222274  Converted HS_Generic VB Project to C# .Net 
//	 27/03/2013		Hema 	 DI-CP-228142  Modify the error handling for delivered H&S symbols
//   6/11/2013     Vijaya   CR-CP-242533  Provide the ability to store GType outputs as a single blob for H&S .net symbols  
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using Ingr.SP3D.Support.Middle;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;

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

    [CacheOption(CacheOptionType.NonCached)] 
    [VariableOutputs]
    [SymbolVersion("1.0.0.0")]
    public class Gen_RoundPad : HangerComponentSymbolDefinition, ICustomHgrBOMDescription
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Generic,Ingr.SP3D.Content.Support.Symbols.Gen_RoundPad"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "Diameter", "Diameter", 0.999999)]
        public InputDouble m_dDiameter;
        [InputDouble(3, "Thickness", "Thickness", 0.999999)]
        public InputDouble m_dThickness;
        [InputString(4, "BOM_DESC1", "BOM_DESC1","No Value")]
        public InputString m_oBOM_DESC1;

        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Cylinder", "Cylinder")]
        [SymbolOutput("HgrPort1", "HgrPort1")]
        [SymbolOutput("HgrPort2", "HgrPort2")]
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
               
                Double diameter = m_dDiameter.Value;               
                Double thickness = m_dThickness.Value;                

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports

                Port port1 = new Port(OccurrenceConnection, part, "HgrPort_1", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["HgrPort1"] = port1;
                Port port2 = new Port(OccurrenceConnection, part, "HgrPort_2", new Position(0, 0, thickness), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["HgrPort2"] = port2;
                if (diameter <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, GenericLocalizer.GetString(GenericResourceIdentifiers.ErrInvalidDiameter, "Diameter should be greater than zero"));
                    return;
                }
                if (thickness == 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, GenericLocalizer.GetString(GenericResourceIdentifiers.ErrInvalidThickness, "Thickness cannot be zero"));
                    return;
                }
                symbolGeometryHelper.ActivePosition = new Position(0,0,0);
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(1, 0, 0));
                Projection3d cylinder = (Projection3d)symbolGeometryHelper.CreateCylinder(null, diameter / 2, thickness);
                m_Symbolic.Outputs["Cylinder"] = cylinder;
            }
            catch    //General Unhandled exception
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, GenericLocalizer.GetString(GenericResourceIdentifiers.ErrConstructOutputs, "Error in ConstructOutputs of Gen_RoundPad"));
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
                String material = "";
                Part part = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];

                PropertyValueCodelist materialCodeList = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrGenPadMat", "Material");
                CodelistItem codeList = materialCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(materialCodeList.PropValue);

                if (codeList != null)
                {
                     material = codeList.ShortDisplayName;
                }
                Double thicknessValue = (double)((PropertyValueDouble)part.GetPropertyValue("IJOAHgrGenPadThickness", "Thickness")).PropValue;
                String thickness = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, thicknessValue, UnitName.DISTANCE_MILLIMETER);

                Double diameterValue = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrGenRoundPad", "Diameter")).PropValue;
                String diameter_mm = (MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, diameterValue, UnitName.DISTANCE_MILLIMETER)).Trim();
                string[] diameter = diameter_mm.Split(' ');

                String bomDescription = (String)((PropertyValueString)part.GetPropertyValue("IJOAHgrGenPartBOMDesc", "BOM_DESC1")).PropValue;

                if( (bomDescription).ToUpper() == null)
                    bomString = "";
                else
                    bomString = part.PartDescription + ", D=" + diameter[0] + ", T=" + thickness + ", " + material;
                return bomString;
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, GenericLocalizer.GetString(GenericResourceIdentifiers.ErrBOMDescription, "Error in BOMDescription of Gen_RoundPad"));
                }
                return "";
            }
        }
        #endregion
    }
}
