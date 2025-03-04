//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Utility_TRUNNION.cs
//    Utility,Ingr.SP3D.Content.Support.Symbols.Utility_TRUNNION
//   Author       :  Sasidhar
//   Creation Date:  04/11/2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   04/11/2012   Sasidhar   CR-CP-222288  Converted HS_Utility VB Project to C# .Net  
//	 27/03/2013		Hema 	 DI-CP-228142  Modify the error handling for delivered H&S symbols
//   04/June/2013   Manikanth TR-CP-234520 Implemented TDL Issues
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System;
using System.Linq;
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

    [CacheOption(CacheOptionType.Cached)]
    [SymbolVersion("1.0.0.0")]
    public class Utility_TRUNNION : HangerComponentSymbolDefinition, ICustomHgrBOMDescription, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Utility,Ingr.SP3D.Content.Support.Symbols.Utility_TRUNNION"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "TRUNNION_LENGTH", "TRUNNION_LENGTH", 0.999999)]
        public InputDouble m_TRUNNION_LENGTH;
        [InputDouble(3, "PIPE_DIA", "PIPE_DIA", 0.999999)]
        public InputDouble m_PIPE_DIA;
        [InputDouble(4, "DIAMETER", "DIAMETER", 0.999999)]
        public InputDouble m_DIAMETER;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("LEFT", "LEFT")]
        [SymbolOutput("RIGHT", "RIGHT")]
        public AspectDefinition m_oSymbolic;

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
                Part part =(Part)m_PartInput.Value ;

                Double trunnionLength = m_TRUNNION_LENGTH.Value;
                Double pipeDiameter = m_PIPE_DIA.Value;
                Double tempDiameter = m_DIAMETER.Value;
                Double theta, S, extraLength, diameter;

                diameter = tempDiameter;
                if (tempDiameter > pipeDiameter)
                {
                    diameter = pipeDiameter;
                }
                if (pipeDiameter <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrInvalidpipeDia, "Pipe Diameter should be greater than zero"));
                    return;
                }
                theta = Math.Asin(diameter / pipeDiameter);
                S = pipeDiameter * Math.Sin(theta / 2);
                extraLength = Math.Sqrt((S * S) - (diameter / 2 * diameter / 2));

                Port port1 = new Port(OccurrenceConnection, part, "Route", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_oSymbolic.Outputs["Port1"] = port1;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, (pipeDiameter / 2 - extraLength), 0);
                symbolGeometryHelper.SetOrientation(new Vector(0, 1, 0), new Vector(1, 0, 0));
                Projection3d left = symbolGeometryHelper.CreateCylinder(null, diameter / 2, trunnionLength + extraLength);
                m_oSymbolic.Outputs["LEFT"] = left;

                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, -(pipeDiameter / 2 - extraLength), 0);
                symbolGeometryHelper.SetOrientation(new Vector(0, -1, 0), new Vector(1, 0, 0));
                Projection3d right = symbolGeometryHelper.CreateCylinder(null, diameter / 2, trunnionLength + extraLength);
                m_oSymbolic.Outputs["RIGHT"] = right;
            }
            catch     //General Unhandled exception 
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Utility_TRUNNION"));
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
                Double diameterValue, tempDiameterValue, pipeDiameteraValue, theta, trunnionLengthValue, S, extraLengthValue;

                String size, trunnionLength;
                String[] bomUnits;

                Part part = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];

                size = (String)((PropertyValueString)part.GetPropertyValue("IJUAHgrUtility_TRUNNION", "SIZE")).PropValue;
                tempDiameterValue = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrUtility_TRUNNION", "DIAMETER")).PropValue;
                pipeDiameteraValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrPipe_Dia", "PIPE_DIA")).PropValue;
                trunnionLengthValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJOAHgrUtility_TRUNNION", "TRUNNION_LENGTH")).PropValue;

                trunnionLength = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, trunnionLengthValue, UnitName.DISTANCE_INCH);
                bomUnits = trunnionLength.Split(' ');

                diameterValue = tempDiameterValue;
                if (tempDiameterValue > pipeDiameteraValue)
                {
                    diameterValue = pipeDiameteraValue;
                }
                theta = Math.Asin(diameterValue / pipeDiameteraValue);
                S = pipeDiameteraValue * Math.Sin(theta / 2);
                extraLengthValue = Math.Sqrt((S * S) - (diameterValue / 2 * diameterValue / 2));

                trunnionLengthValue = Double.Parse(MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, trunnionLengthValue, UnitName.DISTANCE_INCH).Split(' ').First());

                extraLengthValue = Double.Parse(MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, extraLengthValue, UnitName.DISTANCE_INCH).Split(' ').First());

                bomString = "Pipe Trunnion, 2 x " + size + " Pipe, Length: " + Microsoft.VisualBasic.Conversion.Str(trunnionLengthValue + extraLengthValue) + " " + bomUnits[1];

                return bomString;

            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of Utility_TRUNNION"));
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

                Double weight, cogX, cogY, cogZ,pipediameter, trunnionLength, diameter, tempDiameter, theta, S, extraLength, trunnionThickness;
                const int getSteelDensityKGPerM = 7900;

                trunnionLength = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrUtility_TRUNNION", "TRUNNION_LENGTH")).PropValue;
                pipediameter = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJOAHgrPipe_Dia", "PIPE_DIA")).PropValue;

                Part part = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];
                tempDiameter = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrUtility_TRUNNION", "DIAMETER")).PropValue;

                diameter = tempDiameter;

                if (tempDiameter > pipediameter)
                {
                    diameter = pipediameter;
                }

                trunnionThickness = diameter / 8;
                theta = Math.Asin(diameter / pipediameter);
                S = pipediameter * Math.Sin(theta / 2);
                extraLength = Math.Sqrt((S * S) - (diameter / 2 * diameter / 2));

                weight = (((Math.PI * ((Math.Pow((diameter / 2), 2))) * (trunnionLength + extraLength)) - (Math.PI * ((Math.Pow((diameter / 2 - trunnionThickness), 2))) * (trunnionLength + extraLength))) * 2) * getSteelDensityKGPerM;
                cogX = 0;
                cogY = 0;
                cogZ = 0;
                SupportComponent supportComponent = (SupportComponent)supportComponentBO;
                supportComponent.SetWeightAndCOG(weight, cogX, cogY, cogZ);
            }
            catch
            {
                if (base.ToDoListMessage == null) //Check ToDoListMessgae created already or not
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, UtilityLocalizer.GetString(UtilitySymbolResourceIDs.ErrWeightCG, "Error in WeightCG of Utility_TRUNNION"));
                }
            }
        }

        #endregion
    }
}
