//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   PSL_226.cs
//    PSL,Ingr.SP3D.Content.Support.Symbols.PSL_226
//   Author       :  Vijay
//   Creation Date:  22-08-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   22-08-2013     Vijay    CR-CP-232036,232037,232038,232039,232040 Convert HS_PSL VB Project to C# .Net 
//   30-Dec-2014    PVK      TR-CP-264951	Resolve P3 coverity issues found in November 2014 report   
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

using System;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Support.Middle;

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
    [SymbolVersion("1.0.0.0")]
    public class PSL_226 : HangerComponentSymbolDefinition, ICustomHgrBOMDescription, ICustomWeightCG
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "PSL,Ingr.SP3D.Content.Support.Symbols.PSL_226"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"
        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "Length", "Length", 0.999999)]
        public InputDouble Length;
        [InputDouble(3, "ROD_DIA", "ROD_DIA", 0.999999)]
        public InputDouble ROD_DIA;
        [InputString(4, "PART_NUMBER", "PART_NUMBER", "No Value")]
        public InputString PART_NUMBER;
        [InputDouble(5, "END_THREAD_1", "END_THREAD_1", 0.999999)]
        public InputDouble END_THREAD_1;
        [InputDouble(6, "END_THREAD_2", "END_THREAD_2", 0.999999)]
        public InputDouble END_THREAD_2;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Symbolic Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("ROD", "ROD")]
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

                Double rodDiameter = ROD_DIA.Value;
                Double length = Length.Value;
                int topThreadCode = (int)END_THREAD_1.Value;
                int bottomThreadCode =(int)END_THREAD_2.Value;
                String partNumber = PART_NUMBER.Value;
                String[] arrayOfComponents = new String[8];
                Double startThreadLength = 0.0;
                Double endThreadLength = 0.0;

                //initialize array of componnents with column names in PSL_226_AUX table
                arrayOfComponents[1] = "VARIABLE_EFFORT_SUPPORT";
                arrayOfComponents[2] = "TURN_BUCKLE_NUT";
                arrayOfComponents[3] = "ROD_COUPLING_NUT";
                arrayOfComponents[4] = "CLEVIS_NUT";
                arrayOfComponents[5] = "WELDLESS_EYE_NUT";
                arrayOfComponents[6] = "WELDLESS_EYE";
                arrayOfComponents[7] = "TWO_NUTS";

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 matrix = new Matrix4X4();

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports
                if (!(topThreadCode == 8))
                    startThreadLength = PSLSymbolServices.GetDataByCondition("PSL_226_AUX", "IJUAHgrPSL_226_AUX", arrayOfComponents[(int)topThreadCode], "IJUAHgrPSL_226_AUX", "PART_NUMBER", partNumber);

                if (!(bottomThreadCode == 8))
                    endThreadLength = PSLSymbolServices.GetDataByCondition("PSL_226_AUX", "IJUAHgrPSL_226_AUX", arrayOfComponents[(int)bottomThreadCode], "IJUAHgrPSL_226_AUX", "PART_NUMBER", partNumber);

                //Warning only
                if (length < (startThreadLength + endThreadLength))
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidThreadLength226, "Overall length must exceed top thread length + bottom thread length"));

                Port port1 = new Port(OccurrenceConnection, part, "ExThdRH", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "ExThdRH", new Position(0, 0, length), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                //Validating Inputs
                if (rodDiameter <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidRodDiameterGTZ, "ROD_DIA should be greater than zero"));
                    return;
                }
                if (HgrCompareDoubleService.cmpdbl(length , 0)==true)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidLengthNEZ, "Length cannot be zero"));
                    return;
                }

                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1));
                Projection3d rodCylinder = symbolGeometryHelper.CreateCylinder(null, rodDiameter / 2.0, length);
                rodCylinder.Transform(matrix);
                m_Symbolic.Outputs["ROD"] = rodCylinder;
            }
            catch  //General Unhandled exception 
            {
                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of PSL_226."));
                    return;
                }
            }
        }

        #endregion


        #region ICustomWeightCG Members

        public void EvaluateWeightCG(BusinessObject supportComponentBO)
        {
            try
            {
                Part catalogPart = (Part)supportComponentBO.GetRelationship("madeFrom", "part").TargetObjects[0];
                double weight, cogX, cogY, cogZ;
                double length = (double)((PropertyValueDouble)supportComponentBO.GetPropertyValue("IJUAHgrOccLength_mm", "Length")).PropValue;
                double weightPerUnitlength = (double)((PropertyValueDouble)catalogPart.GetPropertyValue("IJUAHgrPSL_226", "WEIGHT_PER_LENGTH")).PropValue;

                weight = weightPerUnitlength * length;
                cogX = 0;
                cogY = 0;
                cogZ = -length / 2;

                SupportComponent supportComponent = (SupportComponent)supportComponentBO;
                supportComponent.SetWeightAndCOG(weight, cogX, cogY, cogZ);
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrWeightCG, "Error in EvaluateWeightCG of PSL_226.cs."));
            }
        }

        #endregion

        #region ICustomHgrBOMDescription Members

        public string BOMDescription(BusinessObject oSupportOrComponent)
        {
            string bomDescription = "";
            try
            {

                Part part = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];
                string finish = string.Empty;
                PropertyValueCodelist finishCodelist = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrPSL_FINISH1", "FINISH");
                if ((finishCodelist.PropValue < 1 || finishCodelist.PropValue > 6) && (finishCodelist.PropValue != 11 || finishCodelist.PropValue != 12))
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidFinish1Codelist, "FINISH codelist number should be 1 or 6 or 11 or 12"));
                    finish = "Self Colour";
                }
                else
                    finish = finishCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(finishCodelist.PropValue).ShortDisplayName;

                string partNumber = (string)((PropertyValueString)part.GetPropertyValue("IJUAHgrPSL_PART_NUMBER", "PART_NUMBER")).PropValue;
                double lengthValue = (double)((PropertyValueDouble)oSupportOrComponent.GetPropertyValue("IJUAHgrOccLength_mm", "Length")).PropValue;
                double rodDiameterValue = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrRod_Dia_mm", "ROD_DIA")).PropValue;
                string length = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, lengthValue, UnitName.DISTANCE_MILLIMETER);
                string rodDiameter = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, rodDiameterValue, UnitName.DISTANCE_MILLIMETER);

                //PSL F226M08 8mm Hanger Rod   oa:IJUAHgrOccLength_mm::Length
                bomDescription = "PSL " + partNumber + " " + rodDiameter + " Hanger Rod, Fisish:" + finish + ", Length:" + length;
                return bomDescription;
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of PSL_226.cs."));
                return "";
            }
        }

        #endregion
    }
}
