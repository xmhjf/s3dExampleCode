//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   PSL_228.cs
//    PSL,Ingr.SP3D.Content.Support.Symbols.PSL_228
//   Author       :  Vijay
//   Creation Date:  22-08-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   22-08-2013     Vijay    CR-CP-232036,232037,232038,232039,232040 Convert HS_PSL VB Project to C# .Net    
//                           WeightCG implementation is commented ,as in VB "WEIGHT_PER_LENGTH" property is not available in refdata.
//   11-Jun-2015    PVK      TR-CP-239160	Issues observed in .Net PSL Parts
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
    public class PSL_228 : HangerComponentSymbolDefinition, ICustomHgrBOMDescription
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "PSL,Ingr.SP3D.Content.Support.Symbols.PSL_228"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"
        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "E", "E", 0.999999)]
        public InputDouble E;
        [InputDouble(3, "ROD_DIA", "ROD_DIA", 0.999999)]
        public InputDouble ROD_DIA;
        [InputDouble(4, "L", "L", 0.999999)]
        public InputDouble L;
        [InputDouble(5, "MAX_THREAD", "MAX_THREAD", 0.999999)]
        public InputDouble MAX_THREAD;
        [InputDouble(6, "THREAD_L", "THREAD_L", 0.999999)]
        public InputDouble THREAD_L;
        #endregion

        #region "Definitions of Aspects and their Outputs"

        [Aspect("Symbolic", "Symbolic Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("CYL", "CYL")]
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
                Double length = L.Value;
                Double e = E.Value;
                Double maximumThread = MAX_THREAD.Value;
                Double threadLength = THREAD_L.Value;

                if (maximumThread < threadLength)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidThreadL, "MTHREAD_L cannot exceed MAX_THREAD length" + Convert.ToSingle(maximumThread * 1000)));
                    threadLength = maximumThread;
                }

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                Matrix4X4 matrix = new Matrix4X4();

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports

                Port port1 = new Port(OccurrenceConnection, part, "Eye", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                Port port2 = new Port(OccurrenceConnection, part, "ExThdRH", new Position(0, 0, length + e / 2.0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port2"] = port2;

                //Validating Inputs
                if (rodDiameter <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidRodDiameterGTZ, "ROD_DIA should be greater than zero"));
                    return;
                }

                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                matrix.Rotate(Math.PI / 2, new Vector(1, 0, 0));
                matrix.Translate(new Vector(0, rodDiameter / 2.0, e / 2.0));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d cylinder = symbolGeometryHelper.CreateCylinder(null, e / 2.0 + rodDiameter, rodDiameter);
                cylinder.Transform(matrix);
                m_Symbolic.Outputs["CYL"] = cylinder;

                matrix = new Matrix4X4();
                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.SetOrientation(new Vector(0, 0, 1), new Vector(0, 0, 1).GetOrthogonalVector());
                matrix.Translate(new Vector(0, 0, e));
                matrix.Rotate(3 * Math.PI / 2, new Vector(0, 0, 1), new Position(0, 0, 0));
                Projection3d rodCylinder = symbolGeometryHelper.CreateCylinder(null, rodDiameter / 2.0, length - e / 2.0);
                rodCylinder.Transform(matrix);
                m_Symbolic.Outputs["ROD"] = rodCylinder;
            }
            catch  //General Unhandled exception 
            {
                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of PSL_228."));
                    return;
                }
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

                string partNumber = (string)((PropertyValueString)part.GetPropertyValue("IJUAHgrPSL_228", "PART_NO_R")).PropValue;
                double lengthValue = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrPSL_228", "L")).PropValue;
                double rodDiameterValue = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrRod_Dia_mm", "ROD_DIA")).PropValue;
                string length = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, lengthValue, UnitName.DISTANCE_MILLIMETER);
                string rodDiameter = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, rodDiameterValue, UnitName.DISTANCE_MILLIMETER);

                //PSL F228M 10 mm Forged Eyebolt
                bomDescription = "PSL " + partNumber + " " + rodDiameter + " Forged Eyebolt, Fisish:" + finish + ", Length:" + length;
                return bomDescription;
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of PSL_228.cs."));
                return "";
            }
        }

        #endregion
    }
}
