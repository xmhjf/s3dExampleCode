//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2013, Intergraph Corporation. All rights reserved.
//
//   PSL_276B.cs
//    PSL,Ingr.SP3D.Content.Support.Symbols.PSL_276B
//   Author       :  Manikanth
//   Creation Date:  21-08-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   21-08-2013    Manikanth CR-CP-232036,232037,232038,232039,232040 Convert HS_PSL VB Project to C# .Net
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

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
    [SymbolVersion("1.0.0.0")]
    [CacheOption(CacheOptionType.Cached)]
    public class PSL_276B : HangerComponentSymbolDefinition, ICustomHgrBOMDescription
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "PSL,Ingr.SP3D.Content.Support.Symbols.PSL_276B"
        //----------------------------------------------------------------------------------
        #region "Definition of Inputs"
        [InputCatalogPart(1)]
        public InputCatalogPart PartInput;
        [InputDouble(2, "L", "L", 0.999999)]
        public InputDouble L1;
        [InputDouble(3, "BOLT_DIA", "BOLT_DIA", 0.999999)]
        public InputDouble BOLT_DIA;
        #endregion

        #region "Definitions of Aspects and their Outputs"
        [Aspect("Symbolic", "Symbolic Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("BOLT", "BOLT")]
        [SymbolOutput("HEAD", "HEAD")]
        [SymbolOutput("Port1", "Port1")]
        public AspectDefinition m_Symbolic;
        #endregion

        #region "Construct Outputs"
        /// <summary>
        /// Construct symbol outputs in aspects.
        /// </summary>
        protected override void ConstructOutputs()
        {
            try
            {
                Part part = (Part)PartInput.Value;
                double L = L1.Value;
                double boltDiameter = BOLT_DIA.Value;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();

                if (boltDiameter <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidBoltDiameter, "BOLT_DIA should be greater than zero"));
                    return;
                }
                symbolGeometryHelper = new SymbolGeometryHelper();
                Vector normal = new Position(0, 0, 0).Subtract(new Position(0, 0, -L));
                symbolGeometryHelper.ActivePosition = new Position(0, 0, -L);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d boltCylinder = symbolGeometryHelper.CreateCylinder(null, boltDiameter / 2, normal.Length);
                m_Symbolic.Outputs["BOLT"] = boltCylinder;

                symbolGeometryHelper = new SymbolGeometryHelper();
                Vector normal1 = new Position(0, 0, boltDiameter / 2).Subtract(new Position(0, 0, 0));
                symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                symbolGeometryHelper.SetOrientation(normal1, normal1.GetOrthogonalVector());
                Projection3d headCylinder = symbolGeometryHelper.CreateCylinder(null, boltDiameter, normal1.Length);
                m_Symbolic.Outputs["HEAD"] = headCylinder;

                Port port1 = new Port(OccurrenceConnection, part, "Other", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;
            }
            catch  //General Unhandled exception 
            {
                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of PSL_276B"));
                    return;
                }
            }
        }
        #endregion

        #region "ICustomHgrBOMDescription Members"
        public string BOMDescription(BusinessObject SupportOrComponent)
        {
            string bomDescrition = "";
            try
            {
                Part part = (Part)SupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];
                string finish = string.Empty;
                PropertyValueCodelist finishCodelist = ((PropertyValueCodelist)SupportOrComponent.GetPropertyValue("IJOAHgrPSL_FINISH6", "FINISH"));
                if (finishCodelist.PropValue != 1 && finishCodelist.PropValue != 5 && finishCodelist.PropValue != 6 && finishCodelist.PropValue != 12)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidFinish6Codelist, "FINISH codelist number should be 1 or 5 or 6 or 12"));
                    finish = "Self Colour";
                }
                else
                    finish = finishCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(finishCodelist.PropValue).DisplayName;
                string partNumber = (string)((PropertyValueString)part.GetPropertyValue("IJUAHgrPSL_PART_NUMBER", "PART_NUMBER")).PropValue;
                double length = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrPSL_276B", "L")).PropValue;
                bomDescrition = "PSL " + partNumber + " " + " Bolt C/W Nut, Finish:" + finish + ", Length:" + MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, length, UnitName.DISTANCE_INCH_SYMBOL);
                return bomDescrition;
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of PSL_276B."));
                return "";
            }
        }
        #endregion
    }
}
