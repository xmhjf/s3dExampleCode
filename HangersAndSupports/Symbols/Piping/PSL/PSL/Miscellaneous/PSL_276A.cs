//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2014R1, Intergraph Corporation. All rights reserved.
//
//   PSL_276A.cs
//    PSL,Ingr.SP3D.Content.Support.Symbols.PSL_276A
//   Author       :  Vijay
//   Creation Date:  23-08-2013
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   23-08-2013     Vijay    CR-CP-232036,232037,232038,232039,232040 Convert HS_PSL VB Project to C# .Net    
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

    [CacheOption(CacheOptionType.Cached)]
    [SymbolVersion("1.0.0.0")]
    public class PSL_276A : HangerComponentSymbolDefinition, ICustomHgrBOMDescription
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "PSL,Ingr.SP3D.Content.Support.Symbols.PSL_274"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"
        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "L", "L", 0.999999)]
        public InputDouble L;
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
        /// <remarks></remarks>
        protected override void ConstructOutputs()
        {
            try
            {
                Part part = (Part)m_PartInput.Value;
                Double length = L.Value;
                Double boltDiameter = BOLT_DIA.Value;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports
                Port port1 = new Port(OccurrenceConnection, part, "Other", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_Symbolic.Outputs["Port1"] = port1;

                //Validating Inputs
                if (boltDiameter <= 0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidBoltDiameterGTZ, "BOLT_DIA value cannot be zero"));
                    return;
                }
                 
                Vector normal = new Position(0, 0, 0).Subtract(new Position(0, 0, -length));
                symbolGeometryHelper.ActivePosition = new Position(0, 0, -length);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d boltCylinder = symbolGeometryHelper.CreateCylinder(null, boltDiameter / 2.0, normal.Length);
                m_Symbolic.Outputs["BOLT"] = boltCylinder;

                normal = new Position(0, 0, boltDiameter / 2.0).Subtract(new Position(0, 0, 0));
                symbolGeometryHelper = new SymbolGeometryHelper();
                symbolGeometryHelper.ActivePosition = new Position(0, 0, 0);
                symbolGeometryHelper.SetOrientation(normal, normal.GetOrthogonalVector());
                Projection3d headCylinder = symbolGeometryHelper.CreateCylinder(null, boltDiameter, normal.Length);
                m_Symbolic.Outputs["HEAD"] = headCylinder;
            }
            catch  //General Unhandled exception 
            {
                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of PSL_276A."));
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
                PropertyValueCodelist finishCodelist = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrPSL_FINISH6", "FINISH");
                if (finishCodelist.PropValue != 1 && finishCodelist.PropValue != 5 && finishCodelist.PropValue != 6 && finishCodelist.PropValue != 12)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrInvalidFinish6Codelist, "FINISH codelist number should be 1 or 5 or 6 or 12"));
                    finish = "Self Colour";
                }
                else
                    finish = finishCodelist.PropertyInfo.CodeListInfo.GetCodelistItem(finishCodelist.PropValue).ShortDisplayName;
                string partNumber = (string)((PropertyValueString)part.GetPropertyValue("IJUAHgrPSL_PART_NUMBER", "PART_NUMBER")).PropValue;
                double lengthValue = (double)((PropertyValueDouble)part.GetPropertyValue("IJUAHgrPSL_276A", "L")).PropValue;
                string length = MiddleServiceProvider.UOMMgr.FormatUnit(UnitType.Distance, lengthValue, UnitName.DISTANCE_MILLIMETER);

                //PSL F230M08 Eye Rod oa:  IJUAHgrOccLength_mm:: LENGTH
                bomDescription = "PSL " + partNumber + " " + " Bolt C/W Nut, Fisish:" + finish + ", Length:" + length;
                return bomDescription;
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, PSLLocalizer.GetString(PSLSymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of PSL_276A.cs."));
                return "";
            }
        }

        #endregion
    }
}
