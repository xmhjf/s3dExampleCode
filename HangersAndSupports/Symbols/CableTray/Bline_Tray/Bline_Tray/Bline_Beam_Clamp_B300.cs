//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Bline_Beam_Clamp_B300.cs
//   Bline_Tray,Ingr.SP3D.Content.Support.Symbols.Bline_Beam_Clamp_B300
//   Author       :  Hema
//   Creation Date:  31.July.2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   31.July.2012    Hema       CR-CP-219113 Converted HS_B-Line VB project to C#.Net Project 
//   23.Nov.2012     Rajeswari  CR-CP-219113 Modified the code with SymbolGeomHelper
//   26.Mar.2013     Vijaya     DI-CP-228142  Modify the error handling for delivered H&S symbols 
//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
using System.Collections.ObjectModel;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.Support.Middle;


namespace Ingr.SP3D.Content.Support.Symbols
{
    //-----------------------------------------------------------------------------------
    //Namespace of this class is Ingr.SP3D.Support.Content.Symbols
    //It is recommended that customers specify namespace of their symbols to be
    //<CompanyName>.SP3D.Content.<Specialization>.
    //It is also recommended that if customers want to change this symbol to suit their
    //requirements, they should change namespace/symbol name so the identity of the modified
    //symbol will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------

    [CacheOption(CacheOptionType.Cached)]
    [SymbolVersion("1.0.0.0")]
    [VariableOutputs]
    public class Bline_Beam_Clamp_B300 : HangerComponentSymbolDefinition, ICustomHgrBOMDescription
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Bline_Tray,Ingr.SP3D.Content.Support.Symbols.Bline_Beam_Clamp_B300"
        //----------------------------------------------------------------------------------
        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "C", "Height of Clamp", 0.0)]
        public InputDouble m_dC;
        [InputDouble(3, "F", "Length of Clamp", 0.0)]
        public InputDouble m_dF;
        [InputDouble(4, "WC", "Width of Clamp", 0.0)]
        public InputDouble m_dWC;
        [InputDouble(5, "HO", "Height of Opening (D)", 0.0)]
        public InputDouble m_dHO;
        [InputDouble(6, "E", "Width of Opening", 0.0)]
        public InputDouble m_dE;
        [InputDouble(7, "B", "Hole Diameter", 0.0)]
        public InputDouble m_B;

        #endregion

        #region "Definitions of Aspects and their outputs"

        // Physical Aspect 
        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("CLAMP_BODY", "CLAMP_BODY")]
        [SymbolOutput("SETSCREW", "SETSCREW")]
        public AspectDefinition m_PhysicalAspect;

        #endregion

        #region "Construction of outputs of all aspects"
        protected override void ConstructOutputs()
        {
            try
            {
                Part part = (Part)m_PartInput.Value;
                double C = m_dC.Value;
                double F = m_dF.Value;
                double WC = m_dWC.Value;
                double D = m_dHO.Value;
                double E = m_dE.Value;
                double B = m_B.Value;

                SymbolGeometryHelper symbolGeometryHelper = new SymbolGeometryHelper();
                if (B <= 0.0)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidB, "B value should not be less than or equal to 0"));
                    return;
                }


                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                //ports
                Port steel = new Port(OccurrenceConnection, part, "Steel", new Position(0, 0, 0), new Vector(1, 0, 0), new Vector(0, 0, -1));
                m_PhysicalAspect.Outputs["Port1"] = steel;

                Port inThdRH1 = new Port(OccurrenceConnection, part, "InThrdRH1", new Position(0, E / 2, -C / 2 + 0.008), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["Port2"] = inThdRH1;

                Collection<Position> pointCollection = new Collection<Position>();
                pointCollection.Add(new Position(-WC / 2, E, C / 2));
                pointCollection.Add(new Position(-WC / 2, E - F, C / 2));
                pointCollection.Add(new Position(-WC / 2, E - F, -C / 2));
                pointCollection.Add(new Position(-WC / 2, E, -C / 2));
                pointCollection.Add(new Position(-WC / 2, E, -D / 2));
                pointCollection.Add(new Position(-WC / 2, 0, -D / 2));
                pointCollection.Add(new Position(-WC / 2, 0, D / 2));
                pointCollection.Add(new Position(-WC / 2, E, D / 2));
                pointCollection.Add(new Position(-WC / 2, E, C / 2));

                Vector projectionVecctor = new Vector(WC, 0, 0);
                Projection3d clampBody = new Projection3d(OccurrenceConnection, new LineString3d(pointCollection), projectionVecctor, projectionVecctor.Length, true);
                m_PhysicalAspect.Outputs["CLAMP_BODY"] = clampBody;

                Vector normal = new Position(0, E / 2, C / 2 + B).Subtract(new Position(0, E / 2, D / 2));
                symbolGeometryHelper.ActivePosition = new Position(0, E / 2, D / 2);
                symbolGeometryHelper.SetOrientation(normal, new Vector(1, 0, 0));
                Projection3d setScrew = symbolGeometryHelper.CreateCylinder(null, B / 2, normal.Length);
                m_PhysicalAspect.Outputs["SETSCREW"] = setScrew;
            }
            catch
            {
                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Bline_Beam_Clamp_B300."));
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
                PropertyValueCodelist finishCodeList = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHgrBlineBmClampFinish", "Finish");

                if (finishCodeList.PropValue <= 0 || finishCodeList.PropValue > 2)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrBeamClampFinish, "Beam Clamp Finish Code list value should be between 1 and 2"));


                string finish = finishCodeList.PropertyInfo.CodeListInfo.GetCodelistItem(finishCodeList.PropValue).DisplayName;
                Part beamClampPart = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];

                bomDescription = beamClampPart.PartDescription + ", Finish: " + finish;
                return bomDescription;
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of Bline_Beam_Clamp_B300."));
                return "";
            }
        }

        #endregion
    }
}


