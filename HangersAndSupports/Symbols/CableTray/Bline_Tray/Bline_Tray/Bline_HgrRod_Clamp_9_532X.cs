//+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
//   Copyright (c) 2012, Intergraph Corporation. All rights reserved.
//
//   Bline_HgrRod_Clamp_9_532X.cs
//   Bline_Tray,Ingr.SP3D.Content.Support.Symbols.Bline_HgrRod_Clamp_9_532X
//   Author       :  Shilpi	
//   Creation Date:  10.August.2012
//   Description:

//   Change History:
//   dd.mmm.yyyy     who     change description
//   -----------     ---     ------------------
//   10.August.2012  Shilpi     CR-CP-219113 Converted HS_B-Line VB project to C#.Net Project 
//   23.Nov.2012      Hema      CR-CP-219113 Modified the code with SymbolGeomHelper
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
    //Namespace of this class is Ingr.SP3D.Content.Support.Symbols
    //It is recommended that customers specify namespace of their symbols to be
    //<CompanyName>.SP3D.Content.<Specialization>.
    //It is also recommended that if customers want to change this symbol to suit their
    //requirements, they should change namespace/symbol name so the identity of the modified
    //symbol will be different from the one delivered by Intergraph.
    //-----------------------------------------------------------------------------------
    [CacheOption(CacheOptionType.Cached)]
    [SymbolVersion("1.0.0.0")]
    [VariableOutputs]
    public class Bline_HgrRod_Clamp_9_532X : HangerComponentSymbolDefinition, ICustomHgrBOMDescription
    {
        //----------------------------------------------------------------------------------
        //DefinitionName/ProgID of this symbol is "Bline_Tray,Ingr.SP3D.Content.Support.Symbols.Bline_HgrRod_Clamp_9_532X"
        //----------------------------------------------------------------------------------

        #region "Definition of Inputs"

        [InputCatalogPart(1)]
        public InputCatalogPart m_PartInput;
        [InputDouble(2, "LC", "Length of Clamp", 0)]
        public InputDouble m_dLC;
        [InputDouble(3, "WC", "Width of Clamp", 0)]
        public InputDouble m_dWC;
        [InputDouble(4, "TH", "Thickness of Clamp", 0)]
        public InputDouble m_dTH;
        [InputDouble(5, "LS", "Distance from edge to hole", 0)]
        public InputDouble m_dLS;
        [InputDouble(6, "RH", "Tray Rail Height", 0)]
        public InputDouble m_dRH;
        [InputDouble(7, "TrayWT", "Tray Web Thickness", 0)]
        public InputDouble m_dTrayWT;
        [InputDouble(8, "WT", "Tray Width", 0)]
        public InputDouble m_dWT;

        #endregion

        #region "Definitions of Aspects and their outputs"

        [Aspect("Symbolic", "Simple Physical Aspect", AspectID.SimplePhysical)]
        [SymbolOutput("L_CLAMP", "L_CLAMP")]
        [SymbolOutput("R_CLAMP", "R_CLAMP")]
        [SymbolOutput("Port1", "Port1")]
        [SymbolOutput("Port2", "Port2")]
        [SymbolOutput("Port3", "Port3")]

        public AspectDefinition m_PhysicalAspect;

        #endregion

        #region "Construction of outputs of all aspects"
        protected override void ConstructOutputs()
        {
            try
            {
                Part part = (Part)m_PartInput.Value;

                double LC = m_dLC.Value;
                double WC = m_dWC.Value;
                double TH = m_dTH.Value;
                double LS = m_dLS.Value;
                double RH = m_dRH.Value;
                double TW = m_dTrayWT.Value;
                double WT = m_dWT.Value;

                //=================================================
                //Construction of Physical Aspect 
                //=================================================
                if (WC == 0)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrInvalidWCgtZero, "Width of Clamp value should not be equal to 0."));

                //ports
                Port trayPort = new Port(OccurrenceConnection, part, "TrayPort", new Position(0, 0, TH), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["Port1"] = trayPort;

                Port inThrdRH1 = new Port(OccurrenceConnection, part, "InThrdRH1", new Position(0, (WT / 2 + TW + LC / 2 - LS), -0.008), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["Port2"] = inThrdRH1;

                Port inThrdRH2 = new Port(OccurrenceConnection, part, "InThrdRH2", new Position(0, -(WT / 2 + TW + LC / 2 - LS), -0.008), new Vector(1, 0, 0), new Vector(0, 0, 1));
                m_PhysicalAspect.Outputs["Port3"] = inThrdRH2;

                Collection<Position> pointCollection = new Collection<Position>();
                Vector projectionVector = new Vector();

                //Add Extruded Polygon for L_CLAMP

                pointCollection.Add(new Position(-WC / 2, WT / 2 + TW - LC / 2, 0));
                pointCollection.Add(new Position(-WC / 2, WT / 2 + TW + LC / 2, 0));
                pointCollection.Add(new Position(-WC / 2, WT / 2 + TW + LC / 2, RH + 2 * TH));
                pointCollection.Add(new Position(-WC / 2, WT / 2 + TW - LC / 2, RH + 2 * TH));
                pointCollection.Add(new Position(-WC / 2, WT / 2 + TW - LC / 2, RH + TH));
                pointCollection.Add(new Position(-WC / 2, WT / 2 + TW + LC / 2 - TH, RH + TH));
                pointCollection.Add(new Position(-WC / 2, WT / 2 + TW + LC / 2 - TH, TH));
                pointCollection.Add(new Position(-WC / 2, WT / 2 + TW - LC / 2, TH));
                pointCollection.Add(new Position(-WC / 2, WT / 2 + TW - LC / 2, 0));

                projectionVector.Set(WC, 0, 0);
                Projection3d LClamp = new Projection3d(new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);

                m_PhysicalAspect.Outputs["L_CLAMP"] = LClamp;
                pointCollection.Clear();

                //Add Extruded Polygon for R_CLAMP

                pointCollection.Add(new Position(-WC / 2, -(WT / 2 + TW - LC / 2), 0));
                pointCollection.Add(new Position(-WC / 2, -(WT / 2 + TW + LC / 2), 0));
                pointCollection.Add(new Position(-WC / 2, -(WT / 2 + TW + LC / 2), RH + 2 * TH));
                pointCollection.Add(new Position(-WC / 2, -(WT / 2 + TW - LC / 2), RH + 2 * TH));
                pointCollection.Add(new Position(-WC / 2, -(WT / 2 + TW - LC / 2), RH + TH));
                pointCollection.Add(new Position(-WC / 2, -(WT / 2 + TW + LC / 2 - TH), RH + TH));
                pointCollection.Add(new Position(-WC / 2, -(WT / 2 + TW + LC / 2 - TH), TH));
                pointCollection.Add(new Position(-WC / 2, -(WT / 2 + TW - LC / 2), TH));
                pointCollection.Add(new Position(-WC / 2, -(WT / 2 + TW - LC / 2), 0));

                projectionVector.Set(WC, 0, 0);
                Projection3d RClamp = new Projection3d(new LineString3d(pointCollection), projectionVector, projectionVector.Length, true);
                m_PhysicalAspect.Outputs["R_CLAMP"] = RClamp;
            }

            catch //General Unhandled exception 
            {

                if (base.ToDoListMessage == null)
                {
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrConstructOutputs, "Error in ConstructOutputs of Bline_HgrRod_Clamp_9_532X."));
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
                Part catalogPart = (Part)oSupportOrComponent.GetRelationship("madeFrom", "part").TargetObjects[0];

                PropertyValueCodelist materialCodeListValue = (PropertyValueCodelist)oSupportOrComponent.GetPropertyValue("IJOAHangerRodClampMat", "Material");
                if (materialCodeListValue.PropValue < 0 || materialCodeListValue.PropValue > 8)
                    ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrClampMaterial, "Clamp Material Code list value should be between 1 and 2"));
                string material = materialCodeListValue.PropertyInfo.CodeListInfo.GetCodelistItem(materialCodeListValue.PropValue).DisplayName;

                bomDescription = catalogPart.PartDescription + ",Material: " + material;
                return bomDescription;
            }
            catch
            {
                ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageWarning, Bline_TrayLocalizer.GetString(Bline_TraySymbolResourceIDs.ErrBOMDescription, "Error in BOMDescription of Bline_HgrRod_Clamp_9_532X."));
                return "";

            }
        }
        #endregion
    }
}



