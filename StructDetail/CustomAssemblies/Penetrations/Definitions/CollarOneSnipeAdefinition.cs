//**************************************************************************************************************************/
//Copyright 1998 - 2016 Intergraph Corporation, Intergraph Process, Power & Marine. All rights reserved.
//
//File
//  CollarOneSnipeADefinition.cs
//
//  Original Dll Name: <InstallDir>\SharedContent\Bin\StructDetail\SmartOccurrence\Release\SMCollarRules.dll
//  Original Class Name: ‘CollarOneSnipeAdef’ in VB content
//
//Abstract
// CollarOneSnipeADefinition is a .NET custom assembly definition which creates the PhysicalConnections, free EdgeTreatments and corner Features needed on a Sniped Collar.
// This class subclasses from CollarCustomAssemblyDefinition.
//
// Change History:
//  dd.mmm.yyyy    who    change description
//****************************************************************************************************************************/
using System;
using System.Collections.Generic;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Content.Structure.Services;
using Ingr.SP3D.Structure.Middle;

namespace Ingr.SP3D.Content.Structure
{
    /// <summary>
    /// Implementation of collar one snipe A type .NET custom assembly definition class.
    /// CollarOneSnipeADefinition is a .NET custom assembly definition which creates PhysicalConnections, EdgeTreatment and corner Features if required.
    /// </summary>
    [SymbolVersion("1.0.0.0")]
    [OutputNotification(MarineSymbolConstants.IID_IJPlate, PenetrationsResourceIds.Thickness)]
    [OutputNotification(MarineSymbolConstants.IID_IJStructureMaterial, PenetrationsResourceIds.MatAndGrade)]
    [OutputNotification(MarineSymbolConstants.IID_IJCollarPart, PenetrationsResourceIds.SideOfPlate)]
    [DefaultLocalizer(PenetrationsResourceIds.PenetrationsResource, PenetrationsResourceIds.PenetrationsAssembly)]
    public class CollarOneSnipeADefinition : CollarCustomAssemblyDefinition
    {
        //============================================================================================================
        //DefinitionName/ProgID of this symbol is "Penetrations,Ingr.SP3D.Content.Structure.CollarOneSnipeADefinition"
        //============================================================================================================

        #region Private members
        //AssemblyOutput names
        private const string BaseRightPhysicalConnection = "BaseRightPC";
        private const string WebRightPhysicalConnection = "WebRightPC";
        private const string WebRightTopCornerPhysicalConnection = "WebRightTopCornerPC";
        private const string TopFlangeRightBottomPhysicalConnection = "TopFlangeRightBottomPC";
        private const string TopFlangeRightBottomCornerPhysicalConnection = "TopFlangeRightBottomCornerPC";
        private const string TopFlangeRightPhysicalConnection = "TopFlangeRightPC";
        private const string NormalSideLapPhysicalConnection = "NormalSideLapPC";
        private const string OppositeSideLapPhysicalConnection = "OppositeSideLapPC";
        private const string SnipeCornerFeature = "CornerSnipe";
        private const string TopRightPhysicalConnection = "TopRightPC";
        private const string TopRightCornerPhysicalConnection = "TopRightCornerPC";
        private const string RightPhysicalConnection = "RightPC";
        private const string BottomFlangeRightBottomFreeEdgeTreatment = "BottomFlangeRightBottomFET";
        private const string BottomFlangeRightBottomCornerFreeEdgeTreatment = "BottomFlangeRightBottomCornerFET";
        private const string TopFlangeRightTopCornerPhysicalConnection = "TopFlangeRightTopCornerPC";
        private const string TopPhysicalConnection = "TopPC";
        private const string NormalSideAdditonalLapPhysicalConnection = "NormalSideLapPC17";
        private const string OppositeSideAdditonalLapPhysicalConnection = "OppositeSideLapPC18";
        private const string BaseRightAdditonalPhysicalConnection = "BaseRight2PC";
        #endregion Private members

        #region Definitions of assembly outputs

        /// <summary>
        /// Tee weld PhysicalConnection between CollarPart bottom right edge and Base Plate
        /// </summary>
        [AssemblyOutput(1, BaseRightPhysicalConnection)]
        public AssemblyOutput baseRightPhysicalConnectionAssemblyOutput;

        /// <summary>
        /// Tee weld PhysicalConnection with the penetrating web right.
        /// </summary>        
        [AssemblyOutput(2, WebRightPhysicalConnection)]
        public AssemblyOutput webRightPhysicalConnectionAssemblyOutput;

        /// <summary>
        /// Tee weld PhysicalConnection with the penetrating web right top corner.
        /// </summary>
        [AssemblyOutput(3, WebRightTopCornerPhysicalConnection)]
        public AssemblyOutput webRightTopCornerPhysicalConnectionAssemblyOutput;

        /// <summary>
        /// Tee weld PhysicalConnection with the penetrating top flange right bottom.
        /// </summary>
        [AssemblyOutput(4, TopFlangeRightBottomPhysicalConnection)]
        public AssemblyOutput topFlangeRightBottomPhysicalConnectionAssemblyOutput;

        /// <summary>
        /// Tee weld PhysicalConnection with the penetrating top flange right bottom corner.
        /// </summary>
        [AssemblyOutput(5, TopFlangeRightBottomCornerPhysicalConnection)]
        public AssemblyOutput topFlangeRightBottomCornerPhysicalConnectionAssemblyOutput;

        /// <summary>
        /// Tee weld PhysicalConnection with the penetrating top flange right.
        /// </summary>
        [AssemblyOutput(6, TopFlangeRightPhysicalConnection)]
        public AssemblyOutput topFlangeRightPhysicalConnectionAssemblyOutput;

        /// <summary>
        /// Lap weld PhysicalConnection with the penetrated base port.
        /// </summary>
        [AssemblyOutput(7, NormalSideLapPhysicalConnection)]
        public AssemblyOutput normalSideLapPhysicalConnectionAssemblyOutput;

        /// <summary>
        /// Lap weld PhysicalConnection with the penetrated offset port.
        /// </summary>
        [AssemblyOutput(8, OppositeSideLapPhysicalConnection)]
        public AssemblyOutput oppositeSideLapPhysicalConnectionAssemblyOutput;

        /// <summary>
        /// Snipe corner Feature with bottom web port.
        /// </summary>
        [AssemblyOutput(9, SnipeCornerFeature)]
        public AssemblyOutput snipeCornerFeatureAssemblyOutput;

        /// <summary>
        /// Top right PhysicalConnection between CollarPart and penetrated plate.
        /// </summary>
        [AssemblyOutput(10, TopRightPhysicalConnection)]
        public AssemblyOutput topRightPhysicalConnectionAssemblyOutput;

        /// <summary>
        /// Top right corner PhysicalConnection between CollarPart and penetrated plate.
        /// </summary>
        [AssemblyOutput(11, TopRightCornerPhysicalConnection)]
        public AssemblyOutput topRightCornerPhysicalConnectionAssemblyOutput;

        /// <summary>
        /// Right PhysicalConnection between CollarPart and penetrated plate.
        /// </summary>
        [AssemblyOutput(12, RightPhysicalConnection)]
        public AssemblyOutput rightPhysicalConnectionAssemblyOutput;

        /// <summary>
        /// Free EdgeTreatment at bottom flange right bottom.
        /// </summary>
        [AssemblyOutput(13, BottomFlangeRightBottomFreeEdgeTreatment)]
        public AssemblyOutput bottomFlangeRightBottomEdgeTreatmentAssemblyOutput;

        /// <summary>
        /// Free EdgeTreatment at bottom flange right bottom corner.
        /// </summary>
        [AssemblyOutput(14, BottomFlangeRightBottomCornerFreeEdgeTreatment)]
        public AssemblyOutput bottomFlangeRightBottomCornerEdgeTreatmentAssemblyOutput;

        /// <summary>
        /// Top flange right top corner PhysicalConnection between CollarPart and profile for CollarBCT_A3
        /// </summary>
        [AssemblyOutput(15, TopFlangeRightTopCornerPhysicalConnection)]
        public AssemblyOutput topFlangeRightTopCornerPhysicalConnectionAssemblyOutput;

        /// <summary>
        /// Top PhysicalConnection for CollarBCT_A3, CollarTCT_A3 CollarParts
        /// </summary>
        [AssemblyOutput(16, TopPhysicalConnection)]
        public AssemblyOutput topPhysicalConnectionAssemblyOutput;

        /// <summary>
        /// Additional lap weld PhysicalConnection with the penetrated base port.
        /// </summary>
        [AssemblyOutput(17, NormalSideAdditonalLapPhysicalConnection)]
        public AssemblyOutput normalSideLapAdditonalPhysicalConnectionAssemblyOutput;

        /// <summary>
        /// Additional lap weld PhysicalConnection with the penetrated offset port.
        /// </summary>
        [AssemblyOutput(18, OppositeSideAdditonalLapPhysicalConnection)]
        public AssemblyOutput oppositeSideLapAdditonalPhysicalConnectionAssemblyOutput;

        /// <summary>
        /// Additional base PhysicalConnection between CollarPart bottom right edge and base plate when CollarPart crosses Seam.
        /// </summary>
        [AssemblyOutput(19, BaseRightAdditonalPhysicalConnection)]
        public AssemblyOutput baseRightAdditonalPhysicalConnectionAssemblyOutput;

        #endregion Definitions of assembly outputs

        #region Public override properties and methods

        /// <summary>
        /// Construct and re-evaluate the custom assembly outputs.
        /// Here we will do the following:
        /// 1. Set the physical properties (Thickness, Material and SideOfPart) on the CollarPart based on the penetrated object.
        /// 2. Decide which assembly outputs are needed.
        /// 3. Create the ones which are needed, delete which are not needed now.        
        /// </summary>
        public override void EvaluateAssembly()
        {
            try
            {
                //Validating the inputs required to create the CollarPart.
                base.ValidateInputs();

                //Check if any ToDo message of type error has been created during the validation of inputs
                if (base.ToDoListMessage != null && base.ToDoListMessage.Type == ToDoMessageTypes.ToDoMessageError)
                {
                    //ToDo list is created while validating the inputs
                    return;
                }

                //It's only required to be explicitly called if the CollarPart is not constructed using its 3D API constructor.
                base.AddCollarGeometry();               

                //get the required assembly outputs for this definition           
                Dictionary<string, bool> requiredAssemblyOutputs = this.RequiredAssemblyOutputs;

                //Create the Tee weld PhysicalConnection between CollarPart bottom right face port and CollarPart's BasePlatePort, if not needed then delete the assembly output.                
                base.CreateOrDeleteBasePhysicalConnection(this.baseRightPhysicalConnectionAssemblyOutput, (int)SectionFaceType.Bottom_Right, DetailingCustomAssembliesConstants.TeeWeld, requiredAssemblyOutputs[BaseRightPhysicalConnection]);

                //Create the Tee weld PhysicalConnection between CollarPart's web right face port and penetrating object's web right face port, if not needed then delete the assembly output.
                base.CreateOrDeletePhysicalConnectionBetweenCollarAndPenetrating(this.webRightPhysicalConnectionAssemblyOutput, (int)SectionFaceType.Web_Right, (int)SectionFaceType.Web_Right, DetailingCustomAssembliesConstants.TeeWeld, requiredAssemblyOutputs[WebRightPhysicalConnection]);

                //Create the Tee weld PhysicalConnection between CollarPart's web right top corner face port and penetrating object's web right top corner face port, if not needed then delete the assembly output.
                base.CreateOrDeletePhysicalConnectionBetweenCollarAndPenetrating(this.webRightTopCornerPhysicalConnectionAssemblyOutput, (int)SectionFaceType.Web_Right_Top_Corner, (int)SectionFaceType.Web_Right_Top_Corner, DetailingCustomAssembliesConstants.TeeWeld, requiredAssemblyOutputs[WebRightTopCornerPhysicalConnection]);

                //Create the Tee weld PhysicalConnection between CollarPart's top flange right bottom face port and penetrating object's top flange right bottom face port, if not needed then delete the assembly output.                
                base.CreateOrDeletePhysicalConnectionBetweenCollarAndPenetrating(this.topFlangeRightBottomPhysicalConnectionAssemblyOutput, (int)SectionFaceType.Top_Flange_Right_Bottom, (int)SectionFaceType.Top_Flange_Right_Bottom, DetailingCustomAssembliesConstants.TeeWeld, requiredAssemblyOutputs[TopFlangeRightBottomPhysicalConnection]);

                //Create the Tee weld PhysicalConnection between CollarPart's top flange right bottom corner face port and penetrating object's top flange right bottom corner face port, if not needed then delete the assembly output.
                base.CreateOrDeletePhysicalConnectionBetweenCollarAndPenetrating(this.topFlangeRightBottomCornerPhysicalConnectionAssemblyOutput, (int)SectionFaceType.Top_Flange_Right_Bottom_Corner, (int)SectionFaceType.Top_Flange_Right_Bottom_Corner, DetailingCustomAssembliesConstants.TeeWeld, requiredAssemblyOutputs[TopFlangeRightBottomCornerPhysicalConnection]);

                //Create the Tee weld PhysicalConnection between CollarPart's top flange right face port and penetrating object's top flange right face port, if not needed then delete the assembly output.
                base.CreateOrDeletePhysicalConnectionBetweenCollarAndPenetrating(this.topFlangeRightPhysicalConnectionAssemblyOutput, (int)SectionFaceType.Top_Flange_Right, (int)SectionFaceType.Top_Flange_Right, DetailingCustomAssembliesConstants.TeeWeld, requiredAssemblyOutputs[TopFlangeRightPhysicalConnection]);

                //Create the Lap weld PhysicalConnection with the penetrated base port, if not needed then delete the assembly output.
                //Since the CollarPart is expanded in the direction same direction as Base to Offset ports from the plate, the port used for the lap weld on the CollarPart is always the Offset
                base.CreateOrDeletePhysicalConnectionBetweenPenetratedAndCollar(this.normalSideLapPhysicalConnectionAssemblyOutput, ContextTypes.Base, ContextTypes.Offset, DetailingCustomAssembliesConstants.LapWeld, requiredAssemblyOutputs[NormalSideLapPhysicalConnection]);

                //Create the Lap weld PhysicalConnection with the penetrated offset port, if not needed then delete the assembly output.
                //Since the CollarPart is expanded in the direction opposite direction as Offset to Base ports from the plate, the port used for the lap weld on the CollarPart is always the Base
                base.CreateOrDeletePhysicalConnectionBetweenPenetratedAndCollar(this.oppositeSideLapPhysicalConnectionAssemblyOutput, ContextTypes.Offset, ContextTypes.Base, DetailingCustomAssembliesConstants.LapWeld, requiredAssemblyOutputs[OppositeSideLapPhysicalConnection]);

                //Create the snipe corner Feature with bottom web port, if not needed then delete the assembly output.
                base.CreateOrDeleteCornerFeature(this.snipeCornerFeatureAssemblyOutput, ContextTypes.Base, (int)SectionFaceType.Web_Right, (int)SectionFaceType.Bottom_Right, DetailingCustomAssembliesConstants.SnipeOnCollar, requiredAssemblyOutputs[SnipeCornerFeature]);

                //Create the Butt weld PhysicalConnection with the penetrated top right, if not needed then delete the assembly output.
                base.CreateOrDeletePhysicalConnectionBetweenPenetratedAndCollar(this.topRightPhysicalConnectionAssemblyOutput, (int)SectionFaceType.Top_Flange_Right_Top, (int)SectionFaceType.Top, DetailingCustomAssembliesConstants.ButtWeld, requiredAssemblyOutputs[TopRightPhysicalConnection]);

                //Create the Butt weld PhysicalConnection with the penetrated top right corner, if not needed then delete the assembly output.
                base.CreateOrDeletePhysicalConnectionBetweenPenetratedAndCollar(this.topRightCornerPhysicalConnectionAssemblyOutput, (int)SectionFaceType.Top_Right_Corner, (int)SectionFaceType.Top_Flange_Right_Top_Corner, DetailingCustomAssembliesConstants.ButtWeld, requiredAssemblyOutputs[TopRightCornerPhysicalConnection]);

                //Create the Butt weld PhysicalConnection with the penetrated right, if not needed then delete the assembly output.
                base.CreateOrDeletePhysicalConnectionBetweenPenetratedAndCollar(this.rightPhysicalConnectionAssemblyOutput, (int)SectionFaceType.Right, (int)SectionFaceType.Right, DetailingCustomAssembliesConstants.ButtWeld, requiredAssemblyOutputs[RightPhysicalConnection]);

                //Create free EdgeTreatment at bottom flange right bottom, if not needed then delete the assembly output.
                base.CreateOrDeleteFreeEdgeTreatment(this.bottomFlangeRightBottomEdgeTreatmentAssemblyOutput, (int)SectionFaceType.Bottom_Flange_Right_Bottom, DetailingCustomAssembliesConstants.Bevel, requiredAssemblyOutputs[BottomFlangeRightBottomFreeEdgeTreatment]);

                //Create free EdgeTreatment at bottom flange right bottom corner, if not needed then delete the assembly output.
                base.CreateOrDeleteFreeEdgeTreatment(this.bottomFlangeRightBottomCornerEdgeTreatmentAssemblyOutput, (int)SectionFaceType.Bottom_Flange_Right_Bottom_Corner, DetailingCustomAssembliesConstants.Bevel, requiredAssemblyOutputs[BottomFlangeRightBottomCornerFreeEdgeTreatment]);

                //Create Tee weld PhysicalConnection with the penetrating top flange right top corner, if not needed then delete the assembly output.
                base.CreateOrDeletePhysicalConnectionBetweenCollarAndPenetrating(this.topFlangeRightTopCornerPhysicalConnectionAssemblyOutput, (int)SectionFaceType.Top_Flange_Right_Top_Corner, (int)SectionFaceType.Top_Flange_Right_Top_Corner, DetailingCustomAssembliesConstants.TeeWeld, requiredAssemblyOutputs[TopFlangeRightTopCornerPhysicalConnection]);

                //Create Tee weld PhysicalConnection with the penetrating top, if not needed then delete the assembly output.
                base.CreateOrDeletePhysicalConnectionBetweenCollarAndPenetrating(this.topPhysicalConnectionAssemblyOutput, (int)SectionFaceType.Top_Flange_Top, (int)SectionFaceType.Top, DetailingCustomAssembliesConstants.TeeWeld, requiredAssemblyOutputs[TopPhysicalConnection]);

                //Create additional Lap weld PhysicalConnection with the penetrated base port, if not needed then delete the assembly output.
                //Tolerance used in getting the overlapping penetrated plate parts to create the PhysicalConnection.
                double tolerance = 0.01;
                base.CreateOrDeleteAdditionalPenetratedPhysicalConnection(this.normalSideLapAdditonalPhysicalConnectionAssemblyOutput, tolerance, ContextTypes.Base, ContextTypes.Offset, DetailingCustomAssembliesConstants.LapWeld, requiredAssemblyOutputs[NormalSideAdditonalLapPhysicalConnection]);

                //Create additional Lap weld PhysicalConnection with the penetrated offset port, if not needed then delete the assembly output.
                base.CreateOrDeleteAdditionalPenetratedPhysicalConnection(this.oppositeSideLapAdditonalPhysicalConnectionAssemblyOutput, tolerance, ContextTypes.Offset, ContextTypes.Base, DetailingCustomAssembliesConstants.LapWeld, requiredAssemblyOutputs[OppositeSideAdditonalLapPhysicalConnection]);

                //Create additional Tee weld PhysicalConnection between CollarPart bottom right edge and base plate, if not needed then delete the assembly output.
                //Tolerance used in getting the overlapping base plate parts to create the PhysicalConnection.
                tolerance = 0.04;
                base.CreateOrDeleteAdditionalBasePhysicalConnection(this.baseRightAdditonalPhysicalConnectionAssemblyOutput, tolerance, (int)SectionFaceType.Bottom_Right, DetailingCustomAssembliesConstants.TeeWeld, requiredAssemblyOutputs[BaseRightAdditonalPhysicalConnection]);
            }
            catch (Exception)
            {
                //may be there could be specific ToDo record is created before, so do not override it.
                if (base.ToDoListMessage == null)
                {
                    base.ToDoListMessage = new ToDoListMessage(ToDoMessageTypes.ToDoMessageError, string.Format(base.GetString(PenetrationsResourceIds.ToDoDefinition,
                            "Unexpected error while evaluating the custom assembly of {0}"), this.ToString()));
                }
            }
        }

         /// <summary>
        /// This method is called prior to the construction of the symbol outputs. This method can be overriden to perform any actions to be done before construction of the symbol outputs.
        /// </summary>
        public override void PreConstructOutputs()
        {            
            //Note: Unlike VB content, with .Net content definition rule is triggered after collar USS.
            //Collar thickness if not set when its plate thickness is set by collar USS.
            //This method PreConstructOutputs helps in executing the methods with in this, to be executed at the right time, inline with VB content.

            //Sets the physical properties (Thickness, material and SideOfPart) on the collar part based on the penetrated object.
            base.SetCollarThickness();
            
            base.SetCollarMaterial();
            
            base.SetCollarSideOfPart();
        }

        #endregion Public override properties and methods

        #region Private Methods

        /// <summary>
        /// Gets all needed assembly outputs for the current configuration of the CollarPart.
        /// </summary>
        /// <returns>This returns the needed outputs of the CollarPart.</returns>
        private Dictionary<string, bool> RequiredAssemblyOutputs
        {
            get
            {
                //Dictionary which holds the data of the assembly output name and the Boolean which indicates
                //whether the corresponding assembly output is needed or not.
                Dictionary<string, bool> requiredAssemblyOutputs = new Dictionary<string, bool>();

                bool isPartialDetailed = false;
                Ingr.SP3D.Interop.StructDetailUtil.StructDetailHelper oHlpr = new Ingr.SP3D.Interop.StructDetailUtil.StructDetailHelper();
                isPartialDetailed = oHlpr.IsPartialDetailed(Ingr.SP3D.Common.Middle.Services.Hidden.COMConverters.ConvertBOToCOMBO( (Ingr.SP3D.Common.Middle.BusinessObject)base.Penetrated)) || oHlpr.IsPartialDetailed(Ingr.SP3D.Common.Middle.Services.Hidden.COMConverters.ConvertBOToCOMBO( (Ingr.SP3D.Common.Middle.BusinessObject)base.Penetrating));
                oHlpr = null;

                requiredAssemblyOutputs.Add(BaseRightPhysicalConnection, ! isPartialDetailed);
                requiredAssemblyOutputs.Add(WebRightPhysicalConnection, ! isPartialDetailed);

                //Get penetrating section type and CollarPart name
                string penetratingSectionTypeName = base.PenetratingSectionTypeName;

                string collarPartName = ((CollarPart)base.Occurrence).PartName;

                PlatePartBase penetratingPlate = base.Penetrating as PlatePartBase;

                PlatePartBase penetratedPlate = base.Penetrated as PlatePartBase;
                StiffenerPartBase penetratedStiffener = base.Penetrated as StiffenerPartBase;

                bool isWebRightTopCornerPhysicalConnectionNeeded = false;

                bool isTopFlangeRightBottomPhysicalConnectionNeeded = false;

                bool isTopFlangeRightBottomCornerPhysicalConnectionNeeded = false;

                bool isTopFlangeRightPhysicalConnectionNeeded = false;

                bool isNormalSideLapPhysicalConnectionNeeded = false;

                bool isOppositeSideLapPhysicalConnectionNeeded = false;

                bool isSnipeCornerFeatureNeeded = false;

                bool isTopRightPhysicalConnectionNeeded = false;

                bool isTopRightCornerPhysicalConnectionNeeded = false;

                bool isRightPhysicalConnectionNeeded = false;

                bool isBottomFlangeRightBottomFreeEdgeTreatmentNeeded = false;

                bool isBottomFlangeRightBottomCornerFreeEdgeTreatmentNeeded = false;

                bool isTopFlangeRightTopCornerPhysicalConnectionNeeded = false;

                bool isTopPhysicalConnectionNeeded = false;

                bool isNormalSideAdditonalLapPhysicalConnectionNeeded = false;

                bool isOppositeSideAdditonalLapPhysicalConnectionNeeded = false;

                bool isBaseRightAdditonalPhysicalConnectionNeeded = false;


                if (!isPartialDetailed)
                {
                    isWebRightTopCornerPhysicalConnectionNeeded = PenetrationsServices.IsCornerPhysicalConnectionNeeded(penetratingSectionTypeName, penetratingPlate);

                    isTopFlangeRightBottomPhysicalConnectionNeeded = PenetrationsServices.IsTopFlangeRightBottomPhysicalConnectionNeeded(penetratingSectionTypeName);

                    isTopFlangeRightBottomCornerPhysicalConnectionNeeded = PenetrationsServices.IsCornerPhysicalConnectionNeeded(penetratingSectionTypeName, penetratingPlate);

                    isTopFlangeRightPhysicalConnectionNeeded = PenetrationsServices.IsTopFlangeRightPhysicalConnectionNeeded(penetratingSectionTypeName);

                    isNormalSideLapPhysicalConnectionNeeded = PenetrationsServices.IsClipSideLapPhysicalConnectionNeeded(NormalSideLapPhysicalConnection, collarPartName, (CollarPart)base.Occurrence);

                    isOppositeSideLapPhysicalConnectionNeeded = PenetrationsServices.IsClipSideLapPhysicalConnectionNeeded(OppositeSideLapPhysicalConnection, collarPartName, (CollarPart)base.Occurrence);
                }

                //Corner Feature
               isSnipeCornerFeatureNeeded = PenetrationsServices.IsCornerSnipeNeeded((CollarPart)base.Occurrence);

               if (!isPartialDetailed)
               {
                   isTopRightPhysicalConnectionNeeded = PenetrationsServices.IsPhysicalConnectionBetweenCollarAndPenetratedNeeded(collarPartName);

                   isTopRightCornerPhysicalConnectionNeeded = PenetrationsServices.IsPhysicalConnectionBetweenCollarAndPenetratedNeeded(collarPartName);

                   isRightPhysicalConnectionNeeded = PenetrationsServices.IsPhysicalConnectionBetweenCollarAndPenetratedNeeded(collarPartName);

                   isBottomFlangeRightBottomFreeEdgeTreatmentNeeded = PenetrationsServices.IsBottomFlangeFreeEdgeTreatmentNeeded(collarPartName);

                   isBottomFlangeRightBottomCornerFreeEdgeTreatmentNeeded = PenetrationsServices.IsBottomFlangeFreeEdgeTreatmentNeeded(collarPartName);

                   isTopFlangeRightTopCornerPhysicalConnectionNeeded = PenetrationsServices.IsTopFlangeRightTopCornerPhysicalConnectionNeeded(collarPartName);

                   isTopPhysicalConnectionNeeded = PenetrationsServices.IsTopPhysicalConnectionNeeded(collarPartName);

                   isNormalSideAdditonalLapPhysicalConnectionNeeded = PenetrationsServices.IsNormalSideAdditonalLapPhysicalConnectionNeeded(collarPartName, (CollarPart)base.Occurrence, penetratedPlate, penetratedStiffener);

                   isOppositeSideAdditonalLapPhysicalConnectionNeeded = PenetrationsServices.IsOppositeSideAdditonalLapPhysicalConnectionNeeded(collarPartName, (CollarPart)base.Occurrence, penetratedPlate, penetratedStiffener);

                   isBaseRightAdditonalPhysicalConnectionNeeded = PenetrationsServices.IsBaseRightAdditonalPhysicalConnectionNeeded((CollarPart)base.Occurrence);
               }
                                
                requiredAssemblyOutputs.Add(WebRightTopCornerPhysicalConnection, isWebRightTopCornerPhysicalConnectionNeeded);

                requiredAssemblyOutputs.Add(TopFlangeRightBottomPhysicalConnection, isTopFlangeRightBottomPhysicalConnectionNeeded);

                requiredAssemblyOutputs.Add(TopFlangeRightBottomCornerPhysicalConnection, isTopFlangeRightBottomCornerPhysicalConnectionNeeded);

                requiredAssemblyOutputs.Add(TopFlangeRightPhysicalConnection, isTopFlangeRightPhysicalConnectionNeeded);

                requiredAssemblyOutputs.Add(NormalSideLapPhysicalConnection, isNormalSideLapPhysicalConnectionNeeded);

                requiredAssemblyOutputs.Add(OppositeSideLapPhysicalConnection, isOppositeSideLapPhysicalConnectionNeeded);

                requiredAssemblyOutputs.Add(SnipeCornerFeature, isSnipeCornerFeatureNeeded);

                requiredAssemblyOutputs.Add(TopRightPhysicalConnection, isTopRightPhysicalConnectionNeeded);
                
                requiredAssemblyOutputs.Add(TopRightCornerPhysicalConnection, isTopRightCornerPhysicalConnectionNeeded);

                requiredAssemblyOutputs.Add(RightPhysicalConnection, isRightPhysicalConnectionNeeded);

                requiredAssemblyOutputs.Add(BottomFlangeRightBottomFreeEdgeTreatment, isBottomFlangeRightBottomFreeEdgeTreatmentNeeded);

                requiredAssemblyOutputs.Add(BottomFlangeRightBottomCornerFreeEdgeTreatment, isBottomFlangeRightBottomCornerFreeEdgeTreatmentNeeded);

                requiredAssemblyOutputs.Add(TopFlangeRightTopCornerPhysicalConnection, isTopFlangeRightTopCornerPhysicalConnectionNeeded);

                requiredAssemblyOutputs.Add(TopPhysicalConnection, isTopPhysicalConnectionNeeded);

                requiredAssemblyOutputs.Add(NormalSideAdditonalLapPhysicalConnection, isNormalSideAdditonalLapPhysicalConnectionNeeded);

                requiredAssemblyOutputs.Add(OppositeSideAdditonalLapPhysicalConnection, isOppositeSideAdditonalLapPhysicalConnectionNeeded);

                requiredAssemblyOutputs.Add(BaseRightAdditonalPhysicalConnection, isBaseRightAdditonalPhysicalConnectionNeeded);



                return requiredAssemblyOutputs;
            }
        }

        #endregion Private Methods
    }
}