/******************************************************************************************
Copyright (C) 2013 Intergraph Corporation.  All rights reserved.
 
Class   :   AmbiguityResolver

ProgId  :   S3DAutoCorrelation.AmbiguityResolver
 
This is a sample rule that will be executed to resolve the ambiguities found during the autocorrelation process.
It is a delivered rule and the user may customize the rule as per their requirements.
The progid should be added to the respective publish map in order to activate this rule from Autocorrelation.

History:
19Apr2013    Sreelekha    DM232231  Correlation of Pump Nozzles that are Not Named  
******************************************************************************************/

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Equipment.Middle;
using Ingr.SP3D.ReferenceData.Middle;

namespace Ingr.S3D.AutoCorrelation.Rules
{
    public class AmbiguityResolver : CorrelationAmbiguityResolverBaseclass
    {
        private enum ErrorLevel
        {
            ERRORLEVEL_ErrInformation = 1,
            ERRORLEVEL_ErrWarning = 2,
            ERRORLEVEL_ErrCritical = 3,
        }

        /// <summary>
        /// Given a S3D object the method attempts to find a matching Design Basis object 
        /// from the ambiguous list presented to this rule.
        /// </summary>
        /// <param name="s3dBOCol"></param>collection of S3D objects (single 3d object currently supported)
        /// <param name="designbasisBOCol"></param>collection of ambiguous Design Basis objects
        /// <param name="s3dBO"></param>3D object for which a matching DB object is to be found
        /// <param name="designbasisBO"></param> returns matching Design Basis object. Null if a match is not found.
        public override void ResolveAmbiguities(ReadOnlyCollection<BusinessObject> s3dBOCol,
                                            ReadOnlyCollection<DesignBasis> designbasisBOCol,
                                            ref BusinessObject s3dBO,
                                            ref DesignBasis designbasisBO)
        {
            try
            {
                designbasisBO = null;

                // validate the input parameters
                if (s3dBOCol == null)
                    throw new ArgumentNullException("s3dBOCol");

                if (designbasisBOCol == null)
                    throw new ArgumentNullException("designbasisBOCol");

                if (s3dBOCol.Count < 1 || designbasisBOCol.Count < 1)
                    throw new ArgumentException("input collection is empty");

                // Note: In the line below the first object is retrieved since the variable 's3dBOCol' 
                // currently holds single 3D object. Need to revisit the code later when multiple 3D objects 
                // are supported in future.
                s3dBO = s3dBOCol.First();

                // Use the service to find the mapped property values for 3D and Design Basis objects
                CorrelationAmbiguityResolverService xmlService = new CorrelationAmbiguityResolverService();

                // Iterate through the ambiguous Design Basis equipment nozzle Collection. For each nozzle, retrieve 
                // the Design Basis parent (i.e Design Basis Equipment) by traversing the respective relationship. 
                // Verify the type of Equipment that the nozzle belongs to - If the equipment is a pump, 
                // use the flowdirection to find a matching Design Basis nozzle otherwise use the nozzle tag to find a match.

                object s3dBOPropertyVal, designbasisBOPropertyVal;

                for (int iIndex = 0; iIndex < designbasisBOCol.Count; iIndex++)
                {
                    DesignBasis designbasisChildBO = designbasisBOCol[iIndex];

                    // Retrieve the parent by traversing 'EquipmentNozzle' relation if it exists.
                    RelationCollection designbasisParentRelCol = designbasisChildBO.GetRelationship("EquipmentNozzle", "Nozzles");
                    ReadOnlyCollection<BusinessObject> designbasisParentBOCol = designbasisParentRelCol.TargetObjects;

                    if (designbasisParentBOCol.Count == 0)
                    {
                        //Traverse 'EquipmentComponentComposition' relation to retrieve the parent
                        designbasisParentRelCol = designbasisChildBO.GetRelationship("EquipmentComponentComposition", "Equipment");
                        designbasisParentBOCol = designbasisParentRelCol.TargetObjects;
                    }

                    if (designbasisParentBOCol.Count < 1 || designbasisParentBOCol.Count > 1)
                        continue;

                    DesignBasis designbasisParentBO = (DesignBasis)designbasisParentBOCol[0];      // Design Basis Parent

                    s3dBOPropertyVal = null;
                    designbasisBOPropertyVal = null;

                    if (designbasisParentBO.SupportsInterface("IPumpOcc"))
                    {
                        //Design Basis parent is a pump. Now confirm whether the S3D parent is also a pump  
                        RelationCollection s3dParentRelCol = s3dBO.GetRelationship("DistribPorts", "DistribPart");     // Get s3D parent
                        ReadOnlyCollection<BusinessObject> s3dParentBOCol = s3dParentRelCol.TargetObjects;

                        if ((s3dParentBOCol.Count < 1) || (s3dParentBOCol.Count > 1))
                            continue;

                        BusinessObject s3dParentBO = s3dParentBOCol[0];      // 3D Parent

                        // Get the Equipment catalog part and verify whether it is a pump
                        Equipment s3dBOEquipment = (Equipment)s3dParentBO;
                        Part s3dCatalogPart = s3dBOEquipment.Part;
                        if ((s3dCatalogPart != null) && (s3dCatalogPart.SupportsInterface("IJEquipmentPart")))
                        {
                            PropertyValue oPropertyValue = s3dCatalogPart.GetPropertyValue("IJEquipmentPart", "ProcessEqTypes3");
                            PropertyValueCodelist oPropertyValueCodelist = (PropertyValueCodelist)oPropertyValue;
                            if (oPropertyValueCodelist.PropValue == 325)    // equipment is a pump
                            {
                                //Get the Flow Direction of the 3D and Design Basis objects                                        
                                xmlService.GetPropertyAndMappedProperty(s3dBO, designbasisChildBO, "IJDPipePort", "IPipingPortComposition:Piping Ports", "FlowDirection", ref s3dBOPropertyVal, ref designbasisBOPropertyVal);

                                // Compare the 3D and Design Basis property values
                                if ((s3dBOPropertyVal != null) && (designbasisBOPropertyVal != null))
                                {
                                    if (s3dBOPropertyVal.Equals(designbasisBOPropertyVal)) // found a match
                                    {
                                        if (designbasisBO == null) // first match
                                        {
                                            designbasisBO = designbasisChildBO; // save it and keep looking
                                        }
                                        else // not first match (ambiguous)
                                        {
                                            designbasisBO = null; // return no matches
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                    }
                } // end For loop
            }
            catch (Exception oEx)
            {
                LogErrorMessage(ErrorLevel.ERRORLEVEL_ErrInformation, this.ToString() + ".ResolveAmbiguities", oEx);
            }
        }

        private void LogErrorMessage(ErrorLevel errLevel, string sMethod, Exception exp)
        {
            LogError logError = MiddleServiceProvider.ErrorLogger;
            string componentName = "AmbiguityResolver : ";
            logError.Log(componentName + sMethod + "" + exp.Message + "", (int)errLevel);
            return;
        }
    }
}
