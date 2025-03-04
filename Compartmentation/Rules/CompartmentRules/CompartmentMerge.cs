using Ingr.SP3D.Common.Exceptions;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Compartmentation.Middle;
using Ingr.SP3D.ReferenceData.Middle;
using Ingr.SP3D.ReferenceData.Middle.Services;
using System;

namespace Ingr.SP3D.Content.Compartmentation
{
    public class CompartmentMerge : CompartmentMergeRuleBase
    {
        public override void Evaluate(Compartment compartment, Part targetPart)
        {
            #region Input Error Handling
            if (compartment == null || targetPart== null)
            {
                throw new CmnArgumentNullException("input arguments are null");
            }
            #endregion
           
            try
            {
                string inputPartClassName =string.Empty, targetPartClassName =string.Empty;
                Part part = compartment.Part;

                if (part != null)
                {
                    PartClass inputPartClass = (PartClass)part.PartClass;
                    PartClass targetPartClass = (PartClass)targetPart.PartClass;

                    if (inputPartClass != null && targetPartClass != null)
                    {
                        inputPartClassName = inputPartClass.PartClassType;
                        targetPartClassName = targetPartClass.PartClassType;
                    }

                    if ((inputPartClassName == targetPartClassName) && (inputPartClassName == "CompartmentClass"))
                    {
                        CatalogBaseHelper catalogBaseHelper = new CatalogBaseHelper();
                        PartClass catalogPartClass = catalogBaseHelper.GetPartClass("VoidSpace") as PartClass;

                        if (catalogPartClass != null && catalogPartClass.Parts.Count > 0)
                        {
                            Part catalogPart = (Part)catalogPartClass.Parts[0];

                            if (part.PartNumber != catalogPart.PartNumber)
                            {
                                compartment.Replace(catalogPart);
                            }
                        }
                    }
                }               
            }
            catch (CmnException e)
            {

                MiddleServiceProvider.ErrorLogger.Log(" (" + e.Message + ")");
            }            
        }
    }
}
