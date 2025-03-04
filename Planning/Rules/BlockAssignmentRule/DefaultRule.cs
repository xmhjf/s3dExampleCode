using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Common.Middle.Services.Hidden;
using Ingr.SP3D.Planning.Middle;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Ingr.SP3D.Content.Planning
{
    public class DefaultRule: PlanningBlockAssignmentRuleBase
    {        
        private const string strTracePath = @"SOFTWARE\Intergraph\Applications\Environments\Planning\Debug";
        private const string overLapValue = "OverLappingValue";

        public override bool IsValidPart(BusinessObject part)
        {
            #region Input Error Handling
            if (part == null)
            {
                throw new ArgumentNullException("Input Prod Routing Data is null");
            }
            #endregion //Input Error Handling

            bool returnValue = false;
            try
            {
                returnValue = CheckValidPart(part);               
            }
            catch (Exception e)
            {                
                MiddleServiceProvider.ErrorLogger.Log(" (" + e.Message + ")");
            }
            return returnValue;
        }
        
        public override void GetPartAssignmentData(BusinessObject part, PlanningBlockAssignmentMethods assignMethod, out double overlapFactor)
        {
            #region Input Error Handling
            if (part == null)
            {
                throw new ArgumentNullException("Input Part is null");
            }
            #endregion //Input Error Handling

            overlapFactor = 0.0;

            try
            {
                bool assignable = IsPartAssignable(part);
                if (assignable == true)
                {
                    RegistryKey localMachineRegistryKey = Registry.LocalMachine;
                    RegistryKey installationSubKey;

                    installationSubKey = localMachineRegistryKey.OpenSubKey(strTracePath);

                    object varValue = null;
                    
                    if (installationSubKey != null)
                    {
                        varValue = installationSubKey.GetValue(overLapValue);
                    }

                    if (!(varValue == null) && !(varValue.ToString() == ""))
                    {
                        overlapFactor = Convert.ToDouble(varValue);
                    }
                    else
                    {
                        bool plnSeamsPresent = false; 
                        assignMethod = ByRangeOrGeometry(part, ref plnSeamsPresent); 

                        if (assignMethod == PlanningBlockAssignmentMethods.ByMinRange)
                        {
                            if (plnSeamsPresent == true)
                            {
                                overlapFactor = 0.8;
                            }
                            else
                            {
                                overlapFactor = 0.95;
                            }
                        }
                        else if (assignMethod == PlanningBlockAssignmentMethods.ByRange)
                        {
                            overlapFactor = 0.95;
                        }
                        else
                        {
                            overlapFactor = 1.0;
                        }
                    }
                    StopTimer();
                }
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log(" (" + e.Message + ")");
            }
        }


        private PlanningBlockAssignmentMethods ByRangeOrGeometry(BusinessObject part, ref bool bPlnSeamsPresent)
        {
            PlanningBlockAssignmentMethods plnBlockAssignmentMethods = PlanningBlockAssignmentMethods.ByRange; 

            try
            {                
                ApplicationDomains ePlnApplicationDomains;
                
                bPlnSeamsPresent = false;

                // Only allow split for structure.     
                ePlnApplicationDomains = GetApplicationDomain(part);

                if (ePlnApplicationDomains == ApplicationDomains.DomainStructure)
                {
                    plnBlockAssignmentMethods = PlanningBlockAssignmentMethods.ByGeometry;
                    
                    if (IsPartCreatedByPlanningSeam(part))
                    {
                        plnBlockAssignmentMethods = PlanningBlockAssignmentMethods.BySmallVol;
                        bPlnSeamsPresent = true;
                    }
                }
                else if (ePlnApplicationDomains == ApplicationDomains.DomainOutfitting)
                {
                    BusinessObject oParentEntity = (BusinessObject)((ISystemChild)part).SystemParent;

                    if ((oParentEntity != null))
                    {                        
                        if (oParentEntity.SupportsInterface("IJRouteSplit"))
                        {
                            plnBlockAssignmentMethods = PlanningBlockAssignmentMethods.ByGeometry;
                        }
                    }
                }
                else
                {
                    plnBlockAssignmentMethods = PlanningBlockAssignmentMethods.ByRange;
                }
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log(" (" + e.Message + ")");
            }
            return plnBlockAssignmentMethods;
        }
    }
}
