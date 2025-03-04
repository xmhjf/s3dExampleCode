using System;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using System.Collections.ObjectModel;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Ingr.SP3D.Content.Interference
{
    /// <summary>
    /// Naming rule computes name for a interference. 
    /// NameRuleBase Class contains common implementation across all name rules and provides properties and methods for use in implementation of custom name rules. 
    /// </summary>
    [WrapperProgID("NameRuleWrapperForNet.CNameRuleWrapper")]
    public class  NamingRule : NameRuleBase
    {
        public NamingRule()
        {
        }
        /// <summary>
        /// Computes a name for the given interference. 
        /// </summary>
        /// <param name="businessObject">The business object whose name is being computed. </param>
        /// <param name="parents">Naming parents of the business object, 
        /// i.e. other business objects which control the naming of the business object whose name is being computed.</param>
        /// <param name="activeEntity">The name rule active entity associated to the business object whose name is being computed.</param>
        public override void ComputeName(BusinessObject businessObject, ReadOnlyCollection<BusinessObject> Parents, BusinessObject activeEntity)
        {
            String objectName = string.Empty;

            if (businessObject == null)
            {
                throw new ArgumentNullException("businessObject");
            }

            if (activeEntity == null)
            {
                throw new ArgumentNullException("activeEnity");
            }

            INamedItem oIfcNamedEnt = businessObject as INamedItem;

            if (oIfcNamedEnt != null)
            {
                objectName = oIfcNamedEnt.Name;
                foreach (BusinessObject oParent in Parents)
                {
                    //string namingParentsString = GetNamingParentsString(activeEntity);   
                    INamedItem oParentName = oParent as INamedItem;
                    if (oParentName != null)
                    {
                        objectName += "-" + oParentName.Name;
                        SetName(businessObject, objectName);
                    }
                    else
                    {
                        ClassInformation oClassInfo = oParent.ClassInfo as ClassInformation;
                        if (oClassInfo != null)
                        {
                            string strNodeName = string.Empty;
                            BOCInformation oBOCInfo = oClassInfo.BOC as BOCInformation;
                            //if the business object is reference object UserClassInfo has the BOC Display Name
                            if (oBOCInfo == null)
                            {
                                ClassInformation oUserClassInfo = oParent.UserClassInfo;
                                if (oUserClassInfo != null)
                                {
                                    oBOCInfo = oUserClassInfo.BOC;
                                    if (oBOCInfo != null)
                                        strNodeName = oBOCInfo.DisplayName;
                                }
                            }
                            else
                            {
                                //if the business object is S3D object ClassInfo has the BOC Display Name.
                                strNodeName = oBOCInfo.DisplayName;
                            }
                            objectName += "-" + strNodeName;
                            SetName(businessObject, objectName);
                        }

                    }
                }
            }
        }

        /// <summary>
        /// Gets the naming parents.
        /// All the Naming Parents that need to participate in an objects naming are added here to the BusinessObject collection. 
        /// The parents added here are used in computing the name of the object in ComputeName() of the same interface. 
        /// Both these methods are called from naming rule semantic. 
        /// </summary>
        /// <param name="entity">BusinessObject for which naming parents are required.</param>
        /// <returns>Returns the collection of naming parents.</returns> 
        public override Collection<BusinessObject> GetNamingParents(BusinessObject Entity)
        {
            return new Collection<BusinessObject>();
        }
    }
}
