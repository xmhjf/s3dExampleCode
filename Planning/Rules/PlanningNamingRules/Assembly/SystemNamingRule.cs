using Ingr.SP3D.Common.Middle;
using System;
using Ingr.SP3D.Planning.Middle;
using Ingr.SP3D.Common.Middle.Services;
using System.Collections.ObjectModel;

namespace Ingr.SP3D.Content.Planning
{
    public class SystemNamingRule : NameRuleBase
    {
        public override void ComputeName(BusinessObject oEntity, ReadOnlyCollection<BusinessObject> oParents, BusinessObject oActiveEntity)
        {
            try
            {
                string parentName = null;

                if (oEntity == null)
                {
                    throw new ArgumentNullException();
                }

                if (oParents.Count > 0)
                {
                    BusinessObject parent = oParents[0];
                    if (parent != null)
                    {
                        parentName = Convert.ToString(parent.GetPropertyValue("IJNamedItem", "Name"));
                        oEntity.SetPropertyValue(parentName, "IJNamedItem", "Name");
                    }
                }
                oEntity.SetPropertyValue(parentName, "IJNamedItem", "Name");
            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log("KoreanNamingRule.ComputeName: Error encountered (" + e.Message + ")");
            } 
        }

        public override Collection<BusinessObject> GetNamingParents(BusinessObject oEntity)
        {
            Collection<BusinessObject> oNamingParents = new Collection<BusinessObject>();
            try
            {
                Assembly assembly = oEntity as Assembly;

                if (assembly != null)
                {
                   ReadOnlyCollection<IAssemblyChild> childrens= assembly.AssemblyChildren;

                   if (childrens  != null)
                   {
                       foreach (IAssemblyChild child in childrens)
                       {
                           if (child is ISystemChild )
                           {
                               if (child is ISystem)
                               {
                                   ISystemChild systemChild = child as ISystemChild;

                                   if (systemChild != null)
                                   {
                                       ISystem sysparent = systemChild.SystemParent;
                                       bool sysChild = false;

                                       while (sysChild == false)
                                       {
                                           systemChild = sysparent as ISystemChild;

                                           if (systemChild != null)
                                           {
                                               sysparent = systemChild.SystemParent;
                                               Model modelDBConnection = oEntity.DBConnection as Model;
                                               if (modelDBConnection != null)
                                               {
                                                   ISystem rootsystem = modelDBConnection.RootSystem;
                                                   if (rootsystem.Equals(sysparent))
                                                   {
                                                       sysChild = true;
                                                       BusinessObject parent = (BusinessObject)systemChild;
                                                       oNamingParents.Add(parent);
                                                   }
                                               }                                               
                                           }
                                       }
                                   }                                  
                               }
                               else if (!(child is ISystem))
                               {
                                   ISystemChild systemChild = child as ISystemChild;
                                   if (systemChild != null)
                                   {
                                       BusinessObject parent = (BusinessObject)systemChild.SystemParent;
                                       oNamingParents.Add(parent);
                                   }                                   
                               }
                           }
                       }
                   }
                }

            }
            catch (Exception e)
            {
                MiddleServiceProvider.ErrorLogger.Log("KoreanNamingRule.GetNamingParents: Error encountered (" + e.Message + ")");
            }
            return oNamingParents;
        }
    }
}
