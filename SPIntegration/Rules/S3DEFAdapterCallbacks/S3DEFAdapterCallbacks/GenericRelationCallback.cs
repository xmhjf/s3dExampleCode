using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Ingr.SP3D.Common.Middle;
using Ingr.SP3D.Common.Middle.Services;
using Ingr.SP3D.Route.Middle;
using System.Collections;
using System.Collections.ObjectModel;
using Ingr.SP3D.Structure.Middle;
using Ingr.SP3D.ReferenceData.Middle;

namespace Ingr.SP3D.Content.EFAdapterCallbacks
{
	/// <summary>
	/// This code is delivered to end user so that he can modify the same to change the default behavior of the EF Adapter.
	/// RelationCallbacks class inherits from EFRelationCallbacks class and overrides methods to provide custom behavior for the rules.
	/// </summary>
	public class GenericRelationCallback : RelationCallback
	{
		/// <summary>
		/// ProcessRelation
		/// <para>Returns: 
		///     Process a relation using an Info map.</para>
		/// </summary>     
		/// <param name="oInfo">The information data for the callback.</param>
		/// <returns>returns input value for base implementation.</returns>
		public override bool ProcessRelation(RelationCallbackInformation oInfo)
		{
			return true;
		}

		/// <summary>
		/// SelectSameAs
		/// <para>Returns: 
		///     Select a SameAS relation from a list.</para>
		/// </summary>     
		/// <param name="oInfo">The information data for the callback.</param>
		/// <returns>returns input value for base implementation.</returns>
		public override bool SelectSameAs(RelationCallbackInformation oInfo)
		{
			oInfo.Value = (int)0;
			return true;
		}

		/// <summary>
		/// Processes IJCustomRelation2 MakeSameAs.
		/// <para>Returns: 
		///     Checks input to determine if the relation should be converted to a SameAs.</para>
		/// </summary>     
		/// <param name="oInfo">The information data for the callback.</param>
		/// <returns>returns true for base implementation.</returns>
		public override bool MakeSameAs(RelationCallbackInformation oInfo)
		{
			oInfo.Value = (bool)true;
			return true;
		}
	}
}
