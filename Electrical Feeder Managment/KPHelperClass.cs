/*
 * Created by SharpDevelop.
 * User: kpphillips
 * Date: 11/21/2015
 * Time: 2:44 PM
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Diagnostics;
//using WinForms = System.Windows.Forms;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Electrical;
using Autodesk.Revit.DB.Mechanical;

namespace ElectricalMacros
{
	/// <summary>
	/// Description of KPHelperClass.
	/// </summary>
	public class KPHelperClass
	{
		public KPHelperClass()
		{
		
		}
		
		/// <summary>
        /// Retrieve all electrical circuit elements in the given document,
        /// identified by the built-in category OST_ElectricalCircuit.
        /// </summary>
        public static List<Element> GetElectricalCircuits(Document doc)
        {
            FilteredElementCollector collector
              = GetElementsOfType(doc, typeof(ElectricalSystem),
                BuiltInCategory.OST_ElectricalCircuit);

            // return a List instead of IList, because we need the method Exists() on it:

            return new List<Element>(collector.ToElements());
        }
        
        /// <summary>
        /// Retrieve all electrical conduit elements in the given document,
        /// identified by the built-in category OST_Conduits.
        /// </summary>
        public static List<Element> GetConduits(Document doc)
        {
            FilteredElementCollector collector
              = GetElementsOfType(doc, typeof(Conduit),
                BuiltInCategory.OST_Conduit);

            // return a List instead of IList, because we need the method Exists() on it:

            return new List<Element>(collector.ToElements());
        }
        
        /// <summary>
        /// Retrieve all electrical conduit fitting elements in the given document,
        /// identified by the built-in category OST_ConduitFittings.
        /// </summary>
        public static List<Element> GetConduitFittings(Document doc)
        {
            FilteredElementCollector collector
              = GetElementsOfType(doc, typeof(FamilyInstance),
                BuiltInCategory.OST_ConduitFitting);

            // return a List instead of IList, because we need the method Exists() on it:

            return new List<Element>(collector.ToElements());
        }
        
        
        
        /// <summary>
        /// Return all elements of the requested class i.e. System.Type
        /// matching the given built-in category in the active document.
        /// </summary>
        public static FilteredElementCollector GetElementsOfType(
          Document doc,
          Type type,
          BuiltInCategory bic)
        {
            FilteredElementCollector collector = new FilteredElementCollector(doc);

            collector.OfCategory(bic);
            collector.OfClass(type);

            return collector;
        }
        
        
        
        #region Parameter Access
        /// <summary>
        /// Helper to get a specific parameter by name.
        /// </summary>
        public static Parameter GetParameterFromName(Element elem, string name)
        {
            foreach (Parameter p in elem.Parameters)
            {
                if (p.Definition.Name == name)
                {
                    return p;
                }
            }
            
            return null;
        }

        public static Definition GetParameterDefinitionFromName(Element elem, string name)
        {
            Parameter p = GetParameterFromName(elem, name);
            return (null == p) ? null : p.Definition;
        }

        public static double GetParameterValueFromName(Element elem, string name)
        {
            Parameter p = GetParameterFromName(elem, name);
            if (null == p)
            {
                throw new ParameterException(name, "element", elem);
            }
            return p.AsDouble();
        }

        public static string GetStringParameterValueFromName(Element elem, string name)
        {
            Parameter p = GetParameterFromName(elem, name);
            if (null == p)
            {
                throw new ParameterException(name, "element", elem);
                //return "No Value";
            }
            return p.AsString();
        }

        static void DumpParameters(Element elem)
        {
            foreach (Parameter p in elem.Parameters)
            {
                Debug.WriteLine(p.Definition.ParameterType + " " + p.Definition.Name);
            }
        }

  
        #endregion // Parameter Access
		
        
        
        #region Exceptions
        public class ParameterException : Exception
        {
            public ParameterException(string parameterName, string description, Element elem)
                : base(string.Format("'{0}' parameter not defined for {1} {2}", parameterName, description, elem.Id.IntegerValue.ToString()))
            {
            	
            	
            	
            }
        }
        
        #endregion // Exceptions
		
	}
}
