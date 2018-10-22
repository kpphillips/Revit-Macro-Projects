/*
 * Created by SharpDevelop.
 * User: KPhillips
 * Date: 11/20/2015
 * Time: 11:15 AM
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB.Electrical;
using System.Text;
using Microsoft.Office.Interop.Excel;

namespace ElectricalMacros
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.DB.Macros.AddInId("CD6BE1CA-707F-4B5F-9EFD-576432E5474D")]
	public partial class ThisDocument
	{
		
		
		//global vars
		const string sharedParameterFile = "nil";
		//private Autodesk.Revit.ApplicationServices.Application app;
		const string FEEDER_PARAMETER_DESIGNATOR_IN_REVIT_AND_EXCEL = "FEEDER_PARAMETER_DESIGNATOR_IN_REVIT_AND_EXCEL";
		
		private void Module_Startup(object sender, EventArgs e)
		{

		}

		private void Module_Shutdown(object sender, EventArgs e)
		{

		}

		#region Revit Macros generated code
		private void InternalStartup()
		{
			//Autodesk.Revit.ApplicationServices.Application app = new Autodesk.Revit.ApplicationServices.Application();
			//this.app = app;
			
			this.Startup += new System.EventHandler(Module_Startup);
			this.Shutdown += new System.EventHandler(Module_Shutdown);
		}
		#endregion
		
		
		#region Macro Methods
		
		public void A_SetUpSharedParameterFile(){

			//Document doc = this.ActiveUIDocument.Document;
			Document doc = this.Document;
			
			//initialize sp file and parameters with helper functions
			SetUpSharedParamFile(doc);
		
		}
		
		//Macro Command
		//Creates Shared Parameters (Hard Coded) into external shared parameter file. Note: Should match Excel file headers 
		//TODO: Read parameters from columns created in Excel file, future update
		public void B_CreateFeederSharedParameters(){
			
			//Name Constants
			//**will get this from a database later
			const string kSharedParamsGroup = "FD_Conduit_Feeders";
			const string kSharedParamsDefA = "FD_FEEDER DESIGNATOR";
			const string kSharedParamsDefB = "FD_SCH 40 PVC";
			const string kSharedParamsDefC = "FD_CONDUCTOR(S)";
			const string kSharedParamsDefD = "FD_GROUND";
			const string kSharedParamsDefE = "FD_FROM";
			const string kSharedParamsDefF = "FD_TO";

			UIApplication uiapp = this.Application;
			
			DefinitionFile defFile = uiapp.Application.OpenSharedParameterFile();
			
			if (defFile == null){
				TaskDialog.Show("Error, no Shared Parameter File!",
				                "You need to create shared parameter file in project directory first, yo!");
				return;
			}
			
			//create a group
			DefinitionGroup sharedParamsGroup = GetOrCreateSharedParamsGroup(defFile,kSharedParamsGroup);
			//create parameters
			GetOrCreateSharedParamsDefinition(sharedParamsGroup, ParameterType.Text, kSharedParamsDefA);
			GetOrCreateSharedParamsDefinition(sharedParamsGroup, ParameterType.Text, kSharedParamsDefB);
			GetOrCreateSharedParamsDefinition(sharedParamsGroup, ParameterType.Text, kSharedParamsDefC);
			GetOrCreateSharedParamsDefinition(sharedParamsGroup, ParameterType.Text, kSharedParamsDefD);
			GetOrCreateSharedParamsDefinition(sharedParamsGroup, ParameterType.Text, kSharedParamsDefE);
			GetOrCreateSharedParamsDefinition(sharedParamsGroup, ParameterType.Text, kSharedParamsDefF);
			
			ShowDefinitionFileInfo(defFile);
			
		}
		
		//Macro Command
		//Adds Project Parameters from Shared Parameter file to (Hard Coded) categories indicated 
		public void C_AddProjectSharedParameters(){
			//Autodesk.Revit.ApplicationServices.Application app = new Autodesk.Revit.ApplicationServices.Application();
			
			Document doc = this.Document;
			UIApplication uiapp = this.Application;
			
			DefinitionFile defFile = uiapp.Application.OpenSharedParameterFile();
			
			// Create a category set and insert category "XX" to it
    		CategorySet myCategories = uiapp.Application.Create.NewCategorySet();
    		// Use BuiltInCategory to get relevant categories
    		//Category myCategory = doc.Settings.Categories.get_Item(BuiltInCategory.OST_ElectricalCircuit);
    		Category myCategory1 = doc.Settings.Categories.get_Item(BuiltInCategory.OST_Conduit);
    		Category myCategory2 = doc.Settings.Categories.get_Item(BuiltInCategory.OST_ConduitFitting);
    		
    		myCategories.Insert(myCategory1);
    		myCategories.Insert(myCategory2);
    		
			//debug show the category 
			foreach(Category cat in myCategories){
    			//TaskDialog.Show("Categories", cat.Name);
    			
    			}
			//Create an object of InstanceBinding according to the Categories
			InstanceBinding instanceBinding = uiapp.Application.Create.NewInstanceBinding(myCategories);
			
    		// Get the BingdingMap of current document.
    		BindingMap bindingMap = doc.ParameterBindings;
    		
    		//get the group
    		//TODO: select this from a dialog later, hard coding not ideal. 
    		DefinitionGroup myGroup = GetOrCreateSharedParamsGroup(defFile, "FD_Conduit_Feeders");
    	
    		//get the definitions
    		//Definition def = GetOrCreateSharedParamsDefinition(myGroup, ParameterType.Text, "FD_FEEDER DESIGNATOR");
    		
    		//Bind the definitions to the document
    		//transaction
    		Transaction trans = new Transaction(doc);
    		trans.Start("Binding Project Shared Parameters");
    		
    		//bool typeBindOK = bindingMap.Insert(def, instanceBinding, BuiltInParameterGroup.PG_IDENTITY_DATA);
    		bool typeBindOK = false;
    		
    		// iterate the defs in group
        	foreach (Definition definition in myGroup.Definitions)
        	{
            	// get definition name
            	//TaskDialog.Show("Definition Name: ", definition.Name);
            	//add paremeters under the construction group in the Revti UI
            	typeBindOK = bindingMap.Insert(definition, instanceBinding, BuiltInParameterGroup.PG_CONSTRUCTION);
        	}
    		
    		trans.Commit();

    		if(typeBindOK){
    			TaskDialog.Show("Type Binding", "Success, on creating Project Shared Parameters");
    		} else {
    			TaskDialog.Show("Type Binding", "No Shared Parameters added. Either they exist already or Shared Parameter file not created");
    		}
			
		}
		
		//Macro Command
		//Retrieves the Circuits in the Project, Current and does some Overcurrent calcs
		public void D_GetElectricalCircuits(){			
			//Document doc = this.ActiveUIDocument.Document;
			Document doc = this.Document;

            List<Element> circuits = KPHelperClass.GetElectricalCircuits(doc);
       		string str = "";
            foreach(Element e in circuits){
            	double amps = KPHelperClass.GetParameterValueFromName(e, "Apparent Current");
            	string panel = KPHelperClass.GetStringParameterValueFromName(e, "Panel");
            	string number = KPHelperClass.GetStringParameterValueFromName(e, "Circuit Number");
            	double amps125percent = amps * 1.25;
            	
            	str += "Circuit: " + panel + "-" + number
            		+ " Connected Load ="
            		+ amps.ToString("0.00")
            		+ " -> " + "Load @ 125%: " 
            		+ amps125percent.ToString("0.00")
            		+ System.Environment.NewLine;
            	
            	//TODO: check what rating needs to be 
            	double ratingProtection = CheckRatingProtection(amps125percent);            	
            	
            	//TODO: set rating value, double amps
       		
       		}
            
            TaskDialog.Show("Circuits", str);
            
		}
		
		//Macro Command
		public void E_GetElementsInConduitRunFromSelectedConduitAndUpdateParameters(){
			
			Document doc = this.Document;
			ISelectionFilter selFilter = new E_ConduitSelectionFilter();
			Reference myRef = Selection.PickObject(ObjectType.Element, selFilter, "Select Conduit");
			Element e = doc.GetElement(myRef);
			
			//List to hold elements in condit run
			List<ElementId> elementsInConduitRun = new List<ElementId>();
			
			//Iterate thru connectors of conduit run
			Conduit c = e as Conduit;
			
			if (c != null){				
				elementsInConduitRun = E_AddConduitAndNeighborFittingsToSelection(c, elementsInConduitRun);
			}
			
			string str = "";
			List<Element> combinedElements = new List<Element>();
			str += "Count: " + elementsInConduitRun.Count.ToString() + "\n";
			foreach(ElementId eId in elementsInConduitRun){				
				Element e2 = doc.GetElement( eId ) as Element;
				combinedElements.Add(e2);
				str += " - " + eId.ToString() + "\n";
			}
			
			
			TaskDialog.Show("Conduit Info to be updated", str);
			
            UpdateParameterValuesForElements(combinedElements, true);
		
		}
		
		//Helper Class for filtering selection to Conduit Category
		private class E_ConduitSelectionFilter : ISelectionFilter
		{
    		public bool AllowElement(Element element)
    		{
        		if (element.Category.Name == "Conduits")
        		{
            		return true;
        		}
        		return false;
    		}

    		public bool AllowReference(Reference refer, XYZ point)
    		{
        		return false;
    		}
		}
		
		//helper method to ittereate thru connectors of a run
		private List<ElementId> E_AddConduitAndNeighborFittingsToSelection(Conduit conduit, List<ElementId> elementsInConduitRun){
			
			Document doc = this.Document;
			
			ConnectorSet connectors = conduit.ConnectorManager.Connectors;
			string str = "";
			
			//Get Conduit and its neighbor connected fitting elements
			foreach (Connector con in connectors) {				
				if (con.IsConnected) {		
					if (!elementsInConduitRun.Contains(con.Owner.Id)){
							elementsInConduitRun.Add(con.Owner.Id);
							str += "Conduit: " + con.Owner.Id.ToString() + "\n";
						}
					ConnectorSet refs = con.AllRefs;
					foreach (Connector conRef in refs) {
						if (conRef.IsConnected) {							
							if (!elementsInConduitRun.Contains(conRef.Owner.Id)){
								elementsInConduitRun.Add(conRef.Owner.Id);
								str += "Conduit Fitting: " + conRef.Owner.Id.ToString() + "\n";
								
								//Get the fittings neighbor conduit and recall the main method
								Categories categories = doc.Settings.Categories;
								ElementId conduitFittingCategoryId = categories.get_Item(BuiltInCategory.OST_ConduitFitting).Id;
								FamilyInstance fi = conRef.Owner as FamilyInstance;
									if (fi != null && fi.Category.Id == conduitFittingCategoryId)
										{
											ConnectorSet cons = fi.MEPModel.ConnectorManager.Connectors;
											foreach (Connector c2 in cons)
												{
												ConnectorSet c2Refs = c2.AllRefs;
													foreach(Connector c2Ref in c2Refs)
													{
														if (!elementsInConduitRun.Contains(c2Ref.Owner.Id))
														{
															str += "Conduit: "+ c2Ref.Owner.Id.ToString() + "\n";	
															Conduit c = c2Ref.Owner as Conduit;
															if (c != null)
															{
																
																//TaskDialog.Show("Before Call", elementsInConduitRun.Count.ToString());
																elementsInConduitRun = E_AddConduitAndNeighborFittingsToSelection(c, elementsInConduitRun);
																
															}
														}							
													}
												}
										}
									
							}
						}
					}				
					//str += "\n";
				}			
			}

			
			
			//TaskDialog.Show("All Itereated Conduit Info last", str + "\n" + "\n" + elementsInConduitRun.Count.ToString());
			return elementsInConduitRun;
			
		}
		
		//Macro Command
		//Update Conduit Feeder Parameters that already have Feeder Designators added
		public void F_UpdateAllProjectConduitAndFittingParamsFromExcel()
		{
			
			 Document doc = this.Document;
            
             List<Element> conduits = KPHelperClass.GetConduits(doc);
             List<Element> conduitFittings = KPHelperClass.GetConduitFittings(doc);
             
             List<Element> combinedElements = new List<Element>();
             foreach (Element cf in conduitFittings){
             	conduits.Add(cf);
             	combinedElements = conduits;
             }
			
             if (combinedElements != null){
             	UpdateParameterValuesForElements(combinedElements, false);
             }

             
		}
		
		//Helper method
		private void UpdateParameterValuesForElements(List<Element> combinedElements, bool isConduitRun){
			
			Document doc = this.Document;
			
			 string feederDesignatorParameterName = "";
             Dictionary<string, List<Tuple<string, string>>> dictXlFeederData = new Dictionary<string,List<Tuple<string,string>>>();
             
             dictXlFeederData = ReadDataFromExcelAndReturn();
             
             List<Tuple<string, string>> dictItem = dictXlFeederData[FEEDER_PARAMETER_DESIGNATOR_IN_REVIT_AND_EXCEL];
             foreach (Tuple<string, string> valueItem in dictItem){
             	feederDesignatorParameterName = valueItem.Item1;
             }
             
			           
             TaskDialog.Show("data", dictXlFeederData.Count.ToString()
                             + "-Rows of Excel Feeder data read. " 
                             + "\n" + "\n" + "Attempting to apply this data to "
                             + combinedElements.Count.ToString()
                             + " Elements");
             
            Transaction trans = new Transaction(doc);
			trans.Start("Updated Conduit Run Parameters");
             
			
             //string str = "";
             int feederDesignatorParameterCount = 0;
             int updatedElements = 0;
             int totalElements = 0;
             //get feeder designator of selected conduit to be applied to conduit run
             //first element should be the selected element  
             string feederDesignatorParameterValue = KPHelperClass.GetStringParameterValueFromName(combinedElements[0], feederDesignatorParameterName);
             
             foreach (Element e in combinedElements){
             	totalElements ++;
             	
             	//use each elements feeder designator value if its not a conduit run
             	if (isConduitRun == false){
             		feederDesignatorParameterValue = KPHelperClass.GetStringParameterValueFromName(e, feederDesignatorParameterName);
             	}
                  	
             	if (feederDesignatorParameterValue != null && dictXlFeederData.ContainsKey(feederDesignatorParameterValue)){
             		
             		//set the feeder designator value based on selected conduit
             		Autodesk.Revit.DB.Parameter p1 = KPHelperClass.GetParameterFromName(e, feederDesignatorParameterName);
             		if (p1 != null) {
             				p1.Set(feederDesignatorParameterValue);            				
             				
             			}
             		
             		var value = dictXlFeederData[feederDesignatorParameterValue];
             		
             		foreach(Tuple<string, string> t in value){
             				
             			Autodesk.Revit.DB.Parameter p2 = KPHelperClass.GetParameterFromName(e, t.Item1);
             			
             			if (p2 != null) {
             				p2.Set(t.Item2);            				
             				
             			}
             		
             		}
             		updatedElements ++;
             		
             	} else {
             		feederDesignatorParameterCount ++;
             	}
             	
             }
             
             trans.Commit();
             
             TaskDialog.Show("Updated Parameters", "Out of " + totalElements.ToString() + " Elements: "
                             + "\n"
                             + "\n"
                             + feederDesignatorParameterCount.ToString()
                             + ": Element(s) did not have a Feeder Designator or did not match the Feeder Designator in the Excel Data"
                             + "\n" 
                             + "\n"
                             + updatedElements.ToString()
                             + ": Element(s) were updated successfully");
			
		}
		
		//Helper method
		private Dictionary<string, List<Tuple<string, string>>> ReadDataFromExcelAndReturn()
		{
			
			Document doc = this.Document;
			
			Microsoft.Office.Interop.Excel.ApplicationClass appExcel;
        	Workbook newWorkbook = null;
        	_Worksheet objsheet = null;
        	Range range;
        	//group xcel data in a dictionaty so we can lookup by Feeder Designator as the key
             Dictionary<string, List<Tuple<string, string>>> dictXlFeederData = new Dictionary<string,List<Tuple<string,string>>>();
			
			string docPath = doc.PathName;
			string docName = doc.Title;
			if (docPath.EndsWith(".rvt")){
				docName += ".rvt";
			}
			
			int index = docPath.IndexOf(docName);
			string cleanPath = (index < 0)
   	 					? docPath
    					: docPath.Remove(index, docName.Length);
			
			string xlpath = cleanPath + "FEEDER SCHEDULE.xlsx";
				
			try
            {
                //TaskDialog.Show("Verify Files Exist", xlpath);
                appExcel = new Microsoft.Office.Interop.Excel.ApplicationClass();
                // then go and load this into excel
                newWorkbook = appExcel.Workbooks.Open(xlpath, true, true);
            }

            catch (System.Runtime.InteropServices.COMException)
            {

                TaskDialog.Show("WHOOPS", "Excel File does not exist. Create and format the following file:" + "\n" + "\n" + xlpath);
                return null;
            }
            
            
			//get the active worksheet
            objsheet = (_Worksheet)appExcel.ActiveWorkbook.ActiveSheet;

            //get the value that the lookup table will be using. 
            string feederDesignatorParameterName = objsheet.get_Range("A1").get_Value().ToString();
            //add it to the dictionary for later reference
            var dictValue = new List<Tuple<string, string>>();
            var dictValueItem = Tuple.Create( feederDesignatorParameterName, "" );
            dictValue.Add(dictValueItem);
            dictXlFeederData.Add(FEEDER_PARAMETER_DESIGNATOR_IN_REVIT_AND_EXCEL, dictValue);
            
            //confirm with user that the value that the calc value is correct
            TaskDialog.Show("Verify Parameter", "This Macro will use this parameter from excel to map to element parameter in Revit. " +
                            "Verify that the case is the same in Revit" + "\n" + "\n" + feederDesignatorParameterName);

            // Read Each Cell and Compare
            //if(Celltostring == string.empty)
			    	
			 int rowCount = 1;
             int columnCount = 1;
             //string rangeStr;
             range = objsheet.UsedRange;
             
             // create arrays that are the length on the excel range (-1 for bc the array starts w/ 0)
             //string[] columnHeaderParameters = new string[range.Columns.Count - 1];
             //string[] feederParameterValues = new string[range.Rows.Count - 1];             
                         
             
              
             int columnHeaderRow = 1;
             //start getting values in second row
             //reset to the first column            
             for (rowCount = 2; rowCount <= range.Rows.Count; rowCount++)
                {
                    
             		string feederDesignatorValue = "nil";
             		var rowData = new List<Tuple<string, string>>();
             		//itterate thru each colum for each row
             		for (columnCount = 1; columnCount <= range.Columns.Count; columnCount++)
                		{
                    		
             			if (columnCount == 1) {
             				
             				feederDesignatorValue = (string)(range.Cells[rowCount, columnCount] as Range).Value2.ToString();
             				
             				} else {
             					var columnHeader = (string)(range.Cells[columnHeaderRow, columnCount] as Range).Value2.ToString();
             					var rowValue = (string)(range.Cells[rowCount, columnCount] as Range).Value2.ToString();
             					rowData.Add(Tuple.Create( columnHeader, rowValue ));
             				}             				
             				
                		}
                     
                   	dictXlFeederData.Add(feederDesignatorValue, rowData);
             		
                }
        
             
             return dictXlFeederData;
             
		}
		
		
		#endregion Macro Methods
		
		
		#region Macro Helper Methods
				
		
		private double CheckRatingProtection(double rating){
			
			
			return rating;
		}
		
		//Create and add shared parameters to the project
		//file is saved in projects filepath location and appended
		private void SetUpSharedParamFile(Document doc)
		{
			//Autodesk.Revit.ApplicationServices.Application app = new Autodesk.Revit.ApplicationServices.Application();
			UIApplication uiapp = this.Application;
			//**Create Shared Parameter File
			//get this docs filepath	
			string docFilePath = doc.PathName.ToString();
			//create the sp file path
			string spFilePath = docFilePath.Replace(".rvt", "_KP_SharedParams.txt");
			CreateExternalSharedParamFile(spFilePath);
			//open the sp file for refrence
			DefinitionFile defFile = SetAndOpenExternalSharedParamFile(uiapp.Application, spFilePath);
			ShowDefinitionFileInfo(defFile); //display to user

		}
		
		//Creating a shared parameter file
		private void CreateExternalSharedParamFile(string sharedParameterFile)
			{
			if (!System.IO.File.Exists(sharedParameterFile)) {
					
				System.IO.FileStream fileStream = System.IO.File.Create(sharedParameterFile);
        		fileStream.Close();
				}
			}
		
		//set current sp file to project
		private DefinitionFile SetAndOpenExternalSharedParamFile(Autodesk.Revit.ApplicationServices.Application application, string sharedParameterFile)
			{
    			// set the path of shared parameter file to current Revit
    			application.SharedParametersFilename = sharedParameterFile;
   				// open the file
    			return application.OpenSharedParameterFile();
			}
		
		//Add sp Group
		public static DefinitionGroup GetOrCreateSharedParamsGroup(DefinitionFile sharedParametersFile,string groupName)
    		{
      		DefinitionGroup g = sharedParametersFile.Groups.get_Item(groupName);
      		if (null == g)
      		{
        		try
        		{
          		g = sharedParametersFile.Groups.Create(groupName);
        		}
        		catch (Exception)
        		{
         		g = null;
        		}
      		}
      		return g;
    		}
		
		//add sp parameter definition
		public static Definition GetOrCreateSharedParamsDefinition(DefinitionGroup defGroup,ParameterType defType,string defName)
    	{
      		Definition definition = defGroup.Definitions.get_Item(defName);
      		if (null == definition)
      		{
        		try
        		{
        			// Create a type definition
        			ExternalDefinitionCreationOptions option = new ExternalDefinitionCreationOptions(defName, defType);
        			definition = defGroup.Definitions.Create(option);
        			
        		
        		}
        		catch (Exception)
        		{
          		definition = null;
        		}
      		}
      		return definition;
		}
		
		//show sp file info
		private void ShowDefinitionFileInfo(DefinitionFile myDefinitionFile)
			{
    			StringBuilder fileInformation = new StringBuilder(500);

    			// get the file name 
   				fileInformation.AppendLine("File Name: " + myDefinitionFile.Filename);
   				fileInformation.AppendLine("");

    			// iterate the Definition groups of this file
    			foreach (DefinitionGroup myGroup in myDefinitionFile.Groups)
    			{
        			// get the group name
        			fileInformation.AppendLine("Group Name: " + myGroup.Name);
        			

        			// iterate the difinitions
        			foreach (Definition definition in myGroup.Definitions)
        			{
            			// get definition name
            			fileInformation.AppendLine("Definition Name: " + definition.Name);
        			}
        			fileInformation.AppendLine("");
    			}
    			TaskDialog.Show("Current Shared Parameter file info with",fileInformation.ToString());
			}
		
		#endregion Macro Helper Methods
		
	}
}