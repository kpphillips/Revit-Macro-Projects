# Revit-Macro-Projects

## 1. Electrical Feeder Managment Macros
[![YouTube Video](https://img.youtube.com/vi/aZcgYRaJrsQ/0.jpg)](https://www.youtube.com/watch?v=aZcgYRaJrsQ)


The current functionality incorporated as separate Macro functions are noted below.  Some of these need to be run in sequential order.  First save the Revit file and Excel file into the same folder. Then run through the commands: Note: The Excel file setup works with the current Macros, ie. The names of the parameters should match in Excel and Revit.  

Notes:
1. Open Revit File -> Select Manage Tab -> Select Macro Manager
2. Run Commands in sequential order
3. Revit User will need to add data to the "FD_Feeder Designator" Parameter that matches Excel before running "F_UpdateAllProjectConduitAndFittingParamsFromExcel"

A_SetUpSharedParameterFile - Creates a blank Shared Parameter file in in the project folder

B_CreateFeederSharedParameters - Adds hard coded parameters to the shared parameter file. 

C_AddProjectSharedParameters - Applies the Shared Parameters to Conduits/Conduit Fittings only.  

D_GetElectricalCircuits - Just a sample Macro to get the circuits in the project.  This can be run alone with no other Macro dependencies. 

E_GetElementsInConduitRunFromSelecedConduit - This is a work in progress and not fully functional yet. 

F_UpdateAllProjectConduitAndFittingParamsFromExcel - Before this Macro is run the user will need to fill out the conduit/conduit fittings “FD_Feeder Designator” that will match that in the Excel file.  If there are typos or noting is filled out, then it will not update the other related Shared Parameters appended with “FD_”.

v2
updated F_UpdateAllProjectConduitAndFittingParamsFromExcel to close Excel file when finished.  

v3 
Updated E_GetElementsInConduitRunFromSelectedConduitAndUpdateParameters
