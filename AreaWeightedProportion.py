# ---------------------------------------------------------------------------
# AreaWeightedProportion.py
# Description: Create proportional values of target features within source features.
# Author: Ryan Callihan; modified by Alex Brasch
# ---------------------------------------------------------------------------

# Import parameters
import arcpy, os

# Script arguments
TargetDataset = arcpy.GetParameterAsText(0)
TargetNameField = arcpy.GetParameterAsText(1)
SourceDataset = arcpy.GetParameterAsText(2)
SourceNameField = arcpy.GetParameterAsText(3)
OutputType = arcpy.GetParameterAsText(4)
OutputDir= arcpy.GetParameterAsText(5)

##OutputDir = 
##AreaToProportion = "X:\F1147.01 Wallowa Land Trust\Deliverables\GIS_Data\ConsPlans_PriorityParcels.gdb\Parcels_Private_PriorityAnalysis"
##BoundariesToDistribute = "X:\F1147.01 Wallowa Land Trust\Deliverables\GIS_Data\ConsPlan_GISAnalysisInputData.gdb\PLN_OR_HighValueForest_2007_DSLV"
##BoundaryNameField = "Forestland"
##BGNameField = "GIS_ID"

#workspace
arcpy.env.workspace = "in_memory"
arcpy.env.overwriteOutput = True

#local variables
if OutputType == "dBASE":
	OutFileName = os.path.basename(TargetDataset) + "_by_" + os.path.basename(SourceDataset) + "_AreaWeightedProportion"
elif OutputType == "Excel (.xls)":
	OutFileName = os.path.basename(TargetDataset) + "_by_" + os.path.basename(SourceDataset) + "_AreaWeightedProportion.xls"
	OutFileFullPath = OutputDir + "//" + OutFileName

OutPath = OutputDir
TargetSelection = TargetNameField + " not in('')"
SourceSelection = SourceNameField + " not in('')"
inMemoryUnionedFC = "inMemoryUnionedFC"
inMemoryUnionedCleanFC = "inMemoryUnionedCleanFC"
inMemoryUnionedCleanAgainFC = "inMemoryUnionedAgainCleanFC"
inMemoryStatsTable = "inMemoryStatsTable"
inMemoryTarget = "inMemoryTarget"

# Process: Add Geometry Attributes
arcpy.AddMessage("Adding Geometry")
arcpy.CopyFeatures_management(TargetDataset, inMemoryTarget)
print "Adding Geometry"
arcpy.AddGeometryAttributes_management(inMemoryTarget, "AREA", "FEET_US", "ACRES", "")
arcpy.AlterField_management(inMemoryTarget, "POLY_AREA", "Target_Area", "", "DOUBLE", "8", "NULLABLE", "false")

# Process: Union
arcpy.AddMessage("Starting Union")
print "Starting Union"
arcpy.Union_analysis([SourceDataset, inMemoryTarget], inMemoryUnionedFC, "ALL", "", "GAPS")

# Process: Add Geometry Attributes
arcpy.AddMessage("Adding Geometry")
print "Adding Geometry"
arcpy.AddGeometryAttributes_management(inMemoryUnionedFC, "AREA", "FEET_US", "ACRES", "")

# Process: Alter Field
arcpy.AlterField_management(inMemoryUnionedFC, "POLY_AREA", "AreaInSource", "", "DOUBLE", "8", "NULLABLE", "false")

# Select only areas that intersected with Target.
arcpy.AddMessage("Making Selection")
print "Making Selection"
arcpy.Select_analysis(inMemoryUnionedFC, inMemoryUnionedCleanFC, TargetSelection)

# Select ony areas that intersected with Source.
arcpy.AddMessage("Making Selection")
print "Making Selection"
arcpy.Select_analysis(inMemoryUnionedCleanFC, inMemoryUnionedCleanAgainFC, SourceSelection)

#Group By Summary - Summarize on Source and on Target
arcpy.AddMessage("Summary Statistics")
print "Summary Statistics"
arcpy.Statistics_analysis(inMemoryUnionedCleanAgainFC,inMemoryStatsTable, [["Target_Area","MEAN"],["AreaInSource","SUM"]],[SourceNameField,TargetNameField])

# Add field for proportion and calculate
arcpy.AddMessage("Calculating Proportions")
print "Calculating Proportions"
arcpy.AddField_management(inMemoryStatsTable, "AW_Prop", "DOUBLE")
arcpy.CalculateField_management(inMemoryStatsTable, "AW_Prop", "!SUM_AreaInSource! / !MEAN_Target_Area!", "PYTHON")

if OutputType == "dBASE":
	# Process: Table To Table
	arcpy.AddMessage("Exporting to dBASE")
	print "Exporting to dBASE"
	arcpy.TableToTable_conversion(inMemoryStatsTable, OutPath, OutFileName)
elif OutputType == "Excel (.xls)":	
	# Process: Table To Excel
	arcpy.AddMessage("Exporting to Excel")
	print "Exporting to Excel"
	arcpy.TableToExcel_conversion(inMemoryStatsTable, OutFileFullPath, "NAME", "CODE")
	
arcpy.Delete_management("in_memory")




