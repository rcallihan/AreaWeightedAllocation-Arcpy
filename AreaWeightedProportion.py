# ---------------------------------------------------------------------------
# AreaWeightedProportion.py
# Description: Create proportional values of target features within source features.
# Author: Ryan Callihan; modified by Alex Brasch
# ---------------------------------------------------------------------------

# Import parameters
import arcpy, os

# Inputs from TBX/User
TargetDataset = arcpy.GetParameterAsText(0)
TargetNameField = arcpy.GetParameterAsText(1)
TargetValueField = arcpy.GetParameterAsText(2)
SourceDataset = arcpy.GetParameterAsText(3)
SourceNameField = arcpy.GetParameterAsText(4)
OutputType = arcpy.GetParameterAsText(5)
OutputDir = arcpy.GetParameterAsText(6)

#workspace
arcpy.env.workspace = "in_memory"
arcpy.env.overwriteOutput = True

#local variables
if OutputType == "dBASE":
	OutFileName = os.path.basename(TargetDataset) + "_by_" + os.path.basename(SourceDataset) + "_AreaWeightedProportion"
	OutFileName_Calculated = os.path.basename(TargetDataset) + "_by_" + os.path.basename(SourceDataset) + "_AreaWeightedProportion_Calculated"

elif OutputType == "Excel (.xls)":
	OutFileName = os.path.basename(TargetDataset) + "_by_" + os.path.basename(SourceDataset) + "_AreaWeightedProportion.xls"
	OutFileFullPath = OutputDir + "//" + OutFileName

	OutFileName_Calculated = os.path.basename(TargetDataset) + "_by_" + os.path.basename(SourceDataset) + "_AreaWeightedProportion_Calculated.xls"
	OutFileFullPath_Calculated = OutputDir + "//" + OutFileName_Calculated

OutPath = OutputDir
TargetSelection = TargetNameField + " not in('')"
SourceSelection = SourceNameField + " not in('')"
inMemoryUnionedFC = "inMemoryUnionedFC"
inMemoryUnionedCleanFC = "inMemoryUnionedCleanFC"
inMemoryUnionedCleanAgainFC = "inMemoryUnionedAgainCleanFC"
inMemoryStatsTable = "inMemoryStatsTable"
inMemoryStatsStatsTable = "inMemoryStatsStatsTable"
inMemoryTarget = "inMemoryTarget"
AWPValueField = "AWP_Value"
ProportionField = "AW_Prop"
ValueField = "MEAN_" + TargetValueField
StatsOutValueField = "SUM_" + AWPValueField
OutputAWPField = "AWP_" + TargetValueField

# Add Geometry Attributes
arcpy.AddMessage("Adding Geometry")
arcpy.CopyFeatures_management(TargetDataset, inMemoryTarget)
print "Adding Geometry"
arcpy.AddGeometryAttributes_management(inMemoryTarget, "AREA", "FEET_US", "ACRES", "")
arcpy.AlterField_management(inMemoryTarget, "POLY_AREA", "Target_Area", "", "DOUBLE", "8", "NULLABLE", "false")

# Union
arcpy.AddMessage("Starting Union")
print "Starting Union"
arcpy.Union_analysis([SourceDataset, inMemoryTarget], inMemoryUnionedFC, "ALL", "", "GAPS")

# Add Geometry Attributes
arcpy.AddMessage("Adding Geometry")
print "Adding Geometry"
arcpy.AddGeometryAttributes_management(inMemoryUnionedFC, "AREA", "FEET_US", "ACRES", "")

# Alter Field
arcpy.AlterField_management(inMemoryUnionedFC, "POLY_AREA", "AreaInSource", "", "DOUBLE", "8", "NULLABLE", "false")

# Select only areas that intersected with Target.
arcpy.AddMessage("Making Selection")
print "Making Selection"
arcpy.Select_analysis(inMemoryUnionedFC, inMemoryUnionedCleanFC, TargetSelection)

# Select ony areas that intersected with Source. (Removed)
# arcpy.AddMessage("Making Selection")
# print "Making Selection"
# arcpy.Select_analysis(inMemoryUnionedFC, inMemoryUnionedCleanAgainFC, SourceSelection)

# Group By Summary - Summarize on Source and on Target
arcpy.AddMessage("Summary Statistics")
print "Summary Statistics"
if not TargetValueField:
	# Output table with proportion (no AWP values calculated, just proportions)
	arcpy.Statistics_analysis(inMemoryUnionedCleanFC,inMemoryStatsTable, [["Target_Area","MEAN"],["AreaInSource","SUM"]],[SourceNameField,TargetNameField])
else:
	# if a target value was defined by a user, pass it through the Statistic_analysis
	arcpy.Statistics_analysis(inMemoryUnionedCleanFC,inMemoryStatsTable, [["Target_Area","MEAN"],["AreaInSource","SUM"],[TargetValueField, "MEAN"]],[SourceNameField,TargetNameField])

# Add field for proportion and calculate
arcpy.AddMessage("Calculating Proportions")
print "Calculating Proportions"
arcpy.AddField_management(inMemoryStatsTable, "AW_Prop", "DOUBLE")
arcpy.CalculateField_management(inMemoryStatsTable, "AW_Prop", "!SUM_AreaInSource! / !MEAN_Target_Area!", "PYTHON")

def outputProportions(intable, outtype, outpath, outfilename, outfullpath): 
	if outtype == "dBASE":
		# Process: Table To Table
		arcpy.AddMessage("Exporting to %s DBASE located to directory %s" % (outfilename, outpath))
		print "Exporting to dBASE"
		arcpy.TableToTable_conversion(intable, outpath, outfilename)
	elif outtype == "Excel (.xls)":	
		# Process: Table To Excel
		arcpy.AddMessage("Exporting to %s Excel to directory %s" % (outfilename, outpath))
		arcpy.TableToExcel_conversion(intable, outfullpath, "NAME", "CODE")

# If a target field exists, calculate the area weighted porportion for that value. 
if not TargetValueField:
	# Output table with proportion (no AWP values calculated, just proportions)
	arcpy.AddMessage("Exporting proportions table.")
	outputProportions(inMemoryStatsTable, OutputType,OutPath, OutFileName, OutFileFullPath)

else:
	# Multiply the value field from the target dataset by the proportion field. 
	Arcpy.AddMessage("Calculating area weighted proportion values.")
	arcpy.AddField_management(inMemoryStatsTable, "AWP_Value", "DOUBLE")
	AWPcalc = "!" + ValueField + "! * !" + ProportionField + "!"
	arcpy.CalculateField_management(inMemoryStatsTable, "AWP_Value" , AWPcalc, "PYTHON")

	#output intermediate proportion file
	arcpy.AddMessage("Exporting proportions table.")
	outputProportions(inMemoryStatsTable, OutputType,OutPath, OutFileName, OutFileFullPath)
	
	# Summarize on Source ID
	arcpy.Statistics_analysis(inMemoryStatsTable, inMemoryStatsStatsTable, [["AWP_Value", "SUM"]],SourceNameField)
	arcpy.AlterField_management(inMemoryStatsStatsTable, StatsOutValueField, OutputAWPField, "", "DOUBLE", "8", "NULLABLE", "false")
	arcpy.AlterField_management(inMemoryStatsStatsTable, "Frequency", "Num_Features")

	# Output Table with calculated AWP values
	arcpy.AddMessage("Exporting calculated proportions table.")
	outputProportions(inMemoryStatsStatsTable, OutputType,OutPath, OutFileName_Calculated, OutFileFullPath_Calculated)

#delete all intermediate files
arcpy.Delete_management("in_memory")




