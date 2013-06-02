'Description: This script is created to generate a copy of FFE_Health_Meter.xlsx and 
'		place the name of the latest file in the environment variable for other scripts to use
'Date: 5/30/2013
'*******************************************************************************************************************************************
'								MAIN SCRIPT
'*******************************************************************************************************************************************
strEnvironmentFilePath = "C:\SkyDrive\QTP\Product Family - FFE\Data\FFE_EnvVariables_Main.xml"
environment.LoadFromFile(strEnvironmentFilePath)
fn_FFM_createHealthMeterSpreadsheet environment("healthMeterDataSheetPath")
fn_FFM_setEnvironmentVariable strEnvironmentFilePath, "healthMeterCurrentDataSheetPath", Environment("latestHealthMeterResults")
'New comment here
'*******************************************************************************************************************************************
				                                           ' FUNCTIONS
'*******************************************************************************************************************************************
' Function Name:   fn_FFM_createHealthMeterSpreadsheet()
' Description: 'Function to create the copy of the main spreadsheet
' Output Parameters:
' Date Created: 
' Usage sample:
' Updates:
Function fn_FFM_createHealthMeterSpreadsheet(strExcelWorkbookPath)
Dim excelApp,excelSheet,cell,objWorkbook1,strLocation
		Set blah
        Set excelApp = CreateObject("Excel.Application")
        excelApp.Visible = True
        excelApp.DisplayAlerts = False
        Set objWorkbook1 = excelApp.Workbooks.Open(strExcelWorkbookPath)


                'CREATE AN UPDATED PATH FOR THE SPREADSHEET, ADD TIMESTAMP
'                strUpdatedHealthMeterPath = Replace(strExcelWorkbookPath,".xlsx","",1,1,1)
'                strUpdatedHealthMeterPath = Replace(strUpdatedHealthMeterPath,".xls","",1,1,1)
'                strUpdatedHealthMeterPath = strUpdatedHealthMeterPath & fn_FFE_generateDate & ".xlsx"
	strNewSheetLocation = "C:\healthMeter_" & fn_FFE_generateDate & ".xlsx"
                excelApp.ActiveWorkbook.SaveAs strNewSheetLocation
                Environment("latestHealthMeterResults") = strNewSheetLocation
            excelAPp.Application.Quit
	Set objWorkbook1 = Nothing
	Set excelApp = Nothing
End Function

'*******************************************************************************************************************************************

' Function Name:   fn_FFM_setEnvironmentVariable()
' Description: 'Function to set enviroment variable
' Output Parameters:
' Date Created: 
' Usage sample:
' Updates:
Public Function fn_FFM_setEnvironmentVariable(ByRef strXMLFilePath, strXMLAttributeName, strXMLValue)
Dim i, xmlDoc, nodes, strTempAttribute, intCounter
		
Set xmlDoc = CreateObject("Microsoft.XMLDOM")
 xmlDoc.Load(strXMLFilePath)
Set nodes = xmlDoc.selectNodes("//*")
         		
intCounter = 0			
			For i = 0 to nodes.length
				
				On error resume next
					strTempAttribute = nodes(i).nodeTypedValue
					strBaseName = nodes(i).baseName
				On error goto 0
				err.clear
				If strTempAttribute = strXMLAttributeName and strBaseName = "Name" Then	
					' Enter value for first match only. Remaining matches are for error reporting only
					nodes(i+1).nodeTypedValue = strXMLValue		
					intCounter 	= intCounter + 1
					Exit For
				End If
			Next

	' Report as error if any of the attributes appear < 1 times
	If intCounter < 1 Then
		Print  "Environment Variable " & strXMLAttributeName & " not found"
	End If
xmlDoc.Save strXMLFilePath
End Function


'*******************************************************************************************************************************************

' Function Name:   fn_FFE_generateDate()
' Description: 
' Output Parameters:
' Date Created: 
' Usage sample:
' Updates:
Public Function fn_FFE_generateDate()	'added
   sYear= DatePart("yyyy",now)
   sMonth = DatePart("m", now)
   sDay = DatePart("d", now)
   strHour = DatePart("h",now)
   strMin = DatePart("n",now)
   strSec = DatePart("s",now)
   fn_FFE_generateDate = " " & sMonth & "_" & sDay & "_"  & sYear&"_" & strHour & "h" & strMin & "m"  & strSec& "s"
End Function

