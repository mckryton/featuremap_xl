Attribute VB_Name = "basFeatureReader"
'------------------------------------------------------------------------
' Description  : this module is about reading files into a data model
'------------------------------------------------------------------------

' Copyright 2016 Matthias Carell
'
'   Licensed under the Apache License, Version 2.0 (the "License");
'   you may not use this file except in compliance with the License.
'   You may obtain a copy of the License at
'
'       http://www.apache.org/licenses/LICENSE-2.0
'
'   Unless required by applicable law or agreed to in writing, software
'   distributed under the License is distributed on an "AS IS" BASIS,
'   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
'   See the License for the specific language governing permissions and
'   limitations under the License.

'Declarations

'Declare variables

'Options
Option Explicit
'-------------------------------------------------------------
' Description   : ask user for the folder containing the feature files
' Parameter     :
' Returnvalue   : feature files directory as string
'-------------------------------------------------------------
Public Function getFeatureFilesDir() As String

    Dim strFeatureDir As Variant
    Dim AppleScript As String
    
    #If Mac Then
    
    #Else
        Dim dlgChooseFolder As FileDialog
    #End If

    On Error GoTo error_handler
    #If Mac Then
        'TODO: fix known bug -> umlauts like Š get converted by vba into a_
        AppleScript = "(choose folder with prompt ""choose feature folder"" default location (path to the desktop folder from user domain)) as string"
        strFeatureDir = MacScript(AppleScript)
    #Else
        Set dlgChooseFolder = Application.FileDialog(msoFileDialogFolderPicker)
        With dlgChooseFolder
            .Title = "Please choose a feature folder"
            .AllowMultiSelect = False
            '.InitialFileName = strPath
            If .Show <> False Then
                strFeatureDir = .SelectedItems(1) & "\"
            End If
        End With
        Set dlgChooseFolder = Nothing
    #End If
    basSystem.log ("feature dir is set to " & strFeatureDir)
    getFeatureFilesDir = strFeatureDir
    Exit Function
    
error_handler:
    basSystem.log_error "basFeatureReader.getFeatureFilesDir"
End Function

'-------------------------------------------------------------
' Description   : find and open all feature files and create a
'                   domain model from extracted data
' Parameter     : pstrFeatureDir    - directory containing all feature files
' Returnvalue   : domain model as collection
'-------------------------------------------------------------
Public Function setupDataModel(pstrFeatureDir As String) As Collection

    Dim colDomainModel As New Collection
    Dim lngFeatureId As Long
    Dim arrFeatureFileNames As Variant
    Dim lngFeatureFileIndex As Long
    Dim arrFileName As Variant
    Dim lngFeatureFileId As Long
    Dim colFeature As Collection
    
    On Error GoTo error_handler
    lngFeatureId = 1
    arrFeatureFileNames = getFeatureFileNames(pstrFeatureDir)
    For lngFeatureFileIndex = 0 To UBound(arrFeatureFileNames)
        'extract feature id from file name (assuming that the file is named <feature id>-<feature name>.feature)
        arrFileName = Split(arrFeatureFileNames(lngFeatureFileIndex), "-")
        If IsNumeric(arrFileName(0)) Then
            lngFeatureFileId = Val(arrFileName(0))
        Else
            lngFeatureFileId = -1
        End If
        Application.StatusBar = "read feature file " & arrFeatureFileNames(lngFeatureFileIndex)
        Set colFeature = readDataFromFeatureFile(pstrFeatureDir & arrFeatureFileNames(lngFeatureFileIndex))
        
'        set vDomainName to domain of vFeature
'        set vAggregateName to aggregate of vFeature
'        -- have to use counters because referencing into the strucure of vDomainmodel seems not to be possible
'        set vDomainCount to 0
'        set vAggregateCount to 0
'        -- domains of vDomainModel is a list of records where each record defines a domain
'        -- now try to figure out out if a record for the given domain already exists
'        set vIsNewItem to true

    
    
    
    
    
    
    
    Next
    
    Exit Function
    
error_handler:
    basSystem.log_error "basFeatureReader.setupDataModel"
End Function
'-------------------------------------------------------------
' Description   : find all feature files
' Parameter     : pstrFeatureDir    - directory containing all feature files
' Returnvalue   : list of feature file names as array
'-------------------------------------------------------------
Private Function getFeatureFileNames(pstrFeatureDir As String) As Variant

    Dim colFeatureFileNames As Collection
    'Applescript code for Mac version
    Dim strScript As String
    Dim varFeatureFiles As Variant
    
    On Error GoTo error_handler
    #If Mac Then
        strScript = "set vFeatureFileNames to {}" & vbLf & _
                    "tell application ""Finder""" & vbLf & _
                        "set vFeaturesFolder to """ & pstrFeatureDir & """ as alias" & vbLf & _
                        "set vFeatureFiles to (get files of vFeaturesFolder whose name ends with "".feature"")" & vbLf & _
                        "repeat with vFeatureFile in vFeatureFiles" & vbLf & _
                                "set end of vFeatureFileNames to get name of vFeatureFile" & vbLf & _
                        "end repeat" & vbLf & _
                    "end tell" & vbLf & _
                    "return vFeatureFileNames"
        varFeatureFiles = MacScript(strScript)
        varFeatureFiles = Split(varFeatureFiles, ",")
    #Else
    
    #End If
    
    basSystem.log "found " & UBound(varFeatureFiles) + 1 & " .feature files"
    getFeatureFileNames = varFeatureFiles
    Exit Function
    
error_handler:
    basSystem.log_error "basFeaturemap.getFeatureFileNames"
End Function

'-------------------------------------------------------------
' Description   : read data from a feature file
' Parameter     : pstrFeatureFile    - full filename of the feature file
' Returnvalue   : collection containing the feature file data
'-------------------------------------------------------------
Private Function readDataFromFeatureFile(ByVal pstrFeatureFile As String) As Collection

    Dim strDomain As String
    Dim strAggregate As String
    Dim strTagFeatureStatus As String
    Dim strTagScenarioStatus As String
    Dim colScenarios As New Collection
    Dim strScenarioName As String
    Dim lngLineNumber As Long
    Dim strScript As String
    Dim varFileText As Variant
    Dim strParagraph As String
    
    
    On Error GoTo error_handler
    strDomain = "undefined"
    strAggregate = "undefined"
    strTagFeatureStatus = ""
    strTagScenarioStatus = ""
    lngLineNumber = 0
    
    'read lines of the feature file into an array
    #If Mac Then
        strScript = "set AppleScript's text item delimiters to ""#@#@""" & vbLf & _
                    "return (paragraphs of (read (""" & pstrFeatureFile & """ as alias) as Çclass utf8È)) as string"
        varFileText = Split(MacScript(strScript), "#@#@")
    #Else
    
    #End If
    'look for the feature
    While lngLineNumber <= UBound(varFileText)
        strParagraph = varFileText(lngLineNumber)
        lngLineNumber = lngLineNumber + 1
'        -- look for a domain tag
'        set AppleScript's text item delimiters to cDomainTag
'        if (count text items of text of vParagraph) > 1 then
'            set vDomain to first word of text item 2 of text of vParagraph
'        End If
'        -- look for a status tag
'        set AppleScript's text item delimiters to cStatusTag
'        if (count text items of text of vParagraph) > 1 then
'            set vTagFeatureStatus to first word of text item 2 of text of vParagraph
'        End If
'        -- look for the feature name
'        set AppleScript's text item delimiters to ": "
'        if first word of vParagraph = "Feature" then
'            set vFeature to text item 2 of vParagraph
'            -- try to get the aggregate name, assum the naming is using this scheme <aggregate name> - <feature name>
'            set AppleScript's text item delimiters to " - "
'            if cDisableAggregates is false and (count text items of vFeature) = 2 then
'                set vAggregate to text item 1 of vFeature
'                set vFeature to text item 2 of vFeature
'            End If
'            exit repeat
'        End If
'    end repeat
'
'    -- look for scenarios
'    repeat with vParagraph in (get items (vLineNumber + 1) thru (length of vData) of vData)
'        -- look for a status tag above the scenario name
'        set AppleScript's text item delimiters to cStatusTag
'        if (count text items of text of vParagraph) > 1 then
'            set vTagScenarioStatus to first word of text item 2 of text of vParagraph
'        End If
'        -- look for the scenarios
'        set AppleScript's text item delimiters to ": "
'        if (count words of text of vParagraph) > 0 and first word of vParagraph = "Scenario" then
'            set vScenarioName to text item 2 of vParagraph
'        End If
'        If vScenarioName Is Not Null Then
'            set end of vScenarios to {name:vScenarioName, tags:{status:vTagScenarioStatus}}
'            set vScenarioName to null
'            set vTagScenarioStatus to null
'        End If
'    end repeat
'
'    set vProcessedData to {domain:vDomain, aggregate:vAggregate, feature:vFeature, scenarios:vScenarios, tags:{status:vTagFeatureStatus}}
'    set AppleScript's text item delimiters to vOldDelimiters
'    --return list of records from text file
'    return vProcessedData
    
    Wend
    
    Exit Function
    
error_handler:
    basSystem.log_error "basFeatureReader.readDataFromFeatureFile"
End Function

