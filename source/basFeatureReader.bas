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
    Dim strFeature As String
    Dim varFeature As Variant            'split feature name into an array if it contains the aggregate
    Dim colScenarios As New Collection
    Dim strScenarioName As String
    Dim lngLineNumber As Long
    Dim strScript As String
    Dim varFileText As Variant
    Dim strParagraph As String
    Dim colFeatureTags As New Collection
    Dim colScenarioTags As Collection
    
    On Error GoTo error_handler
    strDomain = "undefined"
    lngLineNumber = 0
    
    basSystem.log "read data from feature file " & pstrFeatureFile
    'read lines of the feature file into an array
    #If Mac Then
        strScript = "set AppleScript's text item delimiters to ""#@#@""" & vbLf & _
                    "return (paragraphs of (read (""" & pstrFeatureFile & """ as alias) as Çclass utf8È)) as string"
        varFileText = Split(MacScript(strScript), "#@#@")
    #Else
        'TODO add windows support
    #End If
    'read all the lines above Feature:
    Do While lngLineNumber <= UBound(varFileText)
        strParagraph = varFileText(lngLineNumber)
        'found feature?
        If InStr(LCase(strParagraph), "feature:") > 0 Then
            strFeature = Right(strParagraph, Len(strParagraph) - InStr(LCase(strParagraph), "feature:") - 8)
            If cblnGetAggregatesFromFeatureName Then
                varFeature = Split(strFeature, " - ")
                If UBound(varFeature) > 0 Then
                    strAggregate = varFeature(0)
                    strFeature = Right(strFeature, Len(strFeature) - Len(strAggregate) - 3)
                Else
                    strAggregate = "undefined"
                End If
            End If
            Exit Do
        End If
        findTags colFeatureTags, strParagraph
        lngLineNumber = lngLineNumber + 1
    Loop
    Stop
    
    'look for scenarios
    While lngLineNumber <= UBound(varFileText)
        Set colScenarioTags = New Collection
        If InStr(strParagraph, "Scenario:") Then
            
    '        If vScenarioName Is Not Null Then
    '            set end of vScenarios to {name:vScenarioName, tags:{status:vTagScenarioStatus}}
    '            set vScenarioName to null
    '            set vTagScenarioStatus to null
    '        End If
        
        Else
            findTags colScenarioTags, strParagraph
        
        End If
        
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
'-------------------------------------------------------------
' Description   : extract any tag from a string and add them to a given collection
' Parameter     : pcolFeatureTags   - where to save tags
'                 pstrParagraph     - a line from a feature file
' Returnvalue   :
'-------------------------------------------------------------
Private Sub findTags(pcolFeatureTags As Collection, pstrParagraph As String)
    
    Dim varPieces As Variant
    Dim lngPieceIndex As Long
    Dim varTag As Variant
    Dim strTag As String
    Dim strTagKey As String
    Dim strTagValue As String
        
    On Error GoTo error_handler
    If Trim(pstrParagraph) <> "" Then
        varPieces = Split(pstrParagraph, " ")
        For lngPieceIndex = 0 To UBound(varPieces)
            'tags are marked with an @ sign
            If Left(varPieces(lngPieceIndex), 1) = "@" Then
                strTag = Right(varPieces(lngPieceIndex), Len(varPieces(lngPieceIndex)) - 1)
                'key-value tags are separated by a - sign
                varTag = Split(strTag, "-")
                If UBound(varTag) = 0 Then
                    'it's a single tag
                    strTagKey = strTag
                    strTagValue = strTag
                Else
                    strTagKey = varTag(0)
                    strTagValue = Right(strTag, Len(strTag) - Len(strTagKey) - 1)
                End If
            On Error GoTo found_duplicate_tag
            pcolFeatureTags.Add strTagValue, strTagKey
            On Error GoTo error_handler
            End If
        Next
    End If
    Exit Sub
    
found_duplicate_tag:
    basSystem.log "found duplicate value for tag " & strTagKey, cLogWarning
    Resume Next
error_handler:
    basSystem.log_error "basFeatureReader.findTags"
End Sub
