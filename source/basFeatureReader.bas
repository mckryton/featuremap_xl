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

    Dim strFeatureDirInfo As Variant
    Dim strFeatureDirDisk As String
    Dim strFeatureDirPath As String
    Dim strFeatureDirFullPath As String
    Dim strAppleScript As String
    
    #If Mac Then
    
    #Else
        Dim dlgChooseFolder As FileDialog
    #End If

    On Error GoTo error_handler
    #If Mac Then
        'TODO: fix known bug -> umlauts like Š get converted by vba into a_
        strAppleScript = "try" & vbLf & _
                                "tell application ""Finder""" & vbLf & _
                                    "set vPath to (choose folder with prompt ""choose feature folder"" default location (path to the desktop folder from user domain))" & vbLf & _
                                    "return {url of vPath, displayed name of disk of vPath}" & vbLf & _
                                "end tell" & vbLf & _
                            "on error" & vbLf & _
                                "return """"" & vbLf & _
                            "end try"
        strFeatureDirInfo = Split(MacScript(strAppleScript), ", ")
        strFeatureDirPath = basSystem.decomposeUrlPath(strFeatureDirInfo(0))
        strFeatureDirDisk = strFeatureDirInfo(1)
        strFeatureDirFullPath = strFeatureDirDisk & strFeatureDirPath
    #Else
        Set dlgChooseFolder = Application.FileDialog(msoFileDialogFolderPicker)
        With dlgChooseFolder
            .Title = "Please choose a feature folder"
            .AllowMultiSelect = False
            '.InitialFileName = strPath
            If .Show <> False Then
                strFeatureDirFullPath = .SelectedItems(1) & "\"
            End If
        End With
        Set dlgChooseFolder = Nothing
    #End If
    basSystem.log ("feature dir is set to " & strFeatureDirFullPath)
    getFeatureFilesDir = strFeatureDirFullPath
    Exit Function
    
error_handler:
    basSystem.log_error "basFeatureReader.getFeatureFilesDir"
End Function

'-------------------------------------------------------------
' Description   : find and open all feature files and create a
'                   domain model from extracted data
' Parameter     : pstrFeatureNameDir    - directory containing all feature files
' Returnvalue   : domain model as collection
'-------------------------------------------------------------
Public Function setupDataModel(pstrFeatureNameDir As String) As Collection

    Dim colDomainModel As New Collection
    Dim colDomain As Collection
    Dim colDomains As Collection
    Dim colAggregate As Collection
    Dim colAggregates As Collection
    Dim lngFeatureId As Long
    Dim arrFeatureFileNames As Variant
    Dim lngFeatureFileIndex As Long
    Dim arrFileName As Variant
    Dim lngFeatureFileId As Long
    Dim colFeature As Collection
    
    On Error GoTo error_handler
    colDomainModel.Add New Collection, "domains"
    lngFeatureId = 1
    arrFeatureFileNames = getFeatureFileNames(pstrFeatureNameDir)
    For lngFeatureFileIndex = 0 To UBound(arrFeatureFileNames)
        'extract feature id from file name (assuming that the file is named <feature id>-<feature name>.feature)
        arrFileName = Split(arrFeatureFileNames(lngFeatureFileIndex), "-")
        If IsNumeric(arrFileName(0)) Then
            lngFeatureFileId = Val(arrFileName(0))
        Else
            lngFeatureFileId = -1
        End If
        Application.StatusBar = "read feature file " & arrFeatureFileNames(lngFeatureFileIndex)
        Set colFeature = readFeatureFile(pstrFeatureNameDir & arrFeatureFileNames(lngFeatureFileIndex))
        On Error GoTo create_new_domain
        Set colDomain = colDomainModel("domains")(colFeature("domain"))
        On Error GoTo error_handler
        On Error GoTo create_new_aggregate
        Set colAggregate = colDomain("aggregates")(colFeature("aggregate"))
        On Error GoTo error_handler
        colFeature.Add lngFeatureId, "id"
        colFeature.Add lngFeatureId, "fileId"
        colAggregate("features").Add colFeature, colFeature("name")
        basSystem.log "feature " & colFeature("name") & " added to the model"
    Next
    Set setupDataModel = colDomainModel
    Exit Function
create_new_domain:
    Set colDomain = New Collection
    colDomain.Add colFeature("domain"), "name"
    colDomain.Add New Collection, "aggregates"
    colDomainModel("domains").Add colDomain, colFeature("domain")
    basSystem.log "domain " & colFeature("domain") & " added to the model"
    Resume Next
create_new_aggregate:
    Set colAggregate = New Collection
    colAggregate.Add colFeature("aggregate"), "name"
    colAggregate.Add New Collection, "features"
    colDomain("aggregates").Add colAggregate, colFeature("aggregate")
    basSystem.log "aggregate " & colFeature("aggregate") & " added to the model"
    Resume Next
error_handler:
    basSystem.log_error "basFeatureReader.setupDataModel"
End Function
'-------------------------------------------------------------
' Description   : find all feature files
' Parameter     : pstrFeatureNameDir    - directory containing all feature files
' Returnvalue   : list of feature file names as array
'-------------------------------------------------------------
Private Function getFeatureFileNames(pstrFeatureNameDir As String) As Variant

    Dim colFeatureFileNames As Collection
    'Applescript code for Mac version
    Dim strScript As String
    Dim varFeatureFiles As Variant
    Dim varFeatureFilePath As Variant
    Dim lngFeatureFileIndex As Long
    
    On Error GoTo error_handler
    Application.StatusBar = "retrieve feature file names"
    #If Mac Then
        strScript = "set vFeatureFileNames to {}" & vbLf & _
                    "tell application ""Finder""" & vbLf & _
                        "set vFeaturesFolder to """ & pstrFeatureNameDir & """ as alias" & vbLf & _
                        "set vFeatureFiles to (get files of vFeaturesFolder whose name ends with "".feature"")" & vbLf & _
                        "repeat with vFeatureFile in vFeatureFiles" & vbLf & _
                                "set end of vFeatureFileNames to get URL of vFeatureFile" & vbLf & _
                        "end repeat" & vbLf & _
                    "end tell" & vbLf & _
                    "return vFeatureFileNames"
        varFeatureFiles = MacScript(strScript)
        varFeatureFiles = Split(varFeatureFiles, ", ")
        For lngFeatureFileIndex = 0 To UBound(varFeatureFiles)
            varFeatureFilePath = Split(varFeatureFiles(lngFeatureFileIndex), "/")
            varFeatureFiles(lngFeatureFileIndex) = basSystem.decomposeUrlPath("file://" & _
                                                        varFeatureFilePath(UBound(varFeatureFilePath)))
        Next
    #Else
        Dim fsoFileSystem As Variant
        Dim folFeatures As Variant
        Dim filFeature As Variant
        Dim strFiles As String
        
        strFiles = ""
        Set fsoFileSystem = CreateObject("Scripting.FileSystemObject")
        Set folFeatures = fsoFileSystem.GetFolder(pstrFeatureNameDir)
        For Each filFeature In folFeatures.Files
            If Trim(LCase(Right(filFeature.Name, Len(".feature")))) = ".feature" Then
                strFiles = strFiles & filFeature.Name & "//"
            End If
        Next
        strFiles = Left(strFiles, Len(strFiles) - 2)
        varFeatureFiles = Split(strFiles, "//")
    #End If
    
    basSystem.log "found " & UBound(varFeatureFiles) + 1 & " .feature files"
    getFeatureFileNames = varFeatureFiles
    Exit Function
    
error_handler:
    basSystem.log_error "basFeaturemap.getFeatureFileNames"
End Function
'-------------------------------------------------------------
' Description   : read data from a feature file
' Parameter     : pstrFeatureNameFile    - full filename of the feature file
' Returnvalue   : collection containing the feature file data
'-------------------------------------------------------------
Private Function readFeatureFile(ByVal pstrFeatureNameFile As String) As Collection

    Dim strDomainName As String
    Dim strAggregateName As String
    Dim strFeatureName As String
    Dim varFeatureNameParts As Variant            'split feature name into an array if it contains the aggregate
    Dim colFeature As New Collection
    Dim colScenarios As New Collection
    Dim colScenario As Collection
    Dim strScenarioName As String
    Dim lngLineNumber As Long
    Dim strScript As String
    Dim varFileText As Variant
    Dim strParagraph As String
    Dim colFeatureTags As New Collection
    Dim colScenarioTags As Collection
    
    On Error GoTo error_handler
    strDomainName = "undefined"
    lngLineNumber = 0
    strFeatureName = "undefined"
    strAggregateName = "undefined"
    
    basSystem.log "read data from feature file " & pstrFeatureNameFile
    'read lines of the feature file into an array
    #If Mac Then
        strScript = "set AppleScript's text item delimiters to ""#@#@""" & vbLf & _
                    "return (paragraphs of (read (""" & pstrFeatureNameFile & """ as alias) as Çclass utf8È)) as string"
        varFileText = Split(MacScript(strScript), "#@#@")
    #Else
        Const cForReading = 1, cForWriting = 2
   
        Dim fsoFileSystem As Variant
        Dim ftxFeatureFile As Variant
        
        Set fsoFileSystem = CreateObject("Scripting.FileSystemObject")
        Set ftxFeatureFile = fsoFileSystem.OpenTextFile(pstrFeatureNameFile, cForReading, True)
        varFileText = Split(ftxFeatureFile.readall, vbLf)
        ftxFeatureFile.Close
    #End If
    'read all the lines above "Feature:"
    Do While lngLineNumber <= UBound(varFileText)
        strParagraph = varFileText(lngLineNumber)
        'found the feature name?
        If InStr(LCase(strParagraph), "feature:") > 0 Then
            strFeatureName = Right(strParagraph, Len(strParagraph) - InStr(LCase(strParagraph), "feature:") - 8)
            If cblnGetAggregatesFromFeatureName Then
                varFeatureNameParts = Split(strFeatureName, " - ")
                If UBound(varFeatureNameParts) > 0 Then
                    strAggregateName = varFeatureNameParts(0)
                    strFeatureName = Right(strFeatureName, Len(strFeatureName) - Len(strAggregateName) - 3)
                End If
                basSystem.log "found feature >" & strFeatureName & "< for aggregate >" & strAggregateName & "<"
            Else
            basSystem.log "found feature " & strFeatureName
            End If
            Exit Do
        End If
        findTags colFeatureTags, strParagraph
        lngLineNumber = lngLineNumber + 1
    Loop
    colFeature.Add strFeatureName, "name"
    colFeature.Add colFeatureTags, "tags"
    
    'look for scenarios
    Set colScenario = New Collection
    Set colScenarioTags = New Collection
    While lngLineNumber <= UBound(varFileText)
        strParagraph = varFileText(lngLineNumber)
        'found a scenario name?
        If InStr(LCase(strParagraph), "scenario:") > 0 Then
            strScenarioName = Right(strParagraph, Len(strParagraph) - InStr(LCase(strParagraph), "scenario:") - 9)
            colScenario.Add strScenarioName, "name"
            colScenario.Add colScenarioTags, "tags"
            On Error GoTo duplicate_scenario_name
            colScenarios.Add colScenario, strScenarioName
            On Error GoTo error_handler
            'get ready for the next scenario
            Set colScenario = New Collection
            Set colScenarioTags = New Collection
        Else
            findTags colScenarioTags, strParagraph
        End If
        lngLineNumber = lngLineNumber + 1
    Wend
    colFeature.Add colScenarios, "scenarios"
    'receive domain from feature tags
    On Error Resume Next
    strDomainName = colFeatureTags(cstrDomainTag)
    On Error GoTo error_handler
    colFeature.Add strDomainName, "domain"
    colFeature.Add strAggregateName, "aggregate"
    'return feature collection as result
    Set readFeatureFile = colFeature
    Exit Function
    
duplicate_scenario_name:
    basSystem.log "duplicate scenario name found: >" & strScenarioName & "<", cLogWarning
    Resume Next
error_handler:
    basSystem.log_error "basFeatureReader.readFeatureFile"
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
