Attribute VB_Name = "basFeaturemap"
'------------------------------------------------------------------------
' Description  : exports features into feature files
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
' Description   : main routine for generating a new feature map
' Parameter     :
' Returnvalue   :
'-------------------------------------------------------------
Public Sub run_createFeatureMap()

    'local vFeaturesFolder
    Dim strFeatureDir As String
    'local vDomainModel
    Dim colDomainModel As Collection
    'local vDrawingDoc
    Dim wbkFeatureMap As Workbook
    
    On Error GoTo error_handler
    'select a folder containing feature descriptions, text files with a .feature extension
    strFeatureDir = getFeatureFilesDir()
    
    'extract features and scenarios from feature files
'    set vDomainModel to my setupDataModel(vFeaturesFolder)
    Set colDomainModel = setupDataModel(strFeatureDir)
    
'    --create an empty drawing document from OmniGraffle
'    set vDrawingDoc to my createDrawingDoc()
    
'    --draw domain boxes with all aggregates, features and scenarios
'    my assembleModel(vDrawingDoc, vDomainModel)
    
'    --connect each with it's parent
'    my connectItems(vDrawingDoc)
    
'    --set height of every domain box to max height
'    my levelDomainHeight(vDrawingDoc)



    Exit Sub
    
error_handler:
    basSystem.log_error "basFeaturemap.run_createFeatureMap"
End Sub
'-------------------------------------------------------------
' Description   : ask user for the folder containing the feature files
' Parameter     :
' Returnvalue   : feature files directory as string
'-------------------------------------------------------------
Private Function getFeatureFilesDir() As String

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
    basSystem.log_error "basFeaturemap.getFeatureFilesDir"
End Function

'-------------------------------------------------------------
' Description   : find and open all feature files and create a
'                   domain model from extracted data
' Parameter     : pstrFeatureDir    - directory containing all feature files
' Returnvalue   : domain model as collection
'-------------------------------------------------------------
Private Function setupDataModel(pstrFeatureDir As String) As Collection

    Dim colDomainModel As New Collection
    Dim lngFeatureId As Long
    Dim colFeatureFileNames As Collection
    
    On Error GoTo error_handler
    lngFeatureId = 1
    Set colFeatureFileNames = getFeatureFileNames(pstrFeatureDir)
    
    Exit Function
    
error_handler:
    basSystem.log_error "basFeaturemap.setupDataModel"
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
    Set getFeatureFileNames = varFeatureFiles
    Exit Function
    
error_handler:
    basSystem.log_error "basFeaturemap.setupDataModel"
End Function
