Attribute VB_Name = "basRun"
'------------------------------------------------------------------------
' Description  : this module is about to execute the whole application
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
Public Sub runFeatureMap()

    'local vFeaturesFolder
    Dim strFeatureDir As String
    'local vDomainModel
    Dim colDomainModel As Collection
    'local vDrawingDoc
    Dim wshDrawing As Worksheet
    
    On Error GoTo error_handler
    'select a folder containing feature descriptions, text files with a .feature extension
    strFeatureDir = basFeatureReader.getFeatureFilesDir()
    If strFeatureDir = "" Then
        basSystem.log "choose feature folder dialog was canceled"
        Exit Sub
    End If
    
    'extract features and scenarios from feature files
    Set colDomainModel = basFeatureReader.setupDataModel(strFeatureDir)
    
    'create a new workbook as empty drawing document
    Set wshDrawing = createDrawingDoc()
    
    'draw domain boxes with all aggregates, features and scenarios
    basModelVisualizer.visualizeModel wshDrawing, colDomainModel, cblnHideAggregatesDefault
    
'    --connect each with it's parent
'    my connectItems(vDrawingDoc)
    
    'set height of every domain box to max height
    basModelVisualizer.levelDomainHeight wshDrawing
    
    Application.StatusBar = False
    Exit Sub
    
error_handler:
    basSystem.log_error "basFeaturemap.runFeatureMap"
End Sub

