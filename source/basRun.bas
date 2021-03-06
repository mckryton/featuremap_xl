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
    Dim frmOptions As New frmOptionsTemplate
    
    On Error GoTo error_handler
    'warn for unsupported versions of Excel
    #If MAC_OFFICE_VERSION >= 15 Then
        'MsgBox "Excel 2016 MAC is not yet supported, please use Excel 2011 or a Windows version"
        'Exit Sub
    #End If

    'show options
    frmOptions.Show
    If frmOptions.FormCanceled Then
        Exit Sub
    End If
    
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
    basModelVisualizer.visualizeModel wshDrawing, colDomainModel, frmOptions.DrawingOptions
    
    'set height of every domain box to max height
    basModelVisualizer.levelDomainHeight wshDrawing
    
    Application.StatusBar = False
    Exit Sub
    
error_handler:
    basSystem.log_error "basFeaturemap.runFeatureMap"
End Sub

