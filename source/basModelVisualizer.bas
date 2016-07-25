Attribute VB_Name = "basModelVisualizer"
'------------------------------------------------------------------------
' Description  : this module is about turning a data model into a graphic
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
' Description   : create a new empty workbook for drawing
' Parameter     :
' Returnvalue   : xl workbook
'-------------------------------------------------------------
Public Function createDrawingDoc() As Workbook

    Dim wbkDrawing As Workbook
    Dim wshDrawing As Worksheet

    On Error GoTo error_handler
    basSystem.log "create a new workbook for drawing"
    Set wbkDrawing = Application.Workbooks.Add
    Application.DisplayAlerts = False
    'remove unnecessary sheets
    While wbkDrawing.Worksheets.Count > 1
        wbkDrawing.Worksheets(1).Delete
    Wend
    'hide gridlines
    Set wshDrawing = wbkDrawing.Worksheets(1)
    wshDrawing.Name = "domain model"
    wbkDrawing.Windows(1).DisplayGridlines = False
    
    Application.DisplayAlerts = True
    Set createDrawingDoc = wbkDrawing
    Exit Function
    
error_handler:
    basSystem.log_error "basModelVisualizer.createDrawingDoc"
End Function
