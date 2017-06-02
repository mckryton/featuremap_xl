VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOptionsTemplate 
   Caption         =   "featuremap options"
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   -14720
   ClientWidth     =   10000
   OleObjectBlob   =   "frmOptionsTemplate.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmOptionsTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------
' Description  : UI for configuring the script
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
Dim mblnFormCanceled As Boolean
Dim mcolDrawingOptions As Collection

'Options
Option Explicit



'-------------------------------------------------------------
' Description   : add a new color rule
' Parameter     :
' Returnvalue   :
'-------------------------------------------------------------
Private Sub cmdAddRule_Click()
    
    Dim frmEdit As New frmEditRule
    Dim strTagName As String
    Dim strTarget As String
    Dim strColor As String
    Dim lngListIndex As Long
    
    On Error GoTo error_handler
    Me.Hide
    frmEdit.IsNewRule = True
    frmEdit.Show vbModal
    'add rule if edit form wasn't canceled
    If frmEdit.FormCanceled = False Then
        strTagName = frmEdit.txtTagName
        If frmEdit.optBackground.Value = True Then
            strTarget = "background"
        Else
            strTarget = "border"
        End If
        strColor = frmEdit.txtColor
        lngListIndex = Me.lstRules.ListCount
        With Me.lstRules
            .AddItem
            .List(lngListIndex, 0) = strTagName
            .List(lngListIndex, 1) = strTarget
            .List(lngListIndex, 2) = strColor
        End With
    End If
    Me.Show
    Set frmEdit = Nothing
    Exit Sub
    
error_handler:
    basSystem.log_error "frmOptions.cmdAddRule_Click"
End Sub

'-------------------------------------------------------------
' Description   : cancel the macro
' Parameter     :
' Returnvalue   :
'-------------------------------------------------------------
Private Sub cmdCancel_Click()
    
    On Error GoTo error_handler
    mblnFormCanceled = True
    basSystem.log "form was canceled"
    Me.Hide
    Exit Sub
    
error_handler:
    basSystem.log_error "frmOptions.cmdCancel_Click"
End Sub



'-------------------------------------------------------------
' Description   : start executing the macro
' Parameter     :
' Returnvalue   :
'-------------------------------------------------------------
Private Sub cmdOk_Click()

    On Error GoTo error_handler
    Me.Hide
    Exit Sub
    
error_handler:
    basSystem.log_error "frmOptions.cmdOk_Click"
End Sub



'-------------------------------------------------------------
' Description   : init options form
' Parameter     :
' Returnvalue   :
'-------------------------------------------------------------
Private Sub UserForm_Initialize()

    On Error GoTo error_handler
    mblnFormCanceled = False
    Me.chkHideAggregates.Value = cblnHideAggregatesDefault
    Me.chkDrawDomainsOnSeparatePages = cblnDrawDomainsOnSeparatePagesDefault
    Exit Sub
    
error_handler:
    basSystem.log_error "frmOptions.UserForm_Initialize"
End Sub
'-------------------------------------------------------------
' Description   : return bool if  form was canceled
' Parameter     :
' Returnvalue   :
'-------------------------------------------------------------
Public Property Get FormCanceled() As Boolean
    
    On Error GoTo error_handler
    FormCanceled = mblnFormCanceled
    Exit Property
    
error_handler:
    basSystem.log_error "frmOptions.Get FormCanceled"
End Property
'-------------------------------------------------------------
' Description   : return collection with all drawing options
' Parameter     :
' Returnvalue   : collection containing all options
'-------------------------------------------------------------
Public Property Get DrawingOptions() As Collection
    
    Dim colColorRules As New Collection
    Dim lngRule As Long
    
    On Error GoTo error_handler
    If TypeName(mcolDrawingOptions) = "Nothing" Then
        Set mcolDrawingOptions = New Collection
        mcolDrawingOptions.Add Me.chkHideAggregates.Value, cstrOptionNameHideAggregates
        mcolDrawingOptions.Add Me.chkDrawDomainsOnSeparatePages.Value, cstrOptionNameDrawDomainsOnSeparatePages
        'read color rules into collection
        For lngRule = 0 To Me.lstRules.ListCount - 1
            'color rule contains a hex color code and is identified by tag name@target e.g. status-done@background
            colColorRules.Add Me.lstRules.List(lngRule, 2), _
                Me.lstRules.List(lngRule, 0) & "@" & Me.lstRules.List(lngRule, 1)
            Next
        mcolDrawingOptions.Add colColorRules, cstrOptionNameColorRules
    End If
    Set DrawingOptions = mcolDrawingOptions
    Exit Property
    
error_handler:
    basSystem.log_error "frmOptions.Get DrawingOptions"
End Property

