VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOptionsTemplate 
   Caption         =   "featuremap options"
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   -16100
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
        'remove leading @ sign from tags
        If Left(strTagName, 1) = "@" Then
            strTagName = Right(strTagName, Len(strTagName) - 1)
        End If
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
        Me.saveRule strTagName, strTarget, strColor
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
' Description   : delete selected rule
' Parameter     :
' Returnvalue   :
'-------------------------------------------------------------
Private Sub cmdDeleteRule_Click()
    
    Dim strTagName As String
    Dim strTarget As String
    Dim strRuleKey As String

    On Error GoTo error_handler
    strTagName = Me.lstRules.List(Me.lstRules.ListIndex, 0)
    strTarget = Me.lstRules.List(Me.lstRules.ListIndex, 1)
    strRuleKey = "rule_" & strTagName & "@" & strTarget
    Me.lstRules.RemoveItem Me.lstRules.ListIndex
    updateButtonStatus
    ThisWorkbook.CustomDocumentProperties(strRuleKey).Delete
    Exit Sub
    
error_handler:
    basSystem.log_error "frmOptions.cmdOk_Click"
End Sub
'-------------------------------------------------------------
' Description   : edit an existing rule
' Parameter     :
' Returnvalue   :
'-------------------------------------------------------------
Private Sub cmdEditRule_Click()
    
    Dim frmEdit As New frmEditRule
    Dim strOldTagName As String
    Dim strOldTarget As String
    Dim strNewTagName As String
    Dim strNewTarget As String
    Dim strColor As String
    Dim lngListIndex As Long
    
    On Error GoTo error_handler
    'initialize edit form with list values
    strOldTagName = Me.lstRules.List(Me.lstRules.ListIndex, 0)
    frmEdit.txtTagName = strOldTagName
    strOldTarget = Me.lstRules.List(Me.lstRules.ListIndex, 1)
    If strOldTarget = "background" Then
        frmEdit.optBackground.Value = True
    Else
        frmEdit.optBorder.Value = True
    End If
    strColor = Me.lstRules.List(Me.lstRules.ListIndex, 2)
    frmEdit.txtColor = strColor
    frmEdit.lblColorPreview.BackColor = basSystem.hexToRgb(strColor)
    frmEdit.IsNewRule = False
    Me.Hide
    frmEdit.Show vbModal
    'update rule if edit form wasn't canceled
    If frmEdit.FormCanceled = False Then
        
        strNewTagName = frmEdit.txtTagName
        'remove leading @ sign from tags
        If Left(strNewTagName, 1) = "@" Then
            strNewTagName = Right(strNewTagName, Len(strNewTagName) - 1)
        End If
        
        If frmEdit.optBackground.Value = True Then
            strNewTarget = "background"
        Else
            strNewTarget = "border"
        End If
        strColor = frmEdit.txtColor
        lngListIndex = Me.lstRules.ListIndex
        With Me.lstRules
            .List(lngListIndex, 0) = strNewTagName
            .List(lngListIndex, 1) = strNewTarget
            .List(lngListIndex, 2) = strColor
        End With
        Me.replaceRule strOldTagName, strOldTarget, strNewTagName, strNewTarget, strColor
    End If
    Me.Show
    Set frmEdit = Nothing
    Exit Sub
    
error_handler:
    basSystem.log_error "frmOptions.cmdEditRule_Click"
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
' Description   : update button status for edit and delete
' Parameter     :
' Returnvalue   :
'-------------------------------------------------------------
Private Sub lstRules_AfterUpdate()

    On Error GoTo error_handler
    updateButtonStatus
    Exit Sub
    
error_handler:
    basSystem.log_error "frmOptions.lstRules_AfterUpdate"
End Sub
'-------------------------------------------------------------
' Description   : double click on a list items means edit rule
' Parameter     :
' Returnvalue   :
'-------------------------------------------------------------Ü
Private Sub lstRules_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    cmdEditRule_Click
End Sub

'-------------------------------------------------------------
' Description   : init options form
' Parameter     :
' Returnvalue   :
'-------------------------------------------------------------
Private Sub UserForm_Initialize()
    
    Dim dcpRuleProperty As DocumentProperty
    Dim strTagName As String
    Dim strTarget As String
    Dim strColor As String
    Dim lngListIndex As Long

    On Error GoTo error_handler
    mblnFormCanceled = False
    Me.chkHideAggregates.Value = cblnHideAggregatesDefault
    Me.chkDrawDomainsOnSeparatePages = cblnDrawDomainsOnSeparatePagesDefault
    'load rules from document properties
    For Each dcpRuleProperty In ThisWorkbook.CustomDocumentProperties
        If Left(dcpRuleProperty.Name, 5) = "rule_" Then
            strTagName = Mid(dcpRuleProperty.Name, 6, InStr(5, dcpRuleProperty.Name, "@") - 6)
            strTarget = Right(dcpRuleProperty.Name, Len(dcpRuleProperty.Name) - InStr(5, dcpRuleProperty.Name, "@"))
            strColor = dcpRuleProperty.Value
            lngListIndex = Me.lstRules.ListCount
            With Me.lstRules
                .AddItem
                .List(lngListIndex, 0) = strTagName
                .List(lngListIndex, 1) = strTarget
                .List(lngListIndex, 2) = strColor
            End With
        End If
    Next
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
'-------------------------------------------------------------
' Description   : save rule as document property
' Parameter     : pstrTagName
'                 pstrTarget
'                 pstrColor
'-------------------------------------------------------------
Public Sub saveRule(pstrTagName As String, pstrTarget As String, pstrColor As String)
    
    On Error Resume Next
    ThisWorkbook.CustomDocumentProperties("rule_" & pstrTagName & "@" & pstrTarget).Value = pstrColor
    If Err.Number > 0 Then
        ThisWorkbook.CustomDocumentProperties.Add Name:="rule_" & pstrTagName & "@" & pstrTarget, _
                                                    LinkToContent:=False, _
                                                    Type:=msoPropertyTypeString, _
                                                    Value:=pstrColor
    End If
    Exit Sub
    
error_handler:
    basSystem.log_error "frmOptions.saveRule"
End Sub
'-------------------------------------------------------------
' Description   : replace existing rule as document property after update
' Parameter     : pstrOldTagName
'                 pstrOldTarget
'                 pstrNewTagName
'                 pstrNewTarget
'                 pstrColor
'-------------------------------------------------------------
Public Sub replaceRule(pstrOldTagName As String, pstrOldTarget As String, _
                            pstrNewTagName As String, pstrNewTarget As String, pstrColor As String)
    
    On Error GoTo error_handler
    ThisWorkbook.CustomDocumentProperties("rule_" & pstrOldTagName & "@" & pstrOldTarget).Delete
    saveRule pstrNewTagName, pstrNewTarget, pstrColor
    Exit Sub
    
error_handler:
    basSystem.log_error "frmOptions.replaceRule"
End Sub

'-------------------------------------------------------------
' Description   : update button status for edit and delete
' Parameter     :
' Returnvalue   :
'-------------------------------------------------------------
Private Sub updateButtonStatus()

    On Error GoTo error_handler
    If Me.lstRules.ListIndex > -1 Then
        Me.cmdEditRule.Enabled = True
        Me.cmdDeleteRule.Enabled = True
    Else
        Me.cmdEditRule.Enabled = False
        Me.cmdDeleteRule.Enabled = False
    End If
    Exit Sub
    
error_handler:
    basSystem.log_error "frmOptions.lstRules_AfterUpdate"
End Sub
