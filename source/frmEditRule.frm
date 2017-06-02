VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEditRule 
   Caption         =   "Edit rule"
   ClientHeight    =   5060
   ClientLeft      =   0
   ClientTop       =   -8280.001
   ClientWidth     =   7000
   OleObjectBlob   =   "frmEditRule.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmEditRule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------
' Description  : UI for editing a rule to set color of drawing items
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
Option Explicit

Dim mblnIsNewRule As Boolean
Dim mblnFormCanceled As Boolean

'-------------------------------------------------------------
' Description   : cancel rule update
' Parameter     :
' Returnvalue   :
'-------------------------------------------------------------
Private Sub cmdCancel_Click()

    On Error GoTo error_handler
    Me.FormCanceled = True
    Me.Hide
    Exit Sub
    
error_handler:
    basSystem.log_error "frmEditRule.cmdCancel_Click"
End Sub
'-------------------------------------------------------------
' Description   : detect if rule is new
' Parameter     :
' Returnvalue   :
'-------------------------------------------------------------
Public Property Get IsNewRule() As Boolean
    On Error GoTo error_handler
    IsNewRule = mblnIsNewRule
    Exit Property
    
error_handler:
    basSystem.log_error "frmEditRule.Get IsNewRule"
End Property
'-------------------------------------------------------------
' Description   : define if rule is new
' Parameter     :
' Returnvalue   :
'-------------------------------------------------------------
Public Property Let IsNewRule(ByVal pblnIsNewRule As Boolean)
    On Error GoTo error_handler
    mblnIsNewRule = pblnIsNewRule
    Exit Property
    
error_handler:
    basSystem.log_error "frmEditRule.Let IsNewRule"
End Property
'-------------------------------------------------------------
' Description   : show built-in color chooser dialog
' Parameter     :
' Returnvalue   :
'-------------------------------------------------------------
Private Sub cmdChooseColor_Click()
    
    Dim intColorFrmStatus As Integer
    Dim strHexColor As String

    On Error GoTo error_handler
    intColorFrmStatus = Application.Dialogs(xlDialogEditColor).Show(30)
    If intColorFrmStatus = -1 Then
        strHexColor = Right$("000000" & hex$(ThisWorkbook.Colors(30)), 6)
        Me.txtColor = strHexColor
        Me.lblColorPreview.BackColor = ThisWorkbook.Colors(30)
    End If
    Exit Sub
    
error_handler:
    basSystem.log_error "frmEditRule.cmdChooseColor_Click"
End Sub

'-------------------------------------------------------------
' Description   : create/update rule
' Parameter     :
' Returnvalue   :
'-------------------------------------------------------------
Private Sub cmdSave_Click()

    On Error GoTo error_handler
    Me.FormCanceled = False
    Me.Hide
    Exit Sub
    
error_handler:
    basSystem.log_error "frmEditRule.cmdSave_Click"
End Sub

'-------------------------------------------------------------
' Description   : set form title to add or edit
' Parameter     :
' Returnvalue   :
'-------------------------------------------------------------
Private Sub UserForm_Activate()
    
    On Error GoTo error_handler
    If Me.IsNewRule Then
        Me.Caption = "Add rule"
    Else
        Me.Caption = "Edit rule"
    End If
    Exit Sub
    
error_handler:
    basSystem.log_error "frmEditRule.UserForm_Activate"
End Sub

'-------------------------------------------------------------
' Description   : asume new rule as default
' Parameter     :
' Returnvalue   :
'-------------------------------------------------------------
Private Sub UserForm_Initialize()
    
    On Error GoTo error_handler
    mblnIsNewRule = True
    mblnFormCanceled = False
    Exit Sub
    
error_handler:
    basSystem.log_error "frmEditRule.UserForm_Initialize"
End Sub
'-------------------------------------------------------------
' Description   : return true if form was canceled
' Parameter     :
' Returnvalue   :
'-------------------------------------------------------------
Public Property Get FormCanceled() As Boolean
    
    On Error GoTo error_handler
    FormCanceled = mblnFormCanceled
    Exit Property
    
error_handler:
    basSystem.log_error "frmEditRule.Get FormCanceled"
End Property
'-------------------------------------------------------------
' Description   : save status if for  was canceled
' Parameter     :
' Returnvalue   :
'-------------------------------------------------------------
Public Property Let FormCanceled(pblnFormCanceled As Boolean)
    
    On Error GoTo error_handler
    mblnFormCanceled = pblnFormCanceled
    Exit Property
    
error_handler:
    basSystem.log_error "frmEditRule.Let FormCanceled"
End Property


