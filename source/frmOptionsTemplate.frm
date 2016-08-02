VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOptionsTemplate 
   Caption         =   "featuremap options"
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   -3220
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
' Description  : uI for configuring the script
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

'Options
Option Explicit
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
' Description   : return bool to decide if aggrepgates should
'                   be left from the drawing
' Parameter     :
' Returnvalue   : checkbox value as boolean
'-------------------------------------------------------------
Public Property Get HideAggregates() As Boolean
    
    On Error GoTo error_handler
    HideAggregates = Me.chkHideAggregates.Value
    Exit Property
    
error_handler:
    basSystem.log_error "frmOptions.Get HideAggregatesDefault"
End Property
'-------------------------------------------------------------
' Description   : return bool to decide if domains should be
'                   drawn on dedicated worksheets
' Parameter     :
' Returnvalue   : checkbox value as boolean
'-------------------------------------------------------------
Public Property Get DrawDomainsOnSeparatePages() As Boolean
    
    On Error GoTo error_handler
    DrawDomainsOnSeparatePages = Me.chkDrawDomainsOnSeparatePages.Value
    Exit Property
    
error_handler:
    basSystem.log_error "frmOptions.Get DrawDomainsOnSeparatePages"
End Property

