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
' Returnvalue   : the worksheet used for drawing
'-------------------------------------------------------------
Public Function createDrawingDoc() As Worksheet

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
    wshDrawing.Name = "feature map"
    wbkDrawing.Windows(1).DisplayGridlines = False
    
    Application.DisplayAlerts = True
    Set createDrawingDoc = wshDrawing
    Exit Function
    
error_handler:
    basSystem.log_error "basModelVisualizer.createDrawingDoc"
End Function
'-------------------------------------------------------------
' Description   : draw aggregates, features and scenarios as use cases
'                   draw domains as surrounding boxes
' Parameter     : pwshDrawing           - xl worksheet to draw on
'                 pcolDomainModel       - domain model as structured collection
'                 pblnHideAggregates    - if true aggregates are hidden from the drawing
' Returnvalue   :
'-------------------------------------------------------------
Public Sub visualizeModel(pwshDrawing As Worksheet, pcolDomainModel As Collection, pblnHideAggregates As Boolean)

    Dim lngDomainCount As Long
    'the current drawing side of a domain box
    Dim blnDrawOnLeftSide As Boolean
    'number of use case type in the drawing (aggregate, feature, scenario)
    Dim intTypeCount As Integer
    Dim colDomain As Variant
    Dim colAggregate As Variant
    Dim lngScenarioCountLeft As Long
    Dim lngScenarioCountRight As Long
    Dim lngMaxScenarioCount As Long
    
    On Error GoTo error_handler
    lngDomainCount = 0
    'start drawing on the left side of a domain box
    blnDrawOnLeftSide = True
    'types: aggregates, features, scenarios
    If pblnHideAggregates = True Then
        intTypeCount = 2
    Else
        intTypeCount = 3
    End If
    
    For Each colDomain In pcolDomainModel
        'initialise counters
        lngScenarioCountLeft = 0
        lngScenarioCountRight = 0
        'TODO: decide on domain level if there is only one aggregate named undefined don't draw aggregates at all
        For Each colAggregate In colDomain
'        repeat with vAggregate in (get aggregates of vDomain)
'            -- start counting how many scenarios are assigned to the current aggregate
'            set vAggregateScenarioCount to 0
'            repeat with vFeature in (get features of vAggregate)
'                -- set scenario counter depending on the current drawing side
'                If vDrawOnLeftSide Is True Then
'                    set vScenarioCount to vScenarioCountLeft
'                Else
'                    set vScenarioCount to vScenarioCountRight
'                End If
'                repeat with vScenario in (get scenarios of vFeature)
'                    set vScenarioCount to vScenarioCount + 1
'                    my drawScenario(pDrawingDoc, vDomainCount, vDrawOnLeftSide, vScenarioCount, vTypeCount, Â
'                        vScenario, featureid of vFeature, featurefileid of vFeature, featurename of vFeature, domainname of vDomain)
'                end repeat
'                -- if an features has no scenarios it requires the space of one
'                if (length of scenarios of vFeature) = 0 then
'                    set vScenarioCount to vScenarioCount + 1
'                    set vAggregateScenarioCount to vAggregateScenarioCount + 1
'                End If
'                set vAggregateScenarioCount to vAggregateScenarioCount + (length of scenarios of vFeature)
'                my drawFeature(pDrawingDoc, vDomainCount, vDrawOnLeftSide, Â
'                    {currentFeatureCount:(length of scenarios of vFeature), overallCount:vScenarioCount}, Â
'                    vTypeCount, featureid of vFeature, featurefileid of vFeature, featurename of vFeature, tags of vFeature, aggregatename of vAggregate, domainname of vDomain)
'                -- count how many scenarios are on each side of the domain box to be able to calculate the size of the domain box
'                If vDrawOnLeftSide Is True Then
'                    set vScenarioCountLeft to vScenarioCount
'                Else
'                    set vScenarioCountRight to vScenarioCount
'                End If
'                -- switch side after each feature if aggregates are hidden
'                if vHideAggregates is true then set vDrawOnLeftSide to not vDrawOnLeftSide
'            end repeat
'            If vHideAggregates Is False Then
'                my drawAggregate(pDrawingDoc, vDomainCount, vDrawOnLeftSide, Â
'                    {currentAggregateCount:vAggregateScenarioCount, overallCount:vScenarioCount}, Â
'                    vTypeCount, aggregatename of vAggregate, domainname of vDomain)
'            End If
'            -- flip drawing side after each aggregate
'            if vHideAggregates is false then set vDrawOnLeftSide to not vDrawOnLeftSide
            
            'DEBUG - REMOVE
            If TypeName(colAggregate) = "Collection" Then
                lngScenarioCountRight = lngScenarioCountRight + 1
            End If
        Next
        If lngScenarioCountLeft > lngScenarioCountRight Then
            lngMaxScenarioCount = lngScenarioCountLeft
        Else
            lngMaxScenarioCount = lngScenarioCountRight
        End If
        drawDomain pwshDrawing, lngDomainCount, lngMaxScenarioCount, intTypeCount, colDomain("name")
'        my drawDomain(pDrawingDoc, vDomainCount, vMaxScenarioCount, vTypeCount, domainname of vDomain)
        lngDomainCount = lngDomainCount + 1
    Next

    
    Exit Sub
    
error_handler:
    basSystem.log_error "basModelVisualizer.visualizeModel"
End Sub
'-------------------------------------------------------------
' Description   : draw domains as surrounding boxes
' Parameter     : pwshDrawing           - xl worksheet to draw on
'                 plngDomainCount       - index of the current domain
'                 plngMaxScenarioCount  - max number scenarios for one side of the domain box
'                 pintTypeCount        - number of drawn use case types
'                 pstrDomainName
' Returnvalue   :
'-------------------------------------------------------------
Private Sub drawDomain(pwshDrawing As Worksheet, plngDomainCount As Long, plngMaxScenarioCount As Long, _
                        pintTypeCount As Integer, ByVal pstrDomainName As String)
    
    Dim lngDomainOffsetX As Long
    Dim lngOriginX As Long
    Dim lngOriginY As Long
    Dim lngDomainWidth As Long
    Dim lngDomainHeight As Long
    Dim shpDomain As Shape
    
    On Error GoTo error_handler
    lngDomainOffsetX = plngDomainCount * 2 * (pintTypeCount * 2 * clngItemPaddingX + pintTypeCount * clngItemWidth + 2 * clngDomainPaddingX)
        
    lngOriginX = clngDocPaddingX + clngDomainPaddingX + lngDomainOffsetX
    lngOriginY = clngDocPaddingY
    lngDomainWidth = 2 * (pintTypeCount * 2 * clngItemPaddingX + pintTypeCount * clngItemWidth)
    lngDomainHeight = (plngMaxScenarioCount + 1) * (2 * clngItemPaddingY + clngItemHeight)
    
    basSystem.log "draw domain " & pstrDomainName
    Set shpDomain = pwshDrawing.Shapes.AddShape(msoShapeRectangle, lngOriginX, lngOriginY, lngDomainWidth, lngDomainHeight)
    shpDomain.TextFrame.Characters.Text = pstrDomainName
    'format domain box background
    With shpDomain.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ForeColor.TintAndShade = 0
        .Transparency = 0
        .Solid
    End With
    'format domain box frame
    With shpDomain.Line
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
    End With
    'format domain box text
    With shpDomain.TextFrame2.TextRange.Font
        .Size = 24
        .Name = "Helvetica"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.ObjectThemeColor = msoThemeColorText1
        .Fill.ForeColor.TintAndShade = 0
        .Fill.Transparency = 0
        .Fill.Solid
    End With
    shpDomain.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    Exit Sub
    
error_handler:
    basSystem.log_error "basModelVisualizer.drawDomain"
End Sub
