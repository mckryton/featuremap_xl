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
' Returnvalue   : a collection containing all corresponding connection items (e.g. feature and scenario)
'-------------------------------------------------------------
Public Function visualizeModel(pwshDrawing As Worksheet, pcolDomainModel As Collection, _
                                pblnHideAggregates As Boolean) As Collection

    Dim lngDomainCount As Long
    'the current drawing side of a domain box
    Dim blnDrawOnLeftSide As Boolean
    'number of use case type in the drawing (aggregate, feature, scenario)
    Dim intTypeCount As Integer
    Dim colDomain As Collection
    Dim colAggregate As Collection
    Dim colFeature As Collection
    Dim colScenario As Collection
    Dim lngScenarioCount As Long
    Dim lngScenarioCountLeft As Long
    Dim lngScenarioCountRight As Long
    Dim lngMaxScenarioCount As Long
    Dim lngAggregateScenarioCount As Long
    Dim colConnectionTargetScenarios As Collection
    Dim colConnectionTargetFeatures As Collection
    Dim strShapeName As String
    
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
    
    For Each colDomain In pcolDomainModel("domains")
        'initialise counters
        lngScenarioCountLeft = 0
        lngScenarioCountRight = 0
        'TODO: decide on domain level if there is only one aggregate named undefined don't draw aggregates at all
        For Each colAggregate In colDomain("aggregates")
            'start counting how many scenarios are assigned to the current aggregate
            lngAggregateScenarioCount = 0
            Set colConnectionTargetFeatures = New Collection
            For Each colFeature In colAggregate("features")
                'set scenario counter depending on the current drawing side
                If blnDrawOnLeftSide = True Then
                    lngScenarioCount = lngScenarioCountLeft
                Else
                    lngScenarioCount = lngScenarioCountRight
                End If
                Set colConnectionTargetScenarios = New Collection
                For Each colScenario In colFeature("scenarios")
                    lngScenarioCount = lngScenarioCount + 1
                    strShapeName = drawScenario(pwshDrawing, lngDomainCount, blnDrawOnLeftSide, intTypeCount, _
                        colDomain("name"), colFeature("id"), colFeature("fileId"), colFeature("name"), _
                        lngScenarioCount, colScenario)
                    colConnectionTargetScenarios.Add strShapeName
                Next
                'if an features has no scenarios it requires the space of one
                If colFeature("scenarios").Count = 0 Then
                    lngScenarioCount = lngScenarioCount + 1
                    lngAggregateScenarioCount = lngAggregateScenarioCount + 1
                End If
                lngAggregateScenarioCount = lngAggregateScenarioCount + colFeature("scenarios").Count
                strShapeName = drawFeature(pwshDrawing, lngDomainCount, blnDrawOnLeftSide, intTypeCount, _
                        colDomain("name"), colAggregate("name"), colFeature("id"), colFeature("fileId"), _
                        colFeature("name"), colFeature("scenarios").Count, lngScenarioCount)
                colConnectionTargetFeatures.Add strShapeName
                drawConnections pwshDrawing, strShapeName, colConnectionTargetScenarios
                Set colConnectionTargetScenarios = Nothing
                'count how many scenarios are on each side of the domain box to be able to calculate the size of the domain box
                If blnDrawOnLeftSide = True Then
                    lngScenarioCountLeft = lngScenarioCount
                Else
                    lngScenarioCountRight = lngScenarioCount
                End If
                'switch side after each feature if aggregates are hidden
                If pblnHideAggregates = True Then
                    blnDrawOnLeftSide = Not blnDrawOnLeftSide
                End If
            Next
            If pblnHideAggregates = False Then
                strShapeName = drawAggregate(pwshDrawing, lngDomainCount, blnDrawOnLeftSide, intTypeCount, _
                            colDomain("name"), colAggregate("name"), lngScenarioCount, lngAggregateScenarioCount)
                drawConnections pwshDrawing, strShapeName, colConnectionTargetFeatures
                Set colConnectionTargetFeatures = Nothing
            End If
            'flip drawing side after each aggregate
            If pblnHideAggregates = False Then
                blnDrawOnLeftSide = Not blnDrawOnLeftSide
            End If
        Next
        If lngScenarioCountLeft > lngScenarioCountRight Then
            lngMaxScenarioCount = lngScenarioCountLeft
        Else
            lngMaxScenarioCount = lngScenarioCountRight
        End If
        drawDomain pwshDrawing, lngDomainCount, lngMaxScenarioCount, intTypeCount, colDomain("name")
        lngDomainCount = lngDomainCount + 1
    Next

    Exit Function
    
error_handler:
    basSystem.log_error "basModelVisualizer.visualizeModel"
End Function
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
    lngDomainOffsetX = plngDomainCount * 2 * (pintTypeCount * 2 * clngItemPaddingX _
                        + pintTypeCount * clngItemWidth + 2 * clngDomainPaddingX)
    lngOriginX = clngDocPaddingX + clngDomainPaddingX + lngDomainOffsetX
    lngOriginY = clngDocPaddingY
    lngDomainWidth = 2 * (pintTypeCount * 2 * clngItemPaddingX + pintTypeCount * clngItemWidth)
    lngDomainHeight = (plngMaxScenarioCount + 1) * (2 * clngItemPaddingY + clngItemHeight)
    
    basSystem.log "draw domain >" & pstrDomainName & "<"
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
    shpDomain.ZOrder msoSendToBack
    Exit Sub
    
error_handler:
    basSystem.log_error "basModelVisualizer.drawDomain"
End Sub
'-------------------------------------------------------------
' Description   : draw aggregates as use cases
' Parameter     : pwshDrawing           - xl worksheet to draw on
'                 plngDomainCount       - index of the current domain
'                 pblnDrawOnLeftSide    -
'                 pintTypeCount        - number of drawn use case types
'                 pstrDomainName
'                 pstrAggregateName
'                 plngScenarioCount
'                 plngCurrentAggregateScenarioCount
' Returnvalue   : shape name as string
'-------------------------------------------------------------
Private Function drawAggregate(pwshDrawing As Worksheet, plngDomainCount As Long, ByVal pblnDrawOnLeftSide As Boolean, _
                        pintTypeCount As Integer, ByVal pstrDomainName As String, pstrAggregateName As String, _
                        plngScenarioCount As Long, plngCurrentAggregateScenarioCount As Long) As String
    
    Dim lngDomainOffsetX As Long
    Dim lngOriginX As Long
    Dim lngOriginY As Long
    Dim lngSideOffsetX As Long
    Dim lngDomainWidth As Long
    Dim lngDomainHeight As Long
    Dim shpAggregate As Shape
    Dim lngCurrentAggregateScenarioCount As Long
    Dim lngOtherAggregateScenarioCount As Long
    Dim lngScenarioCountOffsetY As Long
    
    On Error GoTo error_handler
    'get the number of the scenarios assigned to the current aggregate
    lngCurrentAggregateScenarioCount = plngCurrentAggregateScenarioCount
    'get the number of all scenarios drawn on the current side of the domain box minus the number of the current feature
    lngOtherAggregateScenarioCount = plngScenarioCount - lngCurrentAggregateScenarioCount

    'calculate aggregate position
    lngScenarioCountOffsetY = (lngOtherAggregateScenarioCount * (2 * clngItemPaddingY + clngItemHeight))
    lngOriginY = clngDocPaddingY + lngScenarioCountOffsetY + (lngCurrentAggregateScenarioCount / 2 * _
                    (2 * clngItemPaddingY + clngItemHeight)) + (clngItemPaddingY + clngItemHeight / 2)
    'TODO: this breaks if some domains hide aggregates and some not
    lngDomainOffsetX = plngDomainCount * 2 * (pintTypeCount * 2 * clngItemPaddingX + pintTypeCount * clngItemWidth _
                        + 2 * clngDomainPaddingX)
    If pblnDrawOnLeftSide = True Then
        'draw aggregate on the left side of the domain box
        lngSideOffsetX = 0
        lngOriginX = clngDocPaddingX + clngDomainPaddingX + lngDomainOffsetX + clngItemPaddingX
    Else
        'draw aggregate on the right side of the domain box
        lngSideOffsetX = (pintTypeCount * (2 * clngItemPaddingX + clngItemWidth))
        lngOriginX = clngDocPaddingX + clngDomainPaddingX + lngDomainOffsetX + lngSideOffsetX _
                        + ((pintTypeCount - 1) * 2 + 1) * clngItemPaddingX + 2 * clngItemWidth
    End If

    basSystem.log "draw aggregate >" & pstrAggregateName & "<"

'    tell application "OmniGraffle"
'        set vLayerModel to layer "functions" of canvas "model" of pDrawingDoc
'        make new shape at end of graphics of vLayerModel with properties Â
'            {name:"Circle", textSize:{0.8, 0.7}, size:{cItemWidth, cItemHeight}, text:{alignment:center, text:pAggregateName}, origin:{vOriginX, vOriginY}, magnets:{{0, 0.5}, {0, -0.5}, {0.5, 0}, {-0.5, 0}, {-0.29, -0.41}, {-0.29, 0.41}, {0.29, 0.41}, {0.29, -0.41}}, textPosition:{0.1, 0.15}, vertical padding:0, thickness:2, user data:{itemtype:"aggregate", domain:pDomainName}}
'    end tell

    Set shpAggregate = pwshDrawing.Shapes.AddShape(msoShapeOval, lngOriginX, lngOriginY, clngItemWidth, clngItemHeight)
    shpAggregate.TextFrame.Characters.Text = pstrAggregateName
    'format domain box background
    With shpAggregate.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ForeColor.TintAndShade = 0
        .Transparency = 0
        .Solid
    End With
    'format domain box frame
    With shpAggregate.Line
        .Weight = 3
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
    End With
    'format domain box text
    With shpAggregate.TextFrame2.TextRange.Font
        .Size = 14
        .Name = "Helvetica"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.ObjectThemeColor = msoThemeColorText1
        .Fill.ForeColor.TintAndShade = 0
        .Fill.Transparency = 0
        .Fill.Solid
    End With
    shpAggregate.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    drawAggregate = shpAggregate.Name
    Exit Function
    
error_handler:
    basSystem.log_error "basModelVisualizer.drawAggregate"
End Function
'-------------------------------------------------------------
' Description   : draw features as use cases
' Parameter     : pwshDrawing           - xl worksheet to draw on
'                 plngDomainCount       - index of the current domain
'                 pblnDrawOnLeftSide    -
'                 pintTypeCount        - number of drawn use case types
'                 pstrDomainName
'                 pstrAggregateName
'                 plngFeatureId
'                 plngFeatureFileId
'                 pstrFeatureName
'                 plngScenarioCount
' Returnvalue   : shape name as string
'-------------------------------------------------------------
Private Function drawFeature(pwshDrawing As Worksheet, plngDomainCount As Long, ByVal pblnDrawOnLeftSide As Boolean, _
                        pintTypeCount As Integer, ByVal pstrDomainName As String, pstrAggregateName As String, _
                        plngFeatureId As Long, plngFeatureFileId As Long, pstrFeatureName As String, _
                        plngCurrentFeatureScenarioCount As Long, plngAllScenarioCount As Long) As String
    
    Dim lngDomainOffsetX As Long
    Dim lngOriginX As Long
    Dim lngOriginY As Long
    Dim lngSideOffsetX As Long
    Dim lngDomainWidth As Long
    Dim lngDomainHeight As Long
    Dim shpFeature As Shape
    Dim lngCurrentFeatureScenarioCount As Long
    Dim lngOtherFeaturesScenarioCount As Long
    Dim lngScenarioCountOffsetY As Long
    
    On Error GoTo error_handler
    'get the number of the scenarios assigned to the current feature
    lngCurrentFeatureScenarioCount = plngCurrentFeatureScenarioCount
    
    If lngCurrentFeatureScenarioCount = 0 Then
        'leave space for one scenario if the feature hasn't one
        lngCurrentFeatureScenarioCount = 1
    End If
    'get the number of all scenarios drawn on the current side of the domain box minus the number of the current feature
    lngOtherFeaturesScenarioCount = plngAllScenarioCount - lngCurrentFeatureScenarioCount

    'calculate feature position
    lngScenarioCountOffsetY = (lngOtherFeaturesScenarioCount * (2 * clngItemPaddingY + clngItemHeight))
    lngOriginY = clngDocPaddingY + lngScenarioCountOffsetY + (lngCurrentFeatureScenarioCount / 2 * (2 * clngItemPaddingY _
                    + clngItemHeight)) + (clngItemPaddingY + clngItemHeight / 2)
    'TODO: this breaks if some domains hide aggregates and some not
    lngDomainOffsetX = plngDomainCount * 2 * (pintTypeCount * 2 * clngItemPaddingX + pintTypeCount * clngItemWidth _
                        + 2 * clngDomainPaddingX)
    If pblnDrawOnLeftSide = True Then
        'draw feature on the left side of the domain box
        lngSideOffsetX = 0
        lngOriginX = clngDocPaddingX + clngDomainPaddingX + lngDomainOffsetX + lngSideOffsetX + ((pintTypeCount - 2) _
                        * (2 * clngItemPaddingX + clngItemWidth)) + clngItemPaddingX
    Else
        'draw feature on the right side of the domain box
        lngSideOffsetX = (pintTypeCount * (2 * clngItemPaddingX + clngItemWidth))
        lngOriginX = clngDocPaddingX + clngDomainPaddingX + lngDomainOffsetX + lngSideOffsetX + 3 * clngItemPaddingX _
                        + clngItemWidth
    End If

'    -- set fill color depending on the feature status
'    set vStatusColor to my getStatusColor(status of pFeatureTags)
'
    basSystem.log "draw feature >" & pstrFeatureName & "<"
'    tell application "OmniGraffle"
'        set vLayerModel to layer "functions" of canvas "model" of pDrawingDoc
'        make new shape at end of graphics of vLayerModel with properties Â
'            {name:"Circle", textSize:{0.8, 0.7}, size:{cItemWidth, cItemHeight}, text:{alignment:center, text:pFeatureName}, origin:{vOriginX, vOriginY}, magnets:{{0, 0.5}, {0, -0.5}, {0.5, 0}, {-0.5, 0}, {-0.29, -0.41}, {-0.29, 0.41}, {0.29, 0.41}, {0.29, -0.41}}, textPosition:{0.1, 0.15}, thickness:1, vertical padding:0, user data:{aggregate:pAggregateName, itemtype:"feature", domain:pDomainName, featureid:pFeatureId, featurefileid:pFeatureFileId}, fill color:vStatusColor}
'    end tell
    
    Set shpFeature = pwshDrawing.Shapes.AddShape(msoShapeOval, lngOriginX, lngOriginY, clngItemWidth, clngItemHeight)
    shpFeature.TextFrame.Characters.Text = pstrFeatureName
    'format domain box background
    With shpFeature.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ForeColor.TintAndShade = 0
        .Transparency = 0
        .Solid
    End With
    'format domain box frame
    With shpFeature.Line
        .Weight = 2
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
    End With
    'format domain box text
    With shpFeature.TextFrame2.TextRange.Font
        .Size = 14
        .Name = "Helvetica"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.ObjectThemeColor = msoThemeColorText1
        .Fill.ForeColor.TintAndShade = 0
        .Fill.Transparency = 0
        .Fill.Solid
    End With
    shpFeature.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    drawFeature = shpFeature.Name
    Exit Function
    
error_handler:
    basSystem.log_error "basModelVisualizer.drawFeature"
End Function
'-------------------------------------------------------------
' Description   : draw scenarios as use cases
' Parameter     : pwshDrawing           - xl worksheet to draw on
'                 plngDomainCount       - index of the current domain
'                 pblnDrawOnLeftSide    -
'                 pintTypeCount        - number of drawn use case types
'                 pstrDomainName
'                 plngFeatureId
'                 plngFeatureFileId
'                 pstrFeatureName
'                 plngScenarioCount
'                 pcolScenario
' Returnvalue   : shape name as string
'-------------------------------------------------------------
Private Function drawScenario(pwshDrawing As Worksheet, plngDomainCount As Long, ByVal pblnDrawOnLeftSide As Boolean, _
                        pintTypeCount As Integer, ByVal pstrDomainName As String, plngFeatureId As Long, _
                        plngFeatureFileId As Long, pstrFeatureName As String, plngScenarioCount As Long, _
                        pcolScenario As Collection) As String
    
    Dim lngDomainOffsetX As Long
    Dim lngOriginX As Long
    Dim lngOriginY As Long
    Dim lngSideOffsetX As Long
    Dim lngDomainWidth As Long
    Dim lngDomainHeight As Long
    Dim shpScenario As Shape
    
    On Error GoTo error_handler
    lngDomainOffsetX = plngDomainCount * 2 * (pintTypeCount * 2 * clngItemPaddingX _
                        + pintTypeCount * clngItemWidth + 2 * clngDomainPaddingX)
    If pblnDrawOnLeftSide = False Then
        'draw scenario on the right side of the domain box
        lngSideOffsetX = (pintTypeCount * (2 * clngItemPaddingX + clngItemWidth))
        lngOriginX = clngDocPaddingX + clngDomainPaddingX + lngDomainOffsetX + lngSideOffsetX + clngItemPaddingX
    Else
        'draw scenario on the left side of the domain box
        lngSideOffsetX = 0
        lngOriginX = clngDocPaddingX + clngDomainPaddingX + lngDomainOffsetX + lngSideOffsetX _
                        + (pintTypeCount - 1) * (2 * clngItemPaddingX + clngItemWidth) _
                        + clngItemPaddingX
    End If
    lngOriginY = clngDocPaddingY + plngScenarioCount * ((2 * clngItemPaddingY) + clngItemHeight)

'
'    -- set fill color depending on the feature status
'    set vStatusColor to my getStatusColor(status of tags of pScenario)
'
    basSystem.log "draw scenario >" & pcolScenario("name") & "<"
'    tell application "OmniGraffle"
'        set vLayerModel to layer "functions" of canvas "model" of pDrawingDoc
'        make new shape at end of graphics of vLayerModel with properties Â
'            {name:"Circle", textSize:{0.8, 0.7}, size:{cItemWidth, cItemHeight}, text:{alignment:center, text:name of pScenario}, origin:{vOriginX, vOriginY}, magnets:{{0, 0.5}, {0, -0.5}, {0.5, 0}, {-0.5, 0}, {-0.29, -0.41}, {-0.29, 0.41}, {0.29, 0.41}, {0.29, -0.41}}, textPosition:{0.1, 0.15}, thickness:0.25, vertical padding:0, user data:{featureid:pFeatureId, featurefileid:pFeatureFileId, feature:pFeatureName, itemtype:"scenario", domain:pDomainName}, fill color:vStatusColor}
'    end tell
    
    Set shpScenario = pwshDrawing.Shapes.AddShape(msoShapeOval, lngOriginX, lngOriginY, clngItemWidth, clngItemHeight)
    shpScenario.TextFrame.Characters.Text = pcolScenario("name")
    'format domain box background
    With shpScenario.Fill
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ForeColor.TintAndShade = 0
        .Transparency = 0
        .Solid
    End With
    'format domain box frame
    With shpScenario.Line
        .Weight = 0.5
        .Visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorText1
        .ForeColor.TintAndShade = 0
    End With
    'format domain box text
    With shpScenario.TextFrame2.TextRange.Font
        .Size = 14
        .Name = "Helvetica"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.ObjectThemeColor = msoThemeColorText1
        .Fill.ForeColor.TintAndShade = 0
        .Fill.Transparency = 0
        .Fill.Solid
    End With
    shpScenario.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    drawScenario = shpScenario.Name
    Exit Function
    
error_handler:
    basSystem.log_error "basModelVisualizer.drawScenario"
End Function
'-------------------------------------------------------------
' Description   : set height of all domain boxes to the same size
' Parameter     : pwshDrawing           - xl worksheet to draw on
' Returnvalue   :
'-------------------------------------------------------------
Public Sub levelDomainHeight(pwshDrawing As Worksheet)

    Dim shpAnyShape As Shape
    Dim shpDomainBox As Shape
    Dim lngMaxHeight As Long
    Dim colDomainShapes As New Collection
    
    On Error GoTo error_handler
    lngMaxHeight = 0
    For Each shpAnyShape In pwshDrawing.Shapes
        'only domain boxes are rectangles
        If shpAnyShape.AutoShapeType = msoShapeRectangle Then
            colDomainShapes.Add shpAnyShape
            If shpAnyShape.Height > lngMaxHeight Then
                lngMaxHeight = shpAnyShape.Height
            End If
        End If
    Next
    'now level domain boxes
    For Each shpDomainBox In colDomainShapes
        shpDomainBox.Height = lngMaxHeight
    Next
    Exit Sub
    
error_handler:
    basSystem.log_error "basModelVisualizer.levelDomainHeight"
End Sub
'-------------------------------------------------------------
' Description   : connect feature map items with each other (eg. feature with scenario)
' Parameter     : pwshDrawing           - xl worksheet to draw on
'                 pstrSourceShape       - name of the source shape
'                 pcolTargetShapes      - collection containing the target shape names
' Returnvalue   :
'-------------------------------------------------------------
Public Sub drawConnections(pwshDrawing As Worksheet, pstrSourceShapeName As String, pcolTargetShapes As Collection)

    Dim strShapeTargetName As Variant
    Dim shpConnector As Shape

    On Error GoTo error_handler
    For Each strShapeTargetName In pcolTargetShapes
        Set shpConnector = pwshDrawing.Shapes.AddConnector(msoConnectorCurve, 0, 0, 50, 50)
        With shpConnector.Line
            .ForeColor.RGB = RGB(0, 0, 0)
            .DashStyle = msoLineDash
            .EndArrowheadStyle = msoArrowheadTriangle
            .EndArrowheadLength = msoArrowheadLong
        End With
        With shpConnector.ConnectorFormat
            .BeginConnect pwshDrawing.Shapes(pstrSourceShapeName), 1
            .EndConnect pwshDrawing.Shapes(strShapeTargetName), 1
        End With
        shpConnector.RerouteConnections
    
    Next
    Exit Sub
    
error_handler:
    basSystem.log_error "basModelVisualizer.drawConnections"
End Sub
