import adsk.core
import os
import io
import sys

import adsk.fusion
from ...lib import fusionAddInUtils as futil
from ... import config
import math

app = adsk.core.Application.get()
ui = app.userInterface

CMD_ID = f'{config.COMPANY_NAME}_{config.ADDIN_NAME}_importAirfoilCSV'
CMD_NAME = 'Import Airfoil CSV'
CMD_Description = 'Import airfoil curve from CSV file.'

IS_PROMOTED = True

WORKSPACE_ID = 'FusionSolidEnvironment'
PANEL_ID = 'SolidCreatePanel'
COMMAND_BESIDE_ID = 'PrimitivePipe'

local_handlers = []
sketch = None

_filePass = None
_swapDirection = False

def start():
    try:
        futil.log(f'{CMD_NAME} Command started')
        cmdDef = ui.commandDefinitions.itemById(CMD_ID)
        if cmdDef:
            cmdDef.deleteMe()
        
        icon_folder = os.path.join(os.path.dirname(os.path.abspath(__file__)), './resources/CommandIcon', '')

        cmdDef = ui.commandDefinitions.addButtonDefinition(CMD_ID, CMD_NAME, CMD_Description, icon_folder)

        futil.add_handler(cmdDef.commandCreated, command_created)

        workspace = ui.workspaces.itemById(WORKSPACE_ID)

        panel = workspace.toolbarPanels.itemById(PANEL_ID)

        control = panel.controls.addCommand(cmdDef, COMMAND_BESIDE_ID, False)

        control.isPromoted = IS_PROMOTED
    except Exception as error:
        ui.messageBox(str(error))

def stop():
    try:
        workspace = ui.workspaces.itemById(WORKSPACE_ID)
        panel = workspace.toolbarPanels.itemById(PANEL_ID)
        command_control = panel.controls.itemById(CMD_ID)
        command_definition = ui.commandDefinitions.itemById(CMD_ID)

        if command_control:
            command_control.deleteMe()

        if command_definition:
            command_definition.deleteMe()
    except Exception as error:
        ui.messageBox(str(error))

def command_created(args: adsk.core.CommandCreatedEventArgs):
    try:
        inputs = args.command.commandInputs
        
        futil.add_handler(args.command.inputChanged, command_changed, local_handlers=local_handlers)
        futil.add_handler(args.command.execute, command_execute, local_handlers=local_handlers)
        futil.add_handler(args.command.executePreview, command_preview, local_handlers=local_handlers)
        futil.add_handler(args.command.destroy, command_destroy, local_handlers=local_handlers)
        futil.add_handler(args.command.validateInputs, command_validate_inputs, local_handlers=local_handlers)
        inputs.addBoolValueInput('fileSelectButton', 'Airfoil file', False, '', False)
        fileNameTextBox = inputs.addTextBoxCommandInput('fileNameText', 'File Name', _filePass.rsplit('/', 1)[1] if _filePass else '', 1, True)
        fileNameTextBox.isFullWidth = True
        fileNameTextBox.isVisible = _filePass is not None
        #inputs.addTextBoxCommandInput('filePassText', 'File Pass', 'Selected file pass', 1, True)
        inputs.addBoolValueInput('edgeLineBool', 'Connect Trailing Edge', True, '', True)
        axisSelectButton = inputs.addButtonRowCommandInput('MethodSelectButton', 'Method', False)
        axisSelectButton.listItems.clear()
        axisSelectButton.listItems.add('Line', True, os.path.join(os.path.dirname(os.path.abspath(__file__)), './resources/OneLineSelect'), -1)
        axisSelectButton.listItems.add('Points', False, os.path.join(os.path.dirname(os.path.abspath(__file__)), './resources/TwoPointsSelect'), -1)
        axisInput = inputs.addSelectionInput('axis', 'Axis', 'Select Axis')
        axisInput.isVisible = True
        axisInput.setSelectionLimits(1, 1)
        axisInput.addSelectionFilter('LinearEdges')
        axisInput.addSelectionFilter('SketchLines')
        #axisInput.addSelectionFilter('ConstructionLines')
        originPointInput = inputs.addSelectionInput('originPoint', 'Origin Point', 'Select Origin Point')
        originPointInput.isVisible = False
        originPointInput.setSelectionLimits(0, 0)
        originPointInput.addSelectionFilter('Vertices')
        originPointInput.addSelectionFilter('SketchPoints')
        originPointInput.addSelectionFilter('ConstructionPoints')
        edgePointInput = inputs.addSelectionInput('edgePoint', 'Trailing Edge Point', 'Select Trailing Edge Point')
        edgePointInput.isVisible = False
        edgePointInput.setSelectionLimits(0, 0)
        edgePointInput.addSelectionFilter('Vertices')
        edgePointInput.addSelectionFilter('SketchPoints')
        edgePointInput.addSelectionFilter('ConstructionPoints')
        inputs.addBoolValueInput('SwapDirection', 'Swap Direction', False, os.path.join(os.path.dirname(os.path.abspath(__file__)), './resources/SwapIcon'), False)
        inputs.addAngleValueCommandInput('tiltAngle', 'Tilt Angle', adsk.core.ValueInput.createByReal(0)).isVisible = False
    except Exception as error:
        ui.messageBox(str(error))

def command_changed(args: adsk.core.InputChangedEventArgs):
    try:
        global _swapDirection
        inputs = args.inputs
        input = args.input
        if input.id == 'fileSelectButton':
            #ui.messageBox(f'ButtonClicked. The status is {input.value}.')
            dialog = ui.createFileDialog()
            dialog.title = 'Open Airfoil File'
            dialog.filter = 'Airfoil File (*.csv;*.dat) ;;All Files (*.*)'
            if dialog.showOpen() != adsk.core.DialogResults.DialogOK:
                return
            filepass = dialog.filename
            filename = filepass.rsplit('/', 1)[1]
            global _filePass
            _filePass = filepass
            args.inputs.itemById('fileNameText').text = filename
            args.inputs.itemById('fileNameText').isVisible = True
        elif input.id == 'axis':
            tiltAngleInput : adsk.core.AngleValueCommandInput = inputs.itemById('tiltAngle')
            if input.selectionCount == 1:
                tiltAngleInput.isVisible = True
                axisEntity = input.selection(0).entity
                #ui.messageBox(f'Selected axis entity type: {axisEntity.objectType}')
                if axisEntity.objectType == 'adsk::fusion::SketchLine':
                    originPoint = axisEntity.startSketchPoint.worldGeometry
                    edgePoint = axisEntity.endSketchPoint.worldGeometry
                else:
                    originPoint = axisEntity.startVertex.geometry
                    edgePoint = axisEntity.endVertex.geometry
                if _swapDirection:
                    originPoint, edgePoint = edgePoint, originPoint
                setAngleManipulator(tiltAngleInput, originPoint, edgePoint)
            else:
                tiltAngleInput.isVisible = False
        elif input.id == 'originPoint' or input.id =='edgePoint':
            tiltAngleInput : adsk.core.AngleValueCommandInput = inputs.itemById('tiltAngle')
    
            if inputs.itemById('originPoint').selectionCount == 1 and inputs.itemById('edgePoint').selectionCount == 1:
                tiltAngleInput.isVisible = True
                originPointEntity = inputs.itemById('originPoint').selection(0).entity
                
                #ui.messageBox(f'Selected axis entity type: {originPointEntity.objectType}')
                if originPointEntity.objectType == 'adsk::fusion::SketchPoint':
                    originPoint = originPointEntity.worldGeometry
                else:
                    originPoint = originPointEntity.geometry
                edgePointEntity = inputs.itemById('edgePoint').selection(0).entity
                if edgePointEntity.objectType == 'adsk::fusion::SketchPoint':
                    edgePoint = edgePointEntity.worldGeometry
                else:
                    edgePoint = edgePointEntity.geometry
                setAngleManipulator(tiltAngleInput, originPoint, edgePoint)

            else:
                tiltAngleInput.isVisible = False
        elif input.id == 'MethodSelectButton':
            axisInput : adsk.core.SelectionCommandInput = inputs.itemById('axis')
            originPointInput : adsk.core.SelectionCommandInput = inputs.itemById('originPoint')
            edgePointInput : adsk.core.SelectionCommandInput = inputs.itemById('edgePoint')
            tiltAngleInput : adsk.core.AngleValueCommandInput = inputs.itemById('tiltAngle')
            tiltAngleInput.isVisible = False
            if input.selectedItem.name == 'Line':
                axisInput.isVisible = True
                axisInput.setSelectionLimits(1, 1)
                originPointInput.isVisible = False
                originPointInput.setSelectionLimits(0, 0)
                edgePointInput.isVisible = False
                edgePointInput.setSelectionLimits(0, 0)
                axisInput.isEnabled = True
            else:
                originPointInput.isVisible = True
                originPointInput.setSelectionLimits(1, 1)
                edgePointInput.isVisible = True
                edgePointInput.setSelectionLimits(1, 1)
                if axisInput.selectionCount == 1:
                    axisEntity = axisInput.selection(0).entity
                    if axisEntity.objectType == 'adsk::fusion::SketchLine':
                        originPoint = axisEntity.startSketchPoint
                        edgePoint = axisEntity.endSketchPoint
                    else:
                        originPoint = axisEntity.startVertex
                        edgePoint = axisEntity.endVertex
                    if _swapDirection:
                        originPoint, edgePoint = edgePoint, originPoint
                    originPointInput.addSelection(originPoint)
                    edgePointInput.addSelection(edgePoint)
                    setAngleManipulator(tiltAngleInput, originPoint.geometry if originPoint.objectType != 'adsk::fusion::SketchPoint' else originPoint.worldGeometry, edgePoint.geometry if edgePoint.objectType != 'adsk::fusion::SketchPoint' else edgePoint.worldGeometry)
                    tiltAngleInput.isVisible = True
                originPointInput.isEnabled = True
                axisInput.isVisible = False
                axisInput.setSelectionLimits(0, 0)
        elif input.id == 'SwapDirection':
            if inputs.itemById('MethodSelectButton').selectedItem.name == 'Points':
                originPointSelection : adsk.core.SelectionCommandInput = inputs.itemById('originPoint')
                edgePointSelection : adsk.core.SelectionCommandInput = inputs.itemById('edgePoint')
                if originPointSelection.selectionCount == 1:
                    if edgePointSelection.selectionCount == 1:
                        edgePoint : adsk.fusion.SketchPoint = originPointSelection.selection(0).entity
                        originPoint : adsk.fusion.SketchPoint = edgePointSelection.selection(0).entity
                        originPointSelection.isEnabled = False
                        originPointSelection.clearSelection()
                        edgePointSelection.clearSelection()
                        originPointSelection.addSelection(originPoint)
                        edgePointSelection.addSelection(edgePoint)
                        originPointSelection.isEnabled = True
                        setAngleManipulator(inputs.itemById('tiltAngle'), originPoint.geometry if originPoint.objectType != 'adsk::fusion::SketchPoint' else originPoint.worldGeometry, edgePoint.geometry if edgePoint.objectType != 'adsk::fusion::SketchPoint' else edgePoint.worldGeometry)
                    else:
                        edgePointSelection.addSelection(originPointSelection.selection(0).entity)
                        originPointSelection.clearSelection()
                elif edgePointSelection.selectionCount == 1:
                    originPointSelection.addSelection(edgePointSelection.selection(0).entity)
                    edgePointSelection.clearSelection()
            else:
                _swapDirection = not _swapDirection
                tiltAngleInput : adsk.core.AngleValueCommandInput = inputs.itemById('tiltAngle')
                axisSelection : adsk.core.SelectionCommandInput = inputs.itemById('axis')
                if axisSelection.selectionCount == 1:
                    axisEntity = axisSelection.selection(0).entity
                    if axisEntity.objectType == 'adsk::fusion::SketchLine':
                        originPoint = axisEntity.startSketchPoint.worldGeometry
                        edgePoint = axisEntity.endSketchPoint.worldGeometry
                    else:
                        originPoint = axisEntity.startVertex.geometry
                        edgePoint = axisEntity.endVertex.geometry
                    if _swapDirection:
                        originPoint, edgePoint = edgePoint, originPoint
                    setAngleManipulator(tiltAngleInput, originPoint, edgePoint)


    except Exception as error:
        exception_type, exception_object, exception_traceback = sys.exc_info()
        filename = exception_traceback.tb_frame.f_code.co_filename
        line_number = exception_traceback.tb_lineno
        ui.messageBox(f"command_changed: {str(error)} ({filename}:{line_number})")

def command_execute(args: adsk.core.CommandEventArgs):
    pass
    #createAirfoilSketch(root, _filePass, originPoint, edgePoint, inputs.itemById('tiltAngle').value, inputs.itemById('edgeLineBool').value)

def command_preview(args: adsk.core.CommandEventArgs):
    try:
        inputs = args.command.commandInputs
        des = adsk.fusion.Design.cast(app.activeProduct)
        root = des.rootComponent
        methodSelectButton : adsk.core.ButtonRowCommandInput = inputs.itemById('MethodSelectButton') 
        if methodSelectButton.selectedItem.name == 'Points':
            originPointEntity = inputs.itemById('originPoint').selection(0).entity
            edgePointEntity = inputs.itemById('edgePoint').selection(0).entity
            sketch = createAirfoilSketchByPoints(root, _filePass, originPointEntity, edgePointEntity, inputs.itemById('tiltAngle').value, inputs.itemById('edgeLineBool').value)
        else:
            axisEntity = inputs.itemById('axis').selection(0).entity
            sketch = createAirfoilSketchByLine(root, _filePass, axisEntity, inputs.itemById('tiltAngle').value, inputs.itemById('edgeLineBool').value, _swapDirection)
        args.isValidResult = sketch is not None
    except Exception as error:
        exception_type, exception_object, exception_traceback = sys.exc_info()
        filename = exception_traceback.tb_frame.f_code.co_filename
        line_number = exception_traceback.tb_lineno
        ui.messageBox(f"command_preview: {str(error)} ({filename}:{line_number})")

def command_destroy(args: adsk.core.CommandEventHandler):
    des = adsk.fusion.Design.cast(app.activeProduct)
    #attribs = des.attributes
    
    global local_handlers
    local_handlers = []
def createAirfoilSketchByLine(comp: adsk.fusion.Component, filePass: str, axis: adsk.core.Base, tiltAngle: float, connectTrailingEdge: bool, swapDirection: bool):
    try:
        app = adsk.core.Application.get()
        viewport = app.activeViewport
        upVector = viewport.frontUpDirection
        root = adsk.fusion.Design.cast(app.activeProduct).rootComponent
        sketches = comp.sketches
        constructionPlanes = comp.constructionPlanes   
        if axis.objectType == 'adsk::fusion::SketchLine':
            originPoint = axis.startSketchPoint
            edgePoint = axis.endSketchPoint
            vectorToOrigin = adsk.core.Vector3D.create(*[(a - b) for a, b in zip(originPoint.worldGeometry.asArray(), edgePoint.worldGeometry.asArray())])
        else:
            originPoint = axis.startVertex
            edgePoint = axis.endVertex
            vectorToOrigin = adsk.core.Vector3D.create(*[(a - b) for a, b in zip(originPoint.geometry.asArray(), edgePoint.geometry.asArray())])
        if upVector.isParallelTo(adsk.core.Vector3D.create(0, 1, 0)) ^ upVector.isParallelTo(vectorToOrigin):
            plarnarPlane = root.xZConstructionPlane
        else:
            plarnarPlane = root.xYConstructionPlane
        constructionPlaneInput = constructionPlanes.createInput()
        planeAngle = math.pi / 2.0 - tiltAngle
        if swapDirection:
            originPoint, edgePoint = edgePoint, originPoint
            planeAngle = -planeAngle
        constructionPlaneInput.setByAngle(axis, adsk.core.ValueInput.createByReal(planeAngle), plarnarPlane)
        constructionPlane = constructionPlanes.add(constructionPlaneInput)
        sketch = sketches.add(constructionPlane)
        createAirfoilSketch(sketch, filePass, originPoint, edgePoint, tiltAngle, connectTrailingEdge)
        return sketch
    except Exception as error:
        exception_type, exception_object, exception_traceback = sys.exc_info()
        filename = exception_traceback.tb_frame.f_code.co_filename
        line_number = exception_traceback.tb_lineno
        ui.messageBox(f"createAirfoilSketch (byLine): {str(error)} ({filename}:{line_number})")

def createAirfoilSketchByPoints(comp: adsk.fusion.Component, filePass: str, originPoint: adsk.core.Base, edgePoint: adsk.core.Base, tiltAngle: float, connectTrailingEdge: bool):
    try:
        app = adsk.core.Application.get()
        viewport = app.activeViewport
        upVector = viewport.frontUpDirection
        root = adsk.fusion.Design.cast(app.activeProduct).rootComponent
        constructionAxes = comp.constructionAxes
        constructionAxesInput = constructionAxes.createInput()
        constructionAxesInput.setByTwoPoints(originPoint, edgePoint)
        constructionAxis = constructionAxes.add(constructionAxesInput)
        constructionPlanes = comp.constructionPlanes
        if upVector.isParallelTo(adsk.core.Vector3D.create(0, 1, 0)) ^ upVector.isParallelTo(constructionAxis.geometry.direction):
            plarnarPlane = root.xZConstructionPlane
        else:
            plarnarPlane = root.xYConstructionPlane
        constructionPlaneInput = constructionPlanes.createInput()
        planeAngle = math.pi / 2.0 - tiltAngle
        constructionPlaneInput.setByAngle(constructionAxis, adsk.core.ValueInput.createByReal(planeAngle), plarnarPlane)
        constructionPlane = constructionPlanes.add(constructionPlaneInput)
        constructionAxis.isLightBulbOn = False
        sketch = comp.sketches.add(constructionPlane)
        createAirfoilSketch(sketch, filePass, originPoint, edgePoint, tiltAngle, connectTrailingEdge)
        return sketch
    except Exception as error:
        exception_type, exception_object, exception_traceback = sys.exc_info()
        filename = exception_traceback.tb_frame.f_code.co_filename
        line_number = exception_traceback.tb_lineno
        ui.messageBox(f"createAirfoilSketch (byPoints): {str(error)} ({filename}:{line_number})")
    return None

def createAirfoilSketch(sketch: adsk.fusion.Sketch, filePass: str, originPoint: adsk.core.Base, edgePoint: adsk.core.Base, tiltAngle: float, connectTrailingEdge: bool):
    try:
        app = adsk.core.Application.get()
        viewport = app.activeViewport
        upVector = viewport.frontUpDirection
        frontEyeVector = viewport.frontEyeDirection
        sketchPoints = sketch.sketchPoints
        sketchCurves = sketch.sketchCurves
        sketchOriginPoint = sketch.project2([originPoint], True)[0]
        sketchEdgePoint = sketch.project2([edgePoint], True)[0]
        vectorToOrigin = adsk.core.Vector3D.create(*[(a - b) for a, b in zip(sketchOriginPoint.worldGeometry.asArray(), sketchEdgePoint.worldGeometry.asArray())])
        if upVector.isParallelTo(vectorToOrigin):
            frontEyeVector.scaleBy(-upVector.dotProduct(vectorToOrigin))
            normalVector = frontEyeVector.crossProduct(vectorToOrigin)
            sketchYVector = normalVector.crossProduct(vectorToOrigin)
        else:
            normalVector = vectorToOrigin.crossProduct(upVector)
            sketchYVector = normalVector.crossProduct(vectorToOrigin)
        normalVector.normalize()
        sketchYVector.normalize()
        sketchYVector.scaleBy(math.cos(tiltAngle))
        normalVector.scaleBy(math.sin(tiltAngle))
        sketchYVector.add(normalVector)
        sketchYVector.scaleBy(vectorToOrigin.length)
        sketchYPoint = sketchOriginPoint.worldGeometry
        sketchYPoint.translateBy(sketchYVector)
        sketchYPoint = sketch.modelToSketchSpace(sketchYPoint)
        matrix = adsk.core.Matrix3D.create()
        matrix.setWithCoordinateSystem(sketchOriginPoint.geometry , adsk.core.Vector3D.create(*[(b - a) for a, b in zip(sketchOriginPoint.geometry.asArray(), sketchEdgePoint.geometry.asArray())]), adsk.core.Vector3D.create(*[(a - b) for a, b in zip(sketchYPoint.asArray(), sketchOriginPoint.geometry.asArray())]), adsk.core.Vector3D.create())
        baseXLinePoint = adsk.core.Point3D.create(-0.1, 0, 0)
        baseYLinePoint = adsk.core.Point3D.create(0, -0.1, 0)
        baseOriginPoint = adsk.core.Point3D.create(-0.1, -0.1, 0)
        baseXLinePoint.transformBy(matrix)
        baseYLinePoint.transformBy(matrix)
        baseOriginPoint.transformBy(matrix)
        baseXLine = sketchCurves.sketchLines.addByTwoPoints(baseXLinePoint, baseOriginPoint)
        baseYLine = sketchCurves.sketchLines.addByTwoPoints(baseYLinePoint, baseXLine.endSketchPoint)
        chordLine = sketchCurves.sketchLines.addByTwoPoints(baseXLine.startSketchPoint, sketchOriginPoint)
        baseXLine.isConstruction = True
        baseYLine.isConstruction = True
        chordLine.isConstruction = True
        constraints = sketch.geometricConstraints
        constraints.addPerpendicular(baseXLine, baseYLine)
        constraints.addPerpendicular(baseXLine, chordLine)
        constraints.addEqual(baseXLine, baseYLine)
        constraints.addCoincident(sketchEdgePoint, chordLine)
        airfoilPoints = []
        with io.open(filePass, 'r', encoding='utf-8-sig') as f:
            fileExtension = os.path.splitext(filePass)[1]
            if fileExtension.lower() == '.dat':
                numBlank = 0
                line = f.readline()
                while line:
                    if line.strip() == '':
                        numBlank += 1
                    line = f.readline()
                if numBlank == 0:
                    f.seek(0)
                    f.readline()
                    line = f.readline()
                    while line:
                        pntStrArr = line.split()
                        if len(pntStrArr) >= 2:
                            try:
                                airfoilPointX = float(pntStrArr[0])
                                airfoilPointY = float(pntStrArr[1])
                                airfoilPoints.append((airfoilPointX, airfoilPointY))
                            except:
                                pass
                        line = f.readline()
                elif numBlank == 2:
                    f.seek(0)
                    f.readline()
                    line = f.readline()
                    f.readline()
                    upSurfacePointsNum, downSurfacePointsNum = line.split()
                    upSurfacePointsNum = int(float(upSurfacePointsNum))
                    downSurfacePointsNum = int(float(downSurfacePointsNum))
                    for i in range(upSurfacePointsNum):
                        line = f.readline()
                        pntStrArr = line.split()
                        if len(pntStrArr) >= 2:
                            try:
                                airfoilPointX = float(pntStrArr[0])
                                airfoilPointY = float(pntStrArr[1])
                                airfoilPoints.append((airfoilPointX, airfoilPointY))
                            except:
                                pass
                    f.readline()
                    line = f.readline()
                    pntStrArr = line.split()
                    if len(pntStrArr) >= 2:
                        try:
                            airfoilPointX = float(pntStrArr[0])
                            airfoilPointY = float(pntStrArr[1])
                            airfoilPoints.append((airfoilPointX, airfoilPointY))
                        except:
                            pass
                    if airfoilPoints[0] == airfoilPoints[-1]:
                        airfoilPoints = airfoilPoints[:-2]
                    airfoilPoints.reverse()

                    for i in range(downSurfacePointsNum - 1):     
                        line = f.readline()
                        pntStrArr = line.split()
                        if len(pntStrArr) >= 2:
                            try:
                                airfoilPointX = float(pntStrArr[0])
                                airfoilPointY = float(pntStrArr[1])
                                airfoilPoints.append((airfoilPointX, airfoilPointY))
                            except:
                                pass
            else:
                line = f.readline()
                while line:
                    pntStrArr = line.split(",")
                    if len(pntStrArr) >= 2:    
                        try:
                            airfoilPointX = float(pntStrArr[0])
                            airfoilPointY = float(pntStrArr[1])
                            airfoilPoints.append((airfoilPointX, airfoilPointY))
                        except:
                            pass
                    line = f.readline()
            f.close()
        minCoordinate = min(min([min(pnt) for pnt in airfoilPoints]), -0.1)
        dimenisions = sketch.sketchDimensions
        chordDimensionTextPos = adsk.core.Point3D.create(0.5, minCoordinate - 0.1, 0)
        baseXLineDimension1TextPos = adsk.core.Point3D.create(minCoordinate - 0.1, 0.1, 0)
        baseXLineDimension1TextPos.transformBy(matrix)
        baseXLineDimension2TextPos = adsk.core.Point3D.create(minCoordinate - 0.1, minCoordinate / 2.0, 0)
        baseXLineDimension2TextPos.transformBy(matrix)
        chordDimension = dimenisions.addDistanceDimension(sketchOriginPoint, sketchEdgePoint, adsk.fusion.DimensionOrientations.AlignedDimensionOrientation, chordDimensionTextPos, False)
        baseXLineDimension1 = dimenisions.addDistanceDimension(baseXLine.startSketchPoint, sketchOriginPoint, adsk.fusion.DimensionOrientations.AlignedDimensionOrientation, baseXLineDimension1TextPos)
        baseXLineDimension2 = dimenisions.addDistanceDimension(baseXLine.startSketchPoint, baseXLine.endSketchPoint, adsk.fusion.DimensionOrientations.AlignedDimensionOrientation, baseXLineDimension2TextPos)
        baseXLineDimension1.parameter.expression = f'{-minCoordinate} * {chordDimension.parameter.name}'
        baseXLineDimension2.parameter.expression = f'{-minCoordinate} * {chordDimension.parameter.name}'
        sketch.isComputeDeferred = True
        points = adsk.core.ObjectCollection.create()
        for i in range(len(airfoilPoints) - 1):
            point = adsk.core.Point3D.create(airfoilPoints[i][0], airfoilPoints[i][1], 0)
            point.transformBy(matrix)
            points.add(point)
        point = adsk.core.Point3D.create(airfoilPoints[-1][0] + (airfoilPoints[0] == airfoilPoints[-1]), airfoilPoints[-1][1], 0)
        point.transformBy(matrix)
        points.add(point)
        spline = sketchCurves.sketchFittedSplines.add(points)
        for i in range(points.count):
            splinePointXDimension = dimenisions.addOffsetDimension(baseXLine, spline.fitPoints.item(i), adsk.core.Point3D.create(0, 0, 0), True)
            splinePointYDimension = dimenisions.addOffsetDimension(baseYLine, spline.fitPoints.item(i), adsk.core.Point3D.create(0, 0, 0), True)
            splinePointXDimension.parameter.expression = f'{(airfoilPoints[i][0] - minCoordinate)} * {chordDimension.parameter.name}'
            splinePointYDimension.parameter.expression = f'{(airfoilPoints[i][1] - minCoordinate)} * {chordDimension.parameter.name}'

        if connectTrailingEdge and airfoilPoints[0] != airfoilPoints[-1]:
            sketchCurves.sketchLines.addByTwoPoints(spline.fitPoints.item(0), spline.fitPoints.item(spline.fitPoints.count - 1))
        sketch.isComputeDeferred = False
        return sketch
    except Exception as error:
        exception_type, exception_object, exception_traceback = sys.exc_info()
        filename = exception_traceback.tb_frame.f_code.co_filename
        line_number = exception_traceback.tb_lineno
        ui.messageBox(f"createAirfoilSketch (byPoints) : {str(error)} ({filename}:{line_number})")
    return None

def command_validate_inputs(args: adsk.core.ValidateInputsEventArgs):
    if _filePass is None:
        args.areInputsValid = False
        return
    args.areInputsValid = True
def setAngleManipulator(input: adsk.core.AngleValueCommandInput, originPoint: adsk.core.Point3D, edgePoint: adsk.core.Point3D):
    try:
        app = adsk.core.Application.get()
        viewport = app.activeViewport
        upVector = viewport.frontUpDirection
        frontEyeVector = viewport.frontEyeDirection
        vectorToOrigin = adsk.core.Vector3D.create(*[(a - b) for a, b in zip(originPoint.asArray(), edgePoint.asArray())])
        if upVector.isParallelTo(vectorToOrigin):
            frontEyeVector.scaleBy(-upVector.dotProduct(vectorToOrigin))
            horizontalVector = frontEyeVector.crossProduct(vectorToOrigin)
            verticalVector = horizontalVector.crossProduct(vectorToOrigin)
        else:
            horizontalVector = vectorToOrigin.crossProduct(upVector)
            verticalVector = horizontalVector.crossProduct(vectorToOrigin)
        horizontalVector.normalize()
        verticalVector.normalize()
        input.setManipulator(adsk.core.Point3D.create(*[(a + b) / 2.0 for a, b in zip(originPoint.asArray(), edgePoint.asArray())]), verticalVector, horizontalVector)
    except Exception as error:
        exception_type, exception_object, exception_traceback = sys.exc_info()
        filename = exception_traceback.tb_frame.f_code.co_filename
        line_number = exception_traceback.tb_lineno
        ui.messageBox(f"setAngleManipulator: {str(error)} ({filename}:{line_number})")