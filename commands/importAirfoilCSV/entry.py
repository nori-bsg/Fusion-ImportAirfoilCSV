import adsk.core
import os
import io

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

def start():
    futil.log(f'{CMD_NAME} Command started')
    cmdDef = ui.commandDefinitions.itemById(CMD_ID)
    if cmdDef:
        cmdDef.deleteMe()
    
    icon_folder = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'resources', '')

    cmdDef = ui.commandDefinitions.addButtonDefinition(CMD_ID, CMD_NAME, CMD_Description, icon_folder)

    futil.add_handler(cmdDef.commandCreated, command_created)

    workspace = ui.workspaces.itemById(WORKSPACE_ID)

    panel = workspace.toolbarPanels.itemById(PANEL_ID)

    control = panel.controls.addCommand(cmdDef, COMMAND_BESIDE_ID, False)

    control.isPromoted = IS_PROMOTED

def stop():
    workspace = ui.workspaces.itemById(WORKSPACE_ID)
    panel = workspace.toolbarPanels.itemById(PANEL_ID)
    command_control = panel.controls.itemById(CMD_ID)
    command_definition = ui.commandDefinitions.itemById(CMD_ID)

    if command_control:
        command_control.deleteMe()

    if command_definition:
        command_definition.deleteMe()

def command_created(args: adsk.core.CommandCreatedEventArgs):
    inputs = args.command.commandInputs
    
    futil.add_handler(args.command.inputChanged, command_changed, local_handlers=local_handlers)
    futil.add_handler(args.command.execute, command_execute, local_handlers=local_handlers)
    futil.add_handler(args.command.executePreview, command_preview, local_handlers=local_handlers)
    futil.add_handler(args.command.destroy, command_destroy, local_handlers=local_handlers)

    inputs.addBoolValueInput('fileSelectButton', 'Select CSV or DAT', False, '', False)
    inputs.addTextBoxCommandInput('fileNameText', 'File Name', 'Selected file name', 1, True)
    inputs.addTextBoxCommandInput('filePassText', 'File Pass', 'Selected file pass', 1, True)
    inputs.addBoolValueInput('edgeLineBool', 'Connect Trailing Edge', True, '', True)
    originPointInput = inputs.addSelectionInput('originPoint', 'Origin Point', 'Select Origin Point')
    originPointInput.addSelectionFilter('Vertices')
    originPointInput.addSelectionFilter('SketchPoints')
    originPointInput.addSelectionFilter('ConstructionPoints')
    edgePointInput = inputs.addSelectionInput('edgePoint', 'Trailing Edge Point', 'Select Trailing Edge Point')
    edgePointInput.addSelectionFilter('Vertices')
    edgePointInput.addSelectionFilter('SketchPoints')
    edgePointInput.addSelectionFilter('ConstructionPoints')
    inputs.addAngleValueCommandInput('tiltAngle', 'Tilt Angle', adsk.core.ValueInput.createByReal(0)).isVisible = False
    
def command_changed(args: adsk.core.InputChangedEventArgs):
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
        args.inputs.itemById('filePassText').text = filepass
        args.inputs.itemById('fileNameText').text = filename
    elif input.id == 'originPoint' or input.id =='edgePoint':
        tiltAngleInput :adsk.core.AngleValueCommandInput = inputs.itemById('tiltAngle')
        
        if inputs.itemById('originPoint').selectionCount == 1 and inputs.itemById('edgePoint').selectionCount == 1:
            tiltAngleInput.isVisible = True
            originPointEntity = inputs.itemById('originPoint').selection(0).entity
            if originPointEntity.objectType == 'adsk::core::SketchPoint':
                originPoint = originPointEntity.worldGeometry
            else:
                originPoint = originPointEntity.geometry
            edgePointEntity = inputs.itemById('edgePoint').selection(0).entity
            if edgePointEntity.objectType == 'adsk::core::SketchPoint':
                edgePoint = edgePointEntity.worldGeometry
            else:
                edgePoint = edgePointEntity.geometry
            matrix = adsk.core.Matrix3D.create()
            matrix.setToRotation(math.pi / 2.0, adsk.core.Vector3D.create(0, 0, 1), edgePoint)
            horizontalVector = adsk.core.Vector3D.create(*[(b - a) for a, b in zip(originPoint.asArray(), edgePoint.asArray())])
            horizontalVector.z = 0
            horizontalVector.transformBy(matrix)
            horizontalVector.normalize()
            verticalVector = horizontalVector.crossProduct(adsk.core.Vector3D.create(*[(a - b) for a, b in zip(originPoint.asArray(), edgePoint.asArray())]))
            verticalVector.normalize()
            tiltAngleInput.setManipulator(adsk.core.Point3D.create(*[(a + b) / 2.0 for a, b in zip(originPoint.asArray(), edgePoint.asArray())]), verticalVector, horizontalVector)

        else:
            tiltAngleInput.isVisible = False

def command_execute(args: adsk.core.CommandEventArgs):
    inputs = args.command.commandInputs
    des = adsk.fusion.Design.cast(app.activeProduct)
    root = des.rootComponent
    originPointEntity = inputs.itemById('originPoint').selection(0).entity
    if originPointEntity.objectType == 'adsk::core::SketchPoint':
        originPoint = originPointEntity.worldGeometry
    else:
        originPoint = originPointEntity.geometry
    edgePointEntity = inputs.itemById('edgePoint').selection(0).entity
    if edgePointEntity.objectType == 'adsk::core::SketchPoint':
        edgePoint = edgePointEntity.worldGeometry
    else:
        edgePoint = edgePointEntity.geometry
    createAirfoilSketch(root, inputs.itemById('filePassText').text, originPoint, edgePoint, inputs.itemById('tiltAngle').value, inputs.itemById('edgeLineBool').value)

def command_preview(args: adsk.core.CommandEventArgs):
    des = adsk.fusion.Design.cast(app.activeProduct)
    root = des.rootComponent
    
    try:
        inputs = args.command.commandInputs
        originPointEntity = inputs.itemById('originPoint').selection(0).entity
        if originPointEntity.objectType == 'adsk::core::SketchPoint':
            originPoint = originPointEntity.worldGeometry
        else:
            originPoint = originPointEntity.geometry
        edgePointEntity = inputs.itemById('edgePoint').selection(0).entity
        if edgePointEntity.objectType == 'adsk::core::SketchPoint':
            edgePoint = edgePointEntity.worldGeometry
        else:
            edgePoint = edgePointEntity.geometry
        sketch = createAirfoilSketch(root, inputs.itemById('filePassText').text, originPoint, edgePoint, inputs.itemById('tiltAngle').value, inputs.itemById('edgeLineBool').value)
    except Exception as error:
        ui.messageBox(str(error))

def command_destroy(args: adsk.core.CommandEventHandler):
    des = adsk.fusion.Design.cast(app.activeProduct)
    #attribs = des.attributes
    
    global local_handlers
    local_handlers = []

def createAirfoilSketch(comp: adsk.fusion.Component, filePass: str = None, originPoint: adsk.core.Point3D = None, edgePoint: adsk.core.Point3D = None, tiltAngle: float = 0.0, connectTrailingEdge: bool = False):
    try:
        sketches = comp.sketches
        sketch = sketches.add(comp.xYConstructionPlane)
        sketchCurves = sketch.sketchCurves
        matrix = adsk.core.Matrix3D.create()
        matrix.setToRotation(math.pi / 2.0, adsk.core.Vector3D.create(0, 0, 1), edgePoint)
        xAxis = adsk.core.Vector3D.create(*[(b - a) for a, b in zip(originPoint.asArray(), edgePoint.asArray())])
        zAxis = adsk.core.Vector3D.create(*[(b - a) for a, b in zip(originPoint.asArray(), edgePoint.asArray())])
        zAxis.z = 0
        zAxis.transformBy(matrix)
        zAxis.normalize()
        yAxis = zAxis.crossProduct(adsk.core.Vector3D.create(*[(a - b) for a, b in zip(originPoint.asArray(), edgePoint.asArray())]))
        zAxis.scaleBy(xAxis.length)
        matrix.setToRotation(-tiltAngle, xAxis, adsk.core.Point3D.create())
        yAxis.transformBy(matrix)
        zAxis.transformBy(matrix)
        matrix.setWithCoordinateSystem(originPoint, xAxis, yAxis, zAxis)
        with io.open(filePass, 'r', encoding='utf-8-sig') as f:
            points = adsk.core.ObjectCollection.create()
            line = f.readline()
            data = []
            while line:
                if "," in line :
                    pntStrArr = line.split(",")
                else:
                    pntStrArr = line.split()
                for pntStr in pntStrArr:
                    try:
                        data.append(float(pntStr))
                    except:
                        break
            
                if len(data) >= 2 :
                    point = adsk.core.Point3D.create(data[0], data[1], 0)
                    point.transformBy(matrix)
                    points.add(point)
                data.clear()    
                line = f.readline()        
        if points.count:
            spline = sketchCurves.sketchFittedSplines.add(points)
            if connectTrailingEdge:
                constratins = sketch.geometricConstraints
                line = sketchCurves.sketchLines.addByTwoPoints(points.item(0), points.item(points.count - 1))
                constratins.addCoincident(spline.startSketchPoint, line.startSketchPoint)
                constratins.addCoincident(spline.endSketchPoint, line.endSketchPoint)
        return sketch
    except Exception as error:
        ui.messageBox(str(error))
    return None