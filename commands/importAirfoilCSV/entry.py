import adsk.core
import os

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
    
    futil.add_handler(args.command.execute, command_execute, local_handlers=local_handlers)
    futil.add_handler(args.command.executePreview, command_preview, local_handlers=local_handlers)
    futil.add_handler(args.command.destroy, command_destroy, local_handlers=local_handlers)

    inputs.addBoolValueInput('fileSelectButton', 'CSV Select', False, '', False)
    inputs.addTextBoxCommandInput('fileNameText', 'CSV name', 'Selected CSV file name', 1, True)
    
def command_execute(args: adsk.core.CommandEventArgs):
    inputs = args.command.commandInputs
    des = adsk.fusion.Design.cast(app.activeProduct)
    root = des.rootComponent

def command_preview(args: adsk.core.CommandEventArgs):
    des = adsk.fusion.Design.cast(app.activeProduct)
    root = des.rootComponent
    inputs = args.command.commandInputs
    
def command_destroy(args: adsk.core.CommandEventHandler):
    des = adsk.fusion.Design.cast(app.activeProduct)
    attribs = des.attributes