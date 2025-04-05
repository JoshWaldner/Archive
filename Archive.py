import adsk.core, adsk.fusion
import traceback
import os

app = adsk.core.Application.get()
ui = app.userInterface

# Variables
CMD_ID = "Archive"
CMD_NAME = "Archive"
CMD_Description = "Archive Fusion Documents in .step and .f3d formats"
QATToolbar = ui.toolbars.itemById("QAT")
ICON_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), "resources", "")
FileDirectory = "C://Users//joshua//OneDrive - Spring Prairie Colony//Documents//Fusion 360//Fusion Archives//"
handlers = []

def run(context):

    InitAddIn()

def stop(context):

    cmdDef = ui.commandDefinitions.itemById(CMD_ID)
    if cmdDef:
        cmdDef.deleteMe()

    cmd = QATToolbar.controls.itemById(CMD_ID)
    if cmd:
        cmd.deleteMe()

def InitAddIn():
    try:
        # pass

        
        BTTNdef = ui.commandDefinitions.addButtonDefinition(CMD_ID, CMD_NAME, CMD_Description, ICON_FOLDER)

        commandEventHandler = CommandEventHandler()
        BTTNdef.commandCreated.add(commandEventHandler)
        handlers.append(commandEventHandler)

        QATToolbar.controls.addCommand(BTTNdef, "FusionDrawingFeatureToggleCommand", False)
    except:
        ui.messageBox("Failed:\n{}".format(traceback.format_exc()))


class CommandEventHandler(adsk.core.CommandCreatedEventHandler):
    def __init__(self):
        super().__init__()

    def notify(self, eventArgs):
        try:
            global handlers
            EventArgs = adsk.core.CommandCreatedEventArgs.cast(eventArgs)
            cmd = EventArgs.command

            onExecute = MyExecuteHandler()
            cmd.execute.add(onExecute)
            handlers.append(onExecute)
        except:
            ui.messageBox("Failed:\n{}".format(traceback.format_exc()))



class MyExecuteHandler(adsk.core.CommandEventHandler):
    def __init__(self):
        super().__init__()

    def notify(self, eventArgs):
        try:

            doc = app.activeDocument
            Product = app.activeProduct
            Design = adsk.fusion.Design.cast(Product)

            

            RootComp = Design.rootComponent
            if RootComp.allOccurrences.count == 0 and RootComp.bRepBodies.count == 0:
                FileDialog = ui.createFileDialog()
                FileDialog.initialDirectory = FileDirectory
                if FileDialog.showOpen() == adsk.core.DialogResults.DialogOK:
                    try:
                        impMgr = app.importManager
                        f3dImportOptions = impMgr.createFusionArchiveImportOptions(FileDialog.filename)
                        doc = impMgr.importToNewDocument(f3dImportOptions)
                    except:
                        impMgr = app.importManager
                        f3dImportOptions = impMgr.createSTEPImportOptions(FileDialog.filename)
                        doc = impMgr.importToNewDocument(f3dImportOptions)                        
            else:
                FileName = ui.inputBox("Give the file a name", "Name")
                if FileName[0] == "":
                    Filename = doc.name
                else:
                    Filename = FileName[0]


                if FileName[1] == False:

                
                    Export = Design.exportManager.createFusionArchiveExportOptions(f"{FileDirectory}{Filename}")
                    Design.exportManager.execute(Export)
                    Export = Design.exportManager.createSTEPExportOptions(f"{FileDirectory}{Filename}")
                    Design.exportManager.execute(Export)

                    doc.close(False)
        except:
            ui.messageBox("Failed:\n{}".format(traceback.format_exc()))