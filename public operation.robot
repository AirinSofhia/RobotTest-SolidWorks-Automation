*** Settings ***
Library           ../Library/utility.py
Library           FlaUILibrary
Resource          ../Public/public data.robot
Resource          ../Public/public_file_element.resource
Resource          ../Open_File/open_file_element.resource

*** Keywords ***
Open App
    Launch Application    ${PROGRAM_PATH}
    Sleep    5s
    Attach Application By Name    ${PROGRAM_NAME}
    Sleep    15s

Open App and prepare Data
    [Arguments]    ${action}
    Restore Data    ${action}
    Open App


Close App
    Close Application By Name    ${PROGRAM_NAME}
    Sleep    15s
    # utility.Presskey    k1=f4
    # Sleep    2s
    # utility.Presskey    k1=left
    # utility.Presskey    k1=enter

CloseFileWithoutSave
    CloseFile
    Sleep    1s
    # MoveToCenterAndClick
    # Sleep    1s
    ExitWithoutSave

CloseAllFileWithoutSave
    Close All File
    Sleep    1s
    ExitWithoutSave

ClosewindowsForDenyMessage
    [Arguments]    ${expected_result}
    IF  "${expected_result}" == "deny"
        ClickEnter
    END

Save File To Output Folder
    [Arguments]    ${sleep_time}
    TriggerSaveAsFileByKeyBoard    ${sleep_time}
    Input Output File Path In Dialog
    ClickEnter
    Sleep    ${sleep_time}

TriggerSaveAsFileByKeyBoard
    KeyboardForSaveAs

CloseNoEditPermissionDialog
    Click Img    ${CURDIR}\\..\\VerifyImg\\CommonAction\\CloseNoEditPermissionWindows\\XbuttonForNextLabs.PNG
	
FileOpenToImport
    [Arguments]    ${data_path}    ${open_type}    ${sleep_time}    ${different_action}
    ActiveWindows
    Trigger the open file dialog
    GoToDefaultPath
    Input File Path In Dialog    ${data_path}
    Click Open Button From Dialog
    Sleep    10s
    Open File    ${open_type}    ${sleep_time}

Close_FileOpenToImport
    [Arguments]    ${expected_result}    ${different_close_action}
    CloseWindowsForDenyMessage    ${expected_result}
    CloseWindowsForDenyMessage    ${expected_result}
    ClickCancelButton
    CloseFileWithoutSave

AsmInsertToImport
    [Arguments]    ${data_path}    ${open_type}    ${sleep_time}    ${different_action}
    ActiveWindows
    Trigger the open file dialog
    GoToDefaultPath
    GoToAsmPath
    Sleep    3s
    ClickInsertButton
    ClickComponentButton
    ClickExistingButton
    GoToDefaultPath
    Sleep    3s
    Input File Path In Dialog    ${data_path}
    Select All File    ${different_action}    ${data_path}
    Click Open Button From Dialog
    Open File    ${open_type}    ${sleep_time}
    Sleep    10s
    
    
Close_AsmInsertToImport
    [Arguments]    ${expected_result}    ${different_close_action}
    CloseWindowsForDenyMessage    ${expected_result}
    ClickCancelButton
    CloseWindowsForDenyMessage    ${expected_result}
    ClickGreenTick
    CloseFileWithoutSave


PrtInsertFeaturesToImport
    [Arguments]    ${data_path}    ${open_type}    ${sleep_time}    ${different_action}
    ActiveWindows
    Trigger the open file dialog
    GoToDefaultPath
    GoToPrtPath
    Sleep    3s
    ClickInsertButton
    ClickFeaturesButton
    ClickImportedButton
    GoToDefaultPath
    Sleep    3s
    Input File Path In Dialog    ${data_path}
    Select All File    ${different_action}    ${data_path}
    Click Open Button From Dialog
    Open File    ${open_type}    ${sleep_time}
    Sleep    10s
    ZoomToFit
    
Close_PrtInsertFeaturesToImport
    [Arguments]    ${expected_result}    ${different_close_action}
    CloseWindowsForDenyMessage    ${expected_result}
    ClickCancelButton
    CloseFileWithoutSave

PrtInsertPartToImport
    [Arguments]    ${data_path}    ${open_type}    ${sleep_time}    ${different_action}
    ActiveWindows
    Trigger the open file dialog
    GoToDefaultPath
    GoToPrtPath
    Sleep    3s
    ClickInsertButton
    ClickPartButton
    GoToDefaultPath
    Sleep    3s
    Input File Path In Dialog    ${data_path}
    Select All File    ${different_action}    ${data_path}
    Click Open Button From Dialog
    Open File    ${open_type}    ${sleep_time}
    Sleep    10s
   

Close_PrtInsertPartToImport
    [Arguments]    ${expected_result}    ${different_close_action}
    CloseWindowsForDenyMessage    ${expected_result}
    ClickCancelButton
    ClickGreenTick
    CloseFileWithoutSave

ToolbarsBlocksInsertToImport
    [Arguments]    ${data_path}    ${open_type}    ${sleep_time}    ${different_action}
    ActiveWindows
    Trigger the open file dialog
    GoToDefaultPath
    GoToPrtPath
    Sleep    3s
    ClickToolsButton
    ClickBlocksButton
    SelectInsertButton
    ClickFrontPlane
    ClickBrowseButton
    GoToDefaultPath
    Input File Path In Dialog    ${data_path}
    Select All File    ${different_action}    ${data_path}
    Click Open
    Open File    ${open_type}    ${sleep_time}
    ZoomToFit
    Sleep    10s

Close_ToolbarsBlocksInsertToImport
    [Arguments]    ${expected_result}    ${different_close_action}
    CloseWindowsForDenyMessage    ${expected_result}
    ClickCancelButton
    ClickGreenTick
    CloseFileWithoutSave


DrwInsertPictureToImport
    [Arguments]    ${data_path}    ${open_type}    ${sleep_time}    ${different_action}
    ActiveWindows
    Trigger the open file dialog
    GoToDefaultPath
    GoToDRWPath
    Sleep    3s
    ClickInsertButton
    ClickPictureButton
    GoToDefaultPath
    Sleep    3s
    Input File Path In Dialog    ${data_path}
    Click Open Button From Dialog
    Open File    ${open_type}    ${sleep_time}
    
    Sleep    10s
    
Close_DrwInsertPictureToImport
    [Arguments]    ${expected_result}    ${different_close_action}
    CloseWindowsForDenyMessage    ${expected_result}
    ClickCancelButton
    CloseFileWithoutSave

DrwInsertDXF/DWGToImport
    [Arguments]    ${data_path}    ${open_type}    ${sleep_time}    ${different_action}
    ActiveWindows
    Trigger the open file dialog
    GoToDefaultPath
    GoToDRWPath
    Sleep    3s
    ClickInsertButton
    ClickDxfDwgButton
    GoToDefaultPath
    Sleep    3s
    Input File Path In Dialog    ${data_path}
    Select All File    ${different_action}    ${data_path}
    Click Open Button From Dialog
    Open File    ${open_type}    ${sleep_time}
    Sleep    3s
    ClickGreenTick
    ZoomToFit
    Sleep    10s

Close_DrwInsertDXF/DWGToImport
    [Arguments]    ${expected_result}    ${different_close_action}
    CloseWindowsForDenyMessage    ${expected_result}
    ClickCancelButton
    CloseFileWithoutSave