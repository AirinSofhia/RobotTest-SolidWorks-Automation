*** Settings ***
Suite Setup       Open App and prepare Data    SaveAs
Suite Teardown    Close App
Test Template     UsageControlForSaveAs
Resource          Public/public data.robot
Resource          Public/public operation.robot
Resource          Open_File/open_file_operation.robot
Resource          Run_Macros/run_macros_operation.robot
Library           DataDriver    file=testcase.xlsx    sheet_name=SaveAs    include=1    exclude=0

*** Test Cases ***
Run Solidworks Usage Control for Save As
    default    default    default    default    default    default    default    default    default

*** Keywords ***
UsageControlForSaveAs
    [Arguments]    ${test_data}    ${sleepTime}    ${OpenInNewWindows}    ${is_edit}    ${macro_edit}    ${expected_result}    ${macro_saveas}    ${verify_deny_img}    ${verify_saveas_files}
    [Tags]    default
    Sleep    2s
    #open file
    Load Test Data    ${test_data}
    Sleep    ${sleepTime}
    IF    '${OpenInNewWindows}' == 'yes'
        Open TestPrt InNewWindows
        Sleep    2s
    END
    #if is_edit = yes,then edit the current file, else, not edit.
    IF    '${is_edit}' == 'yes'
        Load Macro    ${macro_edit}
    END
    Sleep    1s
    ClosewindowsForDenyMessage    ${expected_result}
    Sleep    2s
    IsGoBackToRoot   ${OpenInNewWindows}
    #saveas the file
    IF  '${expected_result}' == 'deny'
        TriggerSaveAsFileByKeyBoard
    ELSE
        Load Macro   ${macro_saveas}
    END
    Sleep    1s
    #get the result
    ${result}    ${msg}    Verify Result For Solidworks SaveAs    ${expected_result}    ${CURDIR}\\${verify_deny_img}    ${verify_saveas_files}
    #Close deny message and opened files
    ClosewindowsForDenyMessage    ${expected_result}
    CloseFileWithoutSave
    Sleep    1s
    ClosewindowsForDenyMessage    ${expected_result}
    CloseFileWithoutSave
    Empty Folder    folder_path=C:\\SolidworksUsageControl\\SaveAs\\SaveAsFiles
    IF    ${result} == ${False} or ${result} == ${None}
        Fail    "Case Failed,"
        Log To Console  ${msg}
    END

