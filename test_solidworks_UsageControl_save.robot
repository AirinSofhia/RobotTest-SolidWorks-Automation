*** Settings ***
Suite Setup       Open App and prepare Data    Save
Suite Teardown    Close App
Test Template     UsageControlForSave
Resource          Public/public data.robot
Resource          Public/public operation.robot
Resource          Open_File/open_file_operation.robot
Resource          Run_Macros/run_macros_operation.robot
Library           DataDriver    file=testcase.xlsx    sheet_name=Save    include=1    exclude=0

*** Test Cases ***
Run Solidworks Usage Control for Save
    default    default    default    default    default    default    default    default

*** Keywords ***
UsageControlForSave
    [Arguments]    ${test_data}    ${sleepTime}    ${OpenInNewWindows}    ${macro_edit}    ${macro_save}    ${expected_result}    ${verify_deny_img}    ${verify_allow_files}
    [Tags]    default
    #get old file info
    ${OldFileInfo}    Evaluate    collections.OrderedDict()    modules=collections
    IF    '${expected_result}' == 'allow'
        ${OldFileInfo}    Getfileinfo    ${verify_allow_files}
    END
    #open file
    Load Test Data    ${test_data}
    Sleep    ${sleepTime}
    Sleep    5s
    IF    '${OpenInNewWindows}' == 'yes'
        Open TestPrt InNewWindows
    END
    Sleep    2s
    #edit the file
    Load Macro    ${macro_edit}
    #sometime it will pop info message, click enter to close pop windows
    Sleep    1s
    CloseNoEditPermissionDialog
    Sleep    3s
    IF    '${OpenInNewWindows}' == 'yes'
        Go Back To asm
        Sleep    1s
        CloseNoEditPermissionDialog
        Sleep    1s
        ActiveWindows
    END
    #save the file
    Load Macro    ${macro_save}
    Sleep    1s
    #get the result
    ${result}    ${msg}    Verify Result For Solidworks Save    ${expected_result}    ${CURDIR}\\${verify_deny_img}    ${verify_allow_files}    ${OldFileInfo}
    #Close deny message
    ClosewindowsForDenyMessage    ${expected_result}
    #try to close opened files
    CloseFile
    Sleep    1s
    CloseFileWithoutSave
    IF    ${result} == ${False} or ${result} == ${None}
        Fail    "Case Failed,"
        Log To Console    ${msg}
    END

