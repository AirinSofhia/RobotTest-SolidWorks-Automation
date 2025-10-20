*** Settings ***
Suite Setup       Open App
Suite Teardown    Close App
Test Template     UsageControlForFileImport
Resource          Public/public data.robot
Resource          Public/public operation.robot
Resource          Open_File/open_file_operation.robot
Library           Collections
Library           String
Library           DataDriver  file=testcase.xlsx    sheet_name=file_import    include=1    exclude=0





*** Test Cases ***
Test Import
    ${test_data}     ${verify_file}    ${verify_deny_img}    ${open_type}    ${sleepTime}    ${expected_result}    ${different_action}    ${different_close_action}   # robotcode: ignore

*** Keywords ***
UsageControlForFileImport

    [Arguments]    ${test_data}     ${verify_file}    ${verify_deny_img}    ${open_type}    ${sleepTime}    ${expected_result}    ${different_action}    ${different_close_action}
    Log To Console    Executing Action: ${different_action}

    Run Keyword    ${different_action}    ${test_data}    ${open_type}    ${sleepTime}     ${different_action}


    # Determine Expected Result
    
    IF    '${expected_result}' == 'allow'
        ${result}    ${msg}    Run Allow Case    ${test_data}     ${verify_file}    ${sleepTime}    ${expected_result}
    ELSE IF    '${expected_result}' == 'deny'
        ${result}    ${msg}    Run Deny Case     ${test_data}   ${verify_deny_img}    ${sleepTime}    ${expected_result}    ${different_close_action}
    END
    # Verify and Log 
    
    IF    "${result}" == "False" or "${result}" == "None"
        Fail    "Case Failed: ${msg}"
    ELSE
        Log To Console   "Case Passed: ${msg}"    
    END

Run Allow Case
    [Arguments]       ${test_data}     ${verify_file}    ${sleepTime}    ${expected_result}
    Log To Console    Running Allow Case
    Save File To Output Folder    ${sleepTime}
    ${NEW_FILE}=    Get New File Name    ${OUTPUT_FOLDER}
    ${result}    ${msg}=    Verify File Size   ${NEW_FILE}    ${verify_file}
    CloseFileWithoutSave
    Empty Folder    ${OUTPUT_FOLDER}
    RETURN    ${result}    ${msg}

Run Deny Case
    [Arguments]    ${test_data}    ${verify_deny_img}    ${sleepTime}    ${expected_result}    ${different_close_action}
    Log To Console    Running Deny Case
    ${result}    ${msg}=    Verify Result    ${expected_result}    ${CURDIR}\\${verify_deny_img}
    Run Keyword    ${different_close_action}    ${expected_result}    ${different_close_action}
    RETURN    ${result}    ${msg}

