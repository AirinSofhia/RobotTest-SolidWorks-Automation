*** Settings ***
Suite Setup       Open App and prepare Data    Open
Suite Teardown    Close App
Test Template     UsageControlForOpen
Resource          Public/public data.robot
Resource          Public/public operation.robot
Resource          Open_File/open_file_operation.robot
Library           DataDriver    file=testcase.xlsx    sheet_name=Open    include=1    exclude=0


*** Test Cases ***
Run Solidworks Usage Control for Open
    default    default    default


*** Keywords ***
UsageControlForOpen
    [Arguments]    ${test_data}    ${sleepTime}    ${verifyImg}
    [Tags]    default
    Sleep    5s
    Load Test Data    ${test_data}
    Sleep    ${sleepTime}
    ${result}    Verify Img    ${CURDIR}\\${verifyImg}
    ClickEnter
    Sleep    1s
    ClickEnter
    CloseFileWithoutSave
    IF    "${result}" == "${None}"
        Fail    "Case Failed, fail to find the verifyImg"
    END
