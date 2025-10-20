*** Settings ***
Resource          open_file_element.resource
Library           ../Library/utility.py

*** Keywords ***
Load Test Data
    [Arguments]    ${data_path}
    ActiveWindows
    Trigger the open file dialog
    GoToDefaultPath
    Input File Path In Dialog    ${data_path}
    Click Open Button From Dialog
    Sleep    10s


Go Back To asm
    [Documentation]    go back to asm or drw.
    MinimaziCurrentFile
    ActiveWindows


Open TestPrt InNewWindows
    [Documentation]    for asm or drw, we need to open the test.prt in new windows to do edit action.
    @{ModelTree_list}    Create List    ${CURDIR}\\..\\VerifyImg\\CommonAction\\openTestPrtInNewWindows\\DrawingView1.PNG    ${CURDIR}\\..\\VerifyImg\\CommonAction\\openTestPrtInNewWindows\\asm_arduino.PNG
    ${result}    Verify Img    img=${CURDIR}\\..\\VerifyImg\\CommonAction\\openTestPrtInNewWindows\\DrawingView1.PNG
    IF  "${result}" != "${None}"
        FOR    ${img}     IN    @{ModelTree_list}
            utility.Doubleclick Img    ${img}
            Sleep    1s

        END
    END
    Rightclick Img    img=${CURDIR}\\..\\VerifyImg\\CommonAction\\openTestPrtInNewWindows\\TestPrt.PNG
    @{openbutton_list}    Create List    ${CURDIR}\\..\\VerifyImg\\CommonAction\\openTestPrtInNewWindows\\OpenPart_drw.PNG
    ...        ${CURDIR}\\..\\VerifyImg\\CommonAction\\openTestPrtInNewWindows\\OpenPart_asm.PNG
    FOR    ${img}     IN    @{openbutton_list}
        Click Img    ${img}
        Sleep    1s
    END
    Sleep    7s
    ActiveWindows

IsGoBackToRoot
    [Arguments]    ${OpenInNewWindows}
    IF   '${OpenInNewWindows}' == 'yes'
        Go Back To asm
        Sleep    1s
        ActiveWindows
    END