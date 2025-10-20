*** Settings ***
Resource          run_macros_element.resource
Resource          ../Public/public_file_element.resource

*** Keywords ***
Load Macro
    [Arguments]    ${macro_path}
    Trigger the Macro Dialog
    Input Macro Path In Dialog    ${macro_path}
    ClickEnter
    Sleep    1s
