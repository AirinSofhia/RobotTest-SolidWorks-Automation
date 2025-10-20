*** Settings ***
Library           ../Library/utility.py

*** Keywords ***
Trigger the Macro Dialog
    utility.Presskey    k1=f12

Input Macro Path In Dialog
    [Arguments]    ${macro_path}
    Sleep    1s
    Type Path    path=${macro_path}
    Sleep    1s

# Click Enter
#     utility.Presskey    k1=enter
