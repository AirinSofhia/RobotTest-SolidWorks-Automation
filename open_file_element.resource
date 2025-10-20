*** Settings ***
Library           ../Library/utility.py
Resource          ../Public/public data.robot
Resource          ../Public/public_file_element.resource

*** Keywords ***
Trigger the open file dialog
    Sleep    1s
    Press 2 Keys    k1=ctrl    k2=w
    ClickEnter
    # utility.Presskey    k1=enter
    Sleep    1s
    Press 2 Keys    k1=ctrl    k2=o

GoToDefaultPath
    Type Path    path=c:
    Sleep    2s
    ClickEnter

Input File Path In Dialog
    [Arguments]    ${data_path}
    Sleep    1s
    Type Path    path=${data_path}
    Sleep    1s

Click Open Button From Dialog
    ClickEnter

ActiveWindows
    Active Windows By Click
    Sleep    1s


MinimaziCurrentFile
    Click Img    img=${CURDIR}\\..\\VerifyImg\\CommonAction\\SwitchToRoot\\MinimizeButton_CurrentFile.PNG
    Sleep    1s
    ClickEnter

