*** Settings ***
Library           ../Library/utility.py
Resource            ../Public/public data.robot


*** Keywords ***
Open Model 
    Sleep    1s
    Press 2 Keys    k1=ctrl    k2=o
    Sleep     1s


CloseFile
    Sleep    1s
    Press 2 Keys    k1=ctrl    k2=w
    Sleep    1s

Close All File
    Sleep    1s
    Press 3 Keys    k1=ctrl    k2=shift    k3=w
    Sleep    2s


ExitWithoutSave
    utility.Presskey    k1=down
    Sleep    1s
    utility.Presskey    k1=enter
    Sleep    1s

Restore Data
    [Arguments]    ${action}
    utility.Restore Data    ${action}


ClickEnter
    utility.Presskey    k1=enter

MoveToCenterAndClick
    utility.Move To Center By Click

KeyboardForSaveAs
    utility.Press 3 Keys    k1=ctrl    k2=shift    k3=s
	
####################import#################
ClickGreenTick 
    Click Img   img=${SCREEN_IMAGES_PATH}\\CommonAction\\OpenFile\\green_tick.png
    Sleep    3s
    
ClickCancelButton 
    Click Img   img=${SCREEN_IMAGES_PATH}\\CommonAction\\Cancel\\cancel_button.png 
    Sleep    1s

ZoomToFit    
    utility.Presskey    k1=f   
    
NoImportDiagnostic 
    utility.Presskey    k1=right
    Sleep    1s
    utility.Presskey    k1=enter
    Sleep    1s

NoFeatureRecogniton 
    utility.Presskey    k1=right
    Sleep    1s
    utility.Presskey    k1=enter
    Sleep    1s

Press Enter Based on Open Type
    [Arguments]    ${open_type}    ${sleep_time}
    IF  ${open_type} == 1  
        ClickEnter
        Sleep    ${Sleep_time}
    ELSE IF    ${open_type} == 2
        NoImportDiagnostic   
        Sleep    ${sleep_time}
    ELSE IF    ${open_type} == 3
        FOR    ${i}    IN RANGE    3
            Click Enter
            Sleep    ${sleep_time}
        END
    ELSE IF    ${open_type} == 4
        NoImportDiagnostic
        NoFeatureRecogniton   
        Sleep    ${sleep_time}
    ELSE IF    ${open_type} == 5
        ClickCancelButton
        MoveToCenterAndClick
        Click Enter  
        Sleep    ${sleep_time}
    ELSE IF    ${open_type} ==6
        FOR    ${i}    IN RANGE    2
            Click Enter
            Sleep    5s
        END
    
    END

Open File
    [Arguments]    ${open_type}    ${sleep_time}
    Sleep    ${sleep_time}
    Press Enter Based on Open Type    ${open_type}    ${sleep_time}
    Sleep    ${sleep_time}

Input Output File Path In Dialog
    Type Path    path=${OUTPUT_FOLDER}
    Sleep    1s
    ClickEnter
    Sleep    1s

Click Open
    Click Img   img=${SCREEN_IMAGES_PATH}\\CommonAction\\OpenFile\\open_button.png
    Sleep    1s

GoToAsmPath
    Type Path    path=${ASM}
    ClickEnter
    Sleep    5s

ClickInsertButton
    Click Img   img=${SCREEN_IMAGES_PATH}\\CommonAction\\OpenFile\\insert_button.png
    Sleep    1s

ClickToolsButton
    Click Img   img=${SCREEN_IMAGES_PATH}\\CommonAction\\OpenFile\\tools_button.png
    Sleep    1s

ClickFrontPlane
    Click Img   img=${SCREEN_IMAGES_PATH}\\CommonAction\\OpenFile\\front_plane.png
    Sleep    2s

ClickBrowseButton
    Click Img   img=${SCREEN_IMAGES_PATH}\\CommonAction\\OpenFile\\browse_button.png
    Sleep    1s

ClickComponentButton
    utility.Presskey    k1=down
    Sleep    1s
    utility.Presskey    k1=enter
    Sleep    1s

ClickExistingButton
    utility.Presskey    k1=enter
    Sleep    1s
    

Select All File
    [Arguments]    ${different_action}    ${test_data}
    IF    '${different_action}' == 'AsmInsertToImport'
        Click Img   img=${SCREEN_IMAGES_PATH}\\CommonAction\\OpenFile\\file_type_1.png
    ELSE IF    '${different_action}' == 'PrtInsertFeaturesToImport'
        Click Img   img=${SCREEN_IMAGES_PATH}\\CommonAction\\OpenFile\\file_type_2.png
    ELSE IF    '${different_action}' == 'PrtInsertPartToImport'
        Click Img   img=${SCREEN_IMAGES_PATH}\\CommonAction\\OpenFile\\file_type_3.png
    ELSE IF    '${different_action}' == 'ToolbarsBlocksInsertToImport'
        Click Img   img=${SCREEN_IMAGES_PATH}\\CommonAction\\OpenFile\\file_type_4.png
    ELSE IF    '${different_action}' == 'DrwInsertDXF/DWGToImport'
        Click Img   img=${SCREEN_IMAGES_PATH}\\CommonAction\\OpenFile\\file_type_5.png
    ELSE
        Fail    Unsupported file type: ${different_action}
    END
    Select File Type    ${different_action}    ${test_data}     
    Sleep    1s


GoToPrtPath
    Type Path    path=${PRT}
    ClickEnter
    Sleep    5s

GoToDRWPath
    Type Path    path=${DRW}
    ClickEnter
    Sleep    5s
 