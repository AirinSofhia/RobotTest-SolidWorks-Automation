import re
import zipfile
import os
import pydirectinput as pdi
import pyautogui
import shutil
import glob
import time

pdi.FAILSAFE = False

screenWidth, screenHeight = pyautogui.size()
currentMouseX, currentMouseY = pyautogui.position()
confidence, grayscale = 0.8, True

pyautogui.FAILSAFE = False

def presskey(k1):
    pdi.keyDown(k1)
    pdi.keyUp(k1)

def press2keys(k1, k2):
    pdi.keyDown(k1)
    pdi.keyDown(k2)
    pdi.keyUp(k2)
    pdi.keyUp(k1)

def press3keys(k1, k2, k3):
    pdi.keyDown(k1)
    pdi.keyDown(k2)
    pdi.keyDown(k3)
    pdi.keyUp(k3)
    pdi.keyUp(k2)
    pdi.keyUp(k1)

def ActiveWindowsByClick():
    pyautogui.moveTo(screenWidth/4, screenHeight/4)
    pyautogui.click()
    # pdi.click()

def MoveToCenterByClick():
    pyautogui.moveTo(screenWidth/2, screenHeight/2)
    pyautogui.click()

def type_path(path):
    pyautogui.typewrite(path)

def unzip_to_currentFolder(path):
    folder_path, filename = os.path.split(path)
    print(folder_path, filename)
    zip_file = zipfile.ZipFile(path)
    zip_file.extractall(folder_path)
    zip_file.close()

def restore_data(action):
    source_path = "C:\\SolidworksUsageControl\\bak\\" + action + "\\Data"
    destination_path =  "C:\\SolidworksUsageControl\\" + action + "\\Data"
    empty_folder(destination_path)
    copy_demo(source_path, destination_path)

def empty_folder(folder_path):
    for filename in os.listdir(folder_path):
        file_path = os.path.join(folder_path, filename)
        if os.path.isfile(file_path):
            os.remove(file_path)
        elif os.path.isdir(file_path):
            empty_folder(file_path)
            os.rmdir(file_path)

def copy_demo(src_dir, dst_dir):
    if not os.path.exists(dst_dir):
        os.makedirs(dst_dir)
    if os.path.exists(src_dir):
        for file in os.listdir(src_dir):
            file_path = os.path.join(src_dir, file)
            dst_path = os.path.join(dst_dir, file)
            if os.path.isfile(os.path.join(src_dir, file)):
                shutil.copy2(file_path, dst_path)
            else:
                copy_demo(file_path, dst_path)
                # print("存在多级文件夹，正在复制。")

def getfileinfo(files):
    files = list(files.split("\n"))
    fileinfo_dict = {}
    # file_file_list = []
    for file in files:
        # os.path.getsize(file)
        # os.path.getmtime(file)
        fileinfo_dict[file] = [os.path.getsize(file), os.path.getmtime(file)]
    return fileinfo_dict

def rightclick_img(img, Optional = 1):
    # location = verify_img(img)
    # if location != None:
    #     pyautogui.rightClick(location)
    start_time = time.time()
    elapse_time = 0
    while elapse_time < Optional:
        location = verify_img(img)
        time.sleep(0.1)
        elapse_time = time.time()-start_time
        if location != None:
            pyautogui.rightClick(location)
            elapse_time = 6

def click_img(img, Optional = 1):
    start_time = time.time()
    elapse_time = 0
    while elapse_time < Optional:
        location = verify_img(img)
        time.sleep(0.1)
        elapse_time = time.time()-start_time
        if location != None:
            pyautogui.click(location)
            elapse_time = 6
        
def doubleclick_img(img, Optional = 1):
    # location = verify_img(img)
    # if location != None:
    #     pyautogui.doubleClick(location)
    start_time = time.time()
    elapse_time = 0
    while elapse_time < Optional:
        location = verify_img(img)
        time.sleep(0.1)
        elapse_time = time.time()-start_time
        if location != None:
            pyautogui.doubleClick(location)
            elapse_time = 6

def verify_result_for_Solidworks_Save(expected_result, img_deny, verify_files, OldFileInfo):
    try:
        isPass = False
        # expect result is allow, then verify the new protected file's file size is biger and modified time is newer than the old file.
        if expected_result == "allow":
            files = list(verify_files.split("\n"))
            filesize_new = []
            for f in files:
                # file is not exit, return failed
                if not os.path.exists(f):
                    isPass = False
                    msg = "file is not exit, check verify_files in excel"                  
                else:
                # file is exist. check file is protect or not, new file size > old file size and new file modified time > old one.
                    file_object2 = open(f, 'r', errors='ignore')
                    all_the_text = file_object2.read()  # 结果为str类型
                    is_nxlfmt = re.findall("NXLFMT", all_the_text)[0]
                    if is_nxlfmt =="NXLFMT" and os.path.getsize(f) > int(OldFileInfo[f][0]) and os.path.getmtime(f) > float(OldFileInfo[f][1]):
                        isPass = True
                        msg = "expect result is allow, file is edited and update."
                    else:
                        isPass = False
                        msg = "expect result is allow, but file is not protected or not update"
                        # msg = str(OldFileInfo[f][0]) +'/n'+ str(os.path.getsize(f)) +'/n'+ str(OldFileInfo[f][1]) +'/n'+ str(os.path.getmtime(f))
                        break
        # expect result is deny, verify deny message
        elif expected_result == "deny":
            if not os.path.exists(img_deny):
                isPass = False
                msg = "file is not exit, check img_deny in excel"
            else:
                location = verify_img(img_deny)
                print(location)
                if location != None:
                    isPass = True
                    msg = "expect result is deny, and find the deny img"
                else:
                    msg = "expect result is deny, but the deny img is not match"
                    isPass = False
        return isPass, msg
    except Exception as msg:
        print(msg)

def verify_result_for_Solidworks_SaveAs(expected_result, img_deny, verify_files):
    try:
        isPass = False
        # expect result is allow, then verify the new protected file's file size is biger and modified time is newer than the old file.
        if expected_result == "allow":
            files = list(verify_files.split("\n"))
            for f in files:
                # file is not exit, return failed
                if not os.path.exists(f):
                    isPass = False
                    msg = "file is not exit, check verify_saveas_files in excel"                  
                else:
                # file is exist. check file is protect or not
                    file_object2 = open(f, 'r', errors='ignore')
                    all_the_text = file_object2.read()  # 结果为str类型
                    is_nxlfmt = re.findall("NXLFMT", all_the_text)[0]
                    if is_nxlfmt =="NXLFMT":
                        isPass = True
                        msg = "expect result is allow, file is edited and update."
                        print(isPass)
                    else:
                        isPass = False
                        msg = "expect result is allow, but file is not protected"
                        break
        # expect result is deny, verify deny message
        elif expected_result == "deny":
            if not os.path.exists(img_deny):
                isPass = False
                msg = "file is not exit, check img_deny in excel"
            else:
                location = verify_img(img_deny)
                print(location)
                if location != None:
                    isPass = True
                    msg = "expect result is deny, and find the deny img"
                else:
                    msg = "expect result is deny, but the deny img is not match"
                    isPass = False
        return isPass, msg
    except Exception as msg:
        print(msg)

def verify_result_for_Solidworks_export(output_file, tag):
    # List of output files
    files = list(output_file.split("\n"))
    
    # Change tag from type string to list
    tag_list = list(tag.split("\n"))

    pattern = re.compile(r'{"gov.+]}')  # The regular expression of the tag of the output file.
    actual_tags = []
    isPass = False
    
    for f in files:
        # Open the generated file and retrieve the content in text format.ng\e
        export_file = open(f, errors="ignore")
        # text = export_file.read().replace("\x00", "")
        text = export_file.read()
        export_file.close()
        

        # Retrieve the tag from the text.
        match = re.search(pattern, text)
        if match != None:
            actual_tags.append(match.group())

    # print(actual_tags)
    # actual_tags
    isPass = all(x == y for x, y in zip(tag_list, actual_tags))

    return actual_tags, isPass

def verify_img(img):
    try:
        location = pyautogui.locateOnScreen(img, confidence=confidence, grayscale=grayscale)
        return location
    except pyautogui.ImageNotFoundException:
        return None





def select_asm_file_type():
    pyautogui.press('down', presses=13)
    pyautogui.press('enter')

def select_prt_features_file_type():
    pyautogui.press('down', presses=10)
    pyautogui.press('enter')

def select_prt_part_file_type():
    for _ in range(11):
        pyautogui.press('down')
        time.sleep(0.5)  
    pyautogui.press('enter')  
    time.sleep(1)  

def select_dxf_dwg_file_type(test_data):
    if test_data == "dwg":
        for _ in range(1):  
            pyautogui.press('down')
            time.sleep(1)
        pyautogui.press('enter')
    else:
        pyautogui.press('enter')
    time.sleep(1)

def select_toolsbar_file_type(test_data):
    print(f"Received test_data: {test_data}")
    if "DWG" in test_data.upper():  # Check if DWG exists in the string
        print("Selecting DWG...")
        for _ in range(1):  # DWG goes down 1
            pyautogui.press('down')
            time.sleep(2)  # Increased timing
        pyautogui.click()
    elif "DXF" in test_data.upper():  # Check if DXF exists in the string
        print("Selecting DXF...")
        for _ in range(2):  # DXF goes down 2
            pyautogui.press('down')
            time.sleep(2)  # Increased timing
        pyautogui.click()
    else:
        print("Selecting default...")
        pyautogui.click()
    time.sleep(3)
    print("Selection complete.")

def select_file_type(file_type, test_data=None):
    print(f"Received file type: {file_type}, test_data: {test_data}")
    if file_type == "AsmInsertToImport":
        print("Executing select_asm_file_type")
        select_asm_file_type()
    elif file_type == "PrtInsertFeaturesToImport":
        print("Executing select_prt_features_file_type")
        select_prt_features_file_type()
    elif file_type == "PrtInsertPartToImport":
        print("Executing select_prt_part_file_type")
        select_prt_part_file_type()
    elif file_type == "ToolbarsBlocksInsertToImport":
        print("Executing select_toolsbar_file_type")
        select_toolsbar_file_type(test_data)
    elif file_type == "DrwInsertDXF/DWGToImport":
        print("Executing_select_dxf_dwg_file_type")
        select_dxf_dwg_file_type(test_data)
    else:
        raise ValueError(f"Unsupported file type: {file_type}")

def click_features_button():
    for _ in range(3):
        pyautogui.press('down')
        time.sleep(1)  
    pyautogui.press('enter')  
    time.sleep(3) 

def click_imported_button():
    for _ in range(4):
        pyautogui.press('up')
        time.sleep(1)  
    pyautogui.press('enter')  
    time.sleep(1)  

def click_part_button():
    pyautogui.press('up', presses=15)
    pyautogui.press('enter')
    time.sleep(1) 

def click_blocks_button():
    pyautogui.press('up', presses=16)
    pyautogui.press('enter')
    time.sleep(1) 

def select_insert_button():
    for _ in range(2):
        pyautogui.press('down')
        time.sleep(1)  
    pyautogui.press('enter')  
    time.sleep(1)  

def click_picture_button():
    for _ in range(6):
        pyautogui.press('up')
        time.sleep(1)  
    pyautogui.press('enter')  
    time.sleep(1) 

def click_dxf_dwg_button():
    for _ in range(3):
        pyautogui.press('up')
        time.sleep(1)  
    pyautogui.press('enter')  
    time.sleep(1) 

def get_latest_file(files):
    # This function takes a list of files and returns the most recently modified file.
    if not files:
        return None
    # Assume the first file is the latest initially
    latest_file = files[0]
    for file in files:
        file_time = os.path.getmtime(file)
        latest_time = os.path.getmtime(latest_file)
        if file_time > latest_time:
            latest_file = file
    return latest_file

def get_new_file_name(output_dir):
    # Get all valid files in the directory (recursive), excluding directories
    files = [f for f in glob.glob(os.path.join(output_dir, '**'), recursive=True) if os.path.isfile(f)]
    # Exclude temporary files (e.g., starting with '~$')
    valid_files = [f for f in files if not os.path.basename(f).startswith('~$')]
    if not valid_files:
        return None  # Return None if no valid files are found
    # Get the most recently created/modified file
    newest_file = get_latest_file(valid_files)
    return newest_file

def verify_file_size(NEW_FILE, OLD_FILE):
    try:
        isPass = False
        msg = ""
        
        # Check if the files exist
        if not os.path.exists(NEW_FILE):
            isPass = False
            msg = f"New file does not exist: {os.path.basename(NEW_FILE)}"
        elif not os.path.exists(OLD_FILE):
            isPass = False
            msg = f"Old file does not exist: {os.path.basename(OLD_FILE)}"
        else:
            # Get file sizes
            new_file_size = os.path.getsize(NEW_FILE)
            old_file_size = os.path.getsize(OLD_FILE)
            
            # Extract only the file names
            new_file_basename = os.path.basename(NEW_FILE)
            old_file_basename = os.path.basename(OLD_FILE)
            
            # Compare the file size and modification time
            if new_file_size >= old_file_size and os.path.getmtime(NEW_FILE) > os.path.getmtime(OLD_FILE):
                isPass = True
                msg = f"File size are valid. New file: {new_file_basename} (Size: {new_file_size} bytes), Old file: {old_file_basename} (Size: {old_file_size} bytes)"
            else:
                isPass = False
                msg = f"File size is invalid. New file: {new_file_basename} (Size: {new_file_size} bytes), Old file: {old_file_basename} (Size: {old_file_size} bytes)"
        
        return isPass, msg
        
    except Exception as e:
        return False, f"An error occurred: {e}"
    except pyautogui.ImageNotFoundException:
        return None
    
def verify_result(expected_result, img_deny):
    try:
        isPass = False
        if expected_result == "deny":
            if not os.path.exists(img_deny):
                isPass = False
                msg = "file is not exit, check img_deny in excel"
            else:
                location = verify_img(img_deny)
                print(location)
                if location != None:
                    isPass = True
                    msg = "expect result is deny, and find the deny img"
                else:
                    msg = "expect result is deny, but the deny img is not match. Path: " + img_deny 
                    isPass = False
        return isPass, msg
    except Exception as msg:
        print(msg)
