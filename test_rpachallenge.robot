*** Setting ***
Documentation    Test case for rpa challenge
Library     SeleniumLibrary    timeout=10
Library     ExcelLibrary

*** Variables ***
${WEB_URL}                      https://www.rpachallenge.com
${BROWSER}                      chrome
${PAGE_BTN_START}               xpath=//button[text() = 'Start']
${PAGE_BTN_SUBMIT_FORM}         xpath=//input[@type='submit']
${PAGE_BTN_ROUND_ONE}           xpath=//button[text() = 'Round 1']
${PAGE_FIRSTNAME}               xpath=//div[label[text() = 'First Name']]/input
${PAGE_LASTNAME}                xpath=//div[label[text() = 'Last Name']]/input
${PAGE_COMPANYNAME}             xpath=//div[label[text() = 'Company Name']]/input
${PAGE_ROLE_IN_COMPANY}         xpath=//div[label[text() = 'Role in Company']]/input
${PAGE_ADDRESS}                 xpath=//div[label[text() = 'Address']]/input
${PAGE_EMAIL}                   xpath=//div[label[text() = 'Email']]/input
${PAGE_PHONE_NUMBER}            xpath=//div[label[text() = 'Phone Number']]/input

*** Test Cases ***
ต้องการดึงข้อมูลจาก Excel เพื่อกรอกลงในเว็บ RPA Challenge
    [Teardown]    Run Keywords    Tear Down RPA Challenge
    เปิดเว็บ RPA Challenge
    คลิกปุ่ม Start เพื่อเริ่มการกรอกแบบฟอร์ม
    เปิด Excel โดยอ่านไฟล์ challenge.xlsx เพื่อทำการดึงข้อมูลมากรอกในแบบฟอร์ม
    FOR   ${i}   IN RANGE    2    12
        ${data_firstname}=    Read Excel Cell               row_num=${i}        col_num=1
        ${data_lastname}=    Read Excel Cell                row_num=${i}        col_num=2
        ${data_company_name}=    Read Excel Cell            row_num=${i}        col_num=3
        ${data_role_in_company_name}=    Read Excel Cell    row_num=${i}        col_num=4
        ${data_address}=    Read Excel Cell                 row_num=${i}        col_num=5
        ${data_email}=    Read Excel Cell                   row_num=${i}        col_num=6
        ${data_phone_number}=    Read Excel Cell            row_num=${i}        col_num=7
        กรอกข้อมูลชื่อ              ${data_firstname}
        กรอกข้อมูลนามสกุล          ${data_lastname}
        กรอกข้อมูลบริษัท            ${data_company_name}
        กรอกข้อมูลตำแหน่งในบริษัท   ${data_role_in_company_name}
        กรอกข้อมูลที่อยู่             ${data_address}
        กรอกข้อมูลอีเมล์            ${data_email}
        กรอกข้อมูลเบอร์โทร         ${data_phone_number}
        คลิกปุ่ม Submit Form
    END
    ตรวจสอบว่าข้อมูลทั้งหมดถูกส่งไปเรียบร้อย

*** Keywords ***
Tear Down RPA Challenge
    Close All Browsers
    Close All Excel Documents

เปิดเว็บ RPA Challenge
    Open Browser    ${WEB_URL}    browser=${BROWSER}

คลิกปุ่ม Start เพื่อเริ่มการกรอกแบบฟอร์ม
    Wait Until Element Is Visible       ${PAGE_BTN_START}    30
    Click Element                       ${PAGE_BTN_START}
    Wait Until Element Is Visible       ${PAGE_BTN_ROUND_ONE}    30

คลิกปุ่ม Submit Form
    Click Element    ${PAGE_BTN_SUBMIT_FORM}

เปิด Excel โดยอ่านไฟล์ challenge.xlsx เพื่อทำการดึงข้อมูลมากรอกในแบบฟอร์ม
    Open Excel Document    filename=challenge.xlsx     doc_id=doc1

กรอกข้อมูลชื่อ
    [Arguments]    ${data_firstname}
    Input Text     ${PAGE_FIRSTNAME}         ${data_firstname}

กรอกข้อมูลนามสกุล
    [Arguments]    ${data_lastname}
    Input Text     ${PAGE_LASTNAME}          ${data_lastname}

กรอกข้อมูลบริษัท
    [Arguments]    ${data_company_name}
    Input Text     ${PAGE_COMPANYNAME}       ${data_company_name}

กรอกข้อมูลตำแหน่งในบริษัท
    [Arguments]    ${data_role_in_company_name}
    Input Text     ${PAGE_ROLE_IN_COMPANY}       ${data_role_in_company_name}

กรอกข้อมูลที่อยู่
    [Arguments]    ${data_address}
    Input Text     ${PAGE_ADDRESS}       ${data_address}

กรอกข้อมูลอีเมล์
    [Arguments]    ${data_email}
    Input Text     ${PAGE_EMAIL}       ${data_email}

กรอกข้อมูลเบอร์โทร
    [Arguments]    ${data_phone_number}
    Input Text     ${PAGE_PHONE_NUMBER}       ${data_phone_number}

ตรวจสอบว่าข้อมูลทั้งหมดถูกส่งไปเรียบร้อย
    Wait Until Page Contains    Congratulations!    30
    Capture page screenshot