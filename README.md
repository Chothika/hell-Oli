*** Setting ***
Library         Selenium2Library
Library         ExcelLibrary
Library         Dialogs
Resource        resources/keywords.robot
Variables       resources/variablesORD.py
# Test Teardown   Close varBrowser

*** Test Cases ***
[TC-001]: New Case
    Login
    Screen Short Proposal       #บันทึกเคสใหม่หน้าย่อ
    Print bPrint                #พิมพ์ใบรับฝากเคส
    Screen Submit Proposal      #ส่งข้อมูลใบคำขอฯ ต่างสาขา
    Screen Full Proposal        #จัดการรายการเคสใหม่
    Screen Verify Proposal      #ตรวจสอบรายละเอียดเกี่ยวกับกรมธรรม์
    Screen Print Receipt        #พิมพ์ใบนำส่งเบี้ยประกัน
    Screen Send AS400           #ถ่ายข้อมูล และ Update เคสใหม่ / ส่งข้อมูลเข้า AS/400

*** Keywords ***
Login
    Open Browser                    ${varURL}    ${varBrowser}
    Maximize Browser Window
    Input Text                      id=username  ${varUsername}
    Input Text                      id=password  ${varPassword}
    Wait Until Element Is Enabled   name=loginButton
    Click Element                   name=loginButton
    Wait Until Page Contains        ออกจากระบบ

Screen Short Proposal
    Open Excel                        ${varExcelFile}
    ${varRows}    Get Row Count       ${varSheetName}
    : For    ${varRow}    In Range    1    ${varRows}
    \   ${varAgent}             Get Value    ${varSheetName}    ${varRow}   AGENT
    \   ${varReqNo}             Get Value    ${varSheetName}    ${varRow}   REQNO
    \   ${varTitle}             Get Value    ${varSheetName}    ${varRow}   TITLE
    \   ${varName}              Get Value    ${varSheetName}    ${varRow}   NAME
    \   ${varSurname}           Get Value    ${varSheetName}    ${varRow}   SURNAME
    \   ${varBirthdate}         Get Value    ${varSheetName}    ${varRow}   BIRTHDATE
    \   ${varCardNumber}        Get Value    ${varSheetName}    ${varRow}   CARDNUMBER
    \   ${varExpireDate}        Get Value    ${varSheetName}    ${varRow}   EXPIREDATE
    \   ${varCheckCard}         Get Value    ${varSheetName}    ${varRow}   CHECKCARD
    \   ${varPhoneHome}         Get Value    ${varSheetName}    ${varRow}   PHONEHOME
    \   ${varPhoneMobile}       Get Value    ${varSheetName}    ${varRow}   PHONEMOBILE
    \   ${varPhoneWOrk}         Get Value    ${varSheetName}    ${varRow}   PHONEWORK
    \   ${varPhoneWorkExt}      Get Value    ${varSheetName}    ${varRow}   PHONEWORKEXT
    \   ${varOccupationList}    Get Value    ${varSheetName}    ${varRow}   OCCUPATIONLIST
    \   ${varHeight}            Get Value    ${varSheetName}    ${varRow}   HEIGHT
    \   ${varWeight}            Get Value    ${varSheetName}    ${varRow}   WEIGHT
    \   ${varPlan}              Get Value    ${varSheetName}    ${varRow}   PLAN
    \   ${varSumAsur}           Get Value    ${varSheetName}    ${varRow}   SUMINSURED
    \   ${varMode}              Get Value    ${varSheetName}    ${varRow}   MODE
    \   ${varRiderCB}           Get Value    ${varSheetName}    ${varRow}   RiderCB
    \   ${varCapCB}             Get Value    ${varSheetName}    ${varRow}   CapitalCB
    \   ${varRiderCPA}          Get Value    ${varSheetName}    ${varRow}   RiderCPA
    \   ${varCapCPA}            Get Value    ${varSheetName}    ${varRow}   CapitalCPA
    #\   ${varRiderCPA6}         Get Value    ${varSheetName}    ${varRow}   RiderCPA26
    #\   ${varCapCPA6}           Get Value    ${varSheetName}    ${varRow}   CapitalCPA26
    \   ${varRiderDAB}          Get Value    ${varSheetName}    ${varRow}   RiderDAB
    \   ${varCapDAB}            Get Value    ${varSheetName}    ${varRow}   CapitalDAB
    #\   ${varRiderDAB2}         Get Value    ${varSheetName}    ${varRow}   RiderDAB2
    #\   ${varCapDAB2}           Get Value    ${varSheetName}    ${varRow}   CapitalDAB2
    \   ${varRiderHC}           Get Value    ${varSheetName}    ${varRow}   RiderHC
    \   ${varCapHC}             Get Value    ${varSheetName}    ${varRow}   CapitalHC
    #\   ${varRiderHCPack}       Get Value    ${varSheetName}    ${varRow}   RiderHCpack
    #\   ${varCapHCPack}         Get Value    ${varSheetName}    ${varRow}   CapitalHCpack
    #\   ${varRiderPB}           Get Value    ${varSheetName}    ${varRow}   RiderPB
    \   ${varRequestDate}       Get Value    ${varSheetName}    ${varRow}   REQUESTDATE
#GOTO NewCase
    \   Go To               ${varURLShortProposal}
    \   Input Text                               id=agent-code-name           ${varAgent}
    \   Sleep                                    1
    \   Press Key                                id=agent-code-name      \\13
    \   Wait Until Element Is Enabled            id=bAddNew
    \   Click Button                             id=bAddNew
    \   Wait Until Element Is Visible            id=newcaseshortly-panel
    \   Input Text                               id=requestId                 ${varReqNo}
    \   Select From List                         id=title                     ${varTitle}
    \   Input Text                               id=name                      ${varName}
    \   Input Text                               id=surname                   ${varSurname}
    \   Click Element                            id=nationalThai
    \   Input Text                               id=birthDate                 ${varBirthdate}
    \   Select From List                         id=cardType                  เลขประจำตัว 13 หลัก
    \   Input Text                               id=cardNumber                ${varCardNumber}
    \   Input Text                               id=expireDate                ${varExpireDate}
    \   Select From List                         id=checkCard                 ${varCheckCard}
    \   Input Text                               id=phoneHome                 ${varPhoneHome}
    \   Input Text                               id=phoneMobile               ${varPhoneMobile}
    \   Input Text                               id=phoneWork                 ${varPhoneWork}
    \   Input Text                               id=phoneWorkExt              ${varPhoneWorkExt}
    \   Click element                            id=occupationList
    \   input text                               id=filteredOccupationList    ${varOccupationList}
    \   Sleep                                    2
    \   Click element                            xpath=//div[3]/div[3]/div/div[3]/div/form[1]/div/table/tbody/tr[2]/td[2]/div[1]/div/ul/li[1]
    \   Click element                            id=rMotorcycleNotuse
    \   Click element                            id=plan
    \   Click element                            id=plan
    \   Wait Until Element Is Visible            id=plan
    \   Select from list                         id=plan                      ${varPlan}
    \   Select from list                         id=sMode                     ${varMode}
    \   Click element                            id=btnPopUpCapital
    \   Click element                            id=popUpCapital
    \   Double click Element                     id=popUpCapital
    \   Press Key                                id=popUpCapital  \\8
    \   Input Text                               id=popUpCapital              ${varSumAsur}
    \   Click element                            xpath=//div[3]/div[10]/div/div[3]/span/button[1]
    \   Wait Until Element Is Visible            id=bAdditionalContract
    \   Click element                            id=bAdditionalContract
    \   Select from list                         id=additionalContract        ${varRiderCB}
    \   Select from list                         id=additionalSumAssure       ${varCapCB}
    \   Click element                            xpath=//div[3]/div[4]/div/div[3]/span/button[1]
    \   Click element                            id=bAdditionalContract
    \   Select from list                         id=additionalContract        ${varRiderCPA}
    \   Select from list                         id=additionalSumAssure       ${varCapCPA}
    \   Click element                            xpath=//div[3]/div[4]/div/div[3]/span/button[1]
    #\   Click element                            id=bAdditionalContract
    #\   Select from list                         id=additionalContract       ${varRiderCPA6}
    #\   Select from list                         id=additionalSumAssure       ${varCapCPA6}
    #\   Click element                            xpath=//div[3]/div[4]/div/div[3]/span/button[1]
    \   Click element                            id=bAdditionalContract
    \   Select from list                         id=additionalContract        ${varRiderDAB}
    \   Select from list                         id=additionalSumAssure       ${varCapDAB}
    \   Click element                            xpath=//div[3]/div[4]/div/div[3]/span/button[1]
    #\   Click element                            id=bAdditionalContract
    #\   Select from list                         id=additionalContract       ${varRiderDAB2}
    #\   Select from list                         id=additionalSumAssure       ${varCapDAB2}
    #\   Click element                            xpath=//div[3]/div[4]/div/div[3]/span/button[1]
    \   Click element                            id=bAdditionalContract
    \   Select from list                         id=additionalContract        ${varRiderHC}
    \   Select from list                         id=additionalSumAssure       ${varCapHC}
    \   Click element                            xpath=//div[3]/div[4]/div/div[3]/span/button[1]
    #\   Click element                            id=bAdditionalContract
    #\   Select from list                         id=additionalContract       ${varRiderHCPack}
    #\   Select from list                         id=additionalSumAssure       ${varCapHCPack}
    #\   Click element                            xpath=//div[3]/div[4]/div/div[3]/span/button[1]
    #\   Click element                            id=bAdditionalContract
    #\   Select from list                         id=additionalContract       ${varRiderPB}
    #\   Click element                            xpath=//div[3]/div[4]/div/div[3]/span/button[1]
    \   Click Element                            id=benefOwnPay1
    \   Click Element                            id=i06Yes
    \   Click Element                            id=i7No
    \   Click Element                            id=i8No
    \   Input Text                               id=height                    ${varHeight}
    \   Input Text                               id=weight                    ${varWeight}
    \   Click Element                            id=i9No
    \   Click Element                            id=i10No
    \   Click Element                            id=i1113No
    \   Click Element                            id=i13No
    \   Click Element                            id=i1416No
    \   Click Element                            id=i17No
    \   Click Element                            id=i18No
    \   Click Element                            id=i19No
    \   Click Element                            id=fatcaNo
    \   Input Text                               id=requestDate               ${varRequestDate}
    \   Wait Until Element Is Enabled            xpath=//div[4]/span/button[2]
    \   Click element                            xpath=//div[4]/span/button[2]
    \   Wait Until Page Contains Element         xpath=//div[7]/div/div[3]/span/button
    \   Click element                            xpath=//div[7]/div/div[3]/span/button
    \   Confirm Action
    \   Click element                            xpath=//div[3]/div/div[4]/span/button[3]
    \   Run Keyword If   "${varRow}"=="${varRows}"   Exit For Loop

Print bPrint
    Click Element                            id=bPrint
    Confirm Action
    Sleep                                    10

Screen Submit Proposal
    Open Excel                        ${varExcelFile}
    ${varRows}    Get Row Count       ${varSheetName}
    : For    ${varRow}    In Range    1    ${varRows}
    \   ${varBranchOwner}               Get Value    ${varSheetName}    ${varRow}   BRANCHOWNER
    \   ${varPdGroupCode}               Get Value    ${varSheetName}    ${varRow}   PDGROUPCODE
    \   ${varAgent}                     Get Value    ${varSheetName}    ${varRow}   AGENT
    \   Go To                           ${varURLSubmitProposal}
    \   Select From List                               id=branchListOwner           ${varBranchOwner}
    \   Select From List                               id=pdGroupCode               ${varPdGroupCode}
    #\   Click Element                                  id=agentName
    #\   Input Text                                     id=agentName                 ${varAgent}
    #\   Sleep                                          1
    #\   Press Key                                      id=agentName  \\13
    \   Click Element                                  id=oSearch
    \   Wait Until Page Contains Element               xpath=//div[2]/div[2]/div[2]/div[2]/div/div/div[2]/div/div/table/tbody[2]/tr[1]/td[1]/Input
    \   Click element                                  xpath=//div[2]/div[2]/div[2]/div[3]/div/div/span/Select
    \   Wait Until Page Contains Element               xpath=//div[2]/div[2]/div[2]/div[3]/div/div/span/Select/option[4]
    \   Click element                                  xpath=//div[2]/div[2]/div[2]/div[3]/div/div/span/Select/option[4]
    \   Wait Until Page Contains Element               xpath=//div[2]/div[2]/div[2]/div[3]/div/div/span/Select/option[4]
    \   Click element                                  xpath=//div[2]/div[2]/div[2]/div[2]/div/div/div[2]/div/table/thead/tr/th[1]/div/input
    \   Click button                                   id=oSendData
    \   Wait Until Page Contains Element               id=datatransmission-before-panel
    \   Click element                                  xpath=//div[2]/div[4]/div/div[3]/span/button[1]
    \   Wait Until Element Is Visible                  id=datatransmission-panel
    \   Click element                                  xpath=//div[2]/div[3]/div/div[3]/span/button
    \   Run Keyword If   "${varRow}"=="${varRow}"   Exit For Loop

Screen Full Proposal
    Open Excel                        ${varExcelFile}
    ${varRows}    Get Row Count       ${varSheetName}
    : For    ${varRow}    In Range    1    ${varRows}
    \   ${varBranchOwner}               Get Value    ${varSheetName}    ${varRow}   BRANCHOWNER
    \   ${varReqNo}                     Get Value    ${varSheetName}    ${varRow}   REQNO
    \   ${varReligion}                  Get Value    ${varSheetName}    ${varRow}   REGION
    \   ${varMarryStatus}               Get Value    ${varSheetName}    ${varRow}   MARRYSTATUS
    \   ${varHomeNo}                    Get Value    ${varSheetName}    ${varRow}   HOMENO
    \   ${varBuilding}                  Get Value    ${varSheetName}    ${varRow}   BUILDING
    \   ${varMoo}                       Get Value    ${varSheetName}    ${varRow}   MOO
    \   ${varSoi}                       Get Value    ${varSheetName}    ${varRow}   SOI
    \   ${varStreet}                    Get Value    ${varSheetName}    ${varRow}   STREET
    \   ${varProvince}                  Get Value    ${varSheetName}    ${varRow}   PROVINCE
    \   ${varDistrict}                  Get Value    ${varSheetName}    ${varRow}   DISTRICT
    \   ${varSubDistrict}               Get Value    ${varSheetName}    ${varRow}   SUBDISTRICT
    #\   ${varCurHomeNo}                 Get Value    ${varSheetName}    ${varRow}   CURHOMENO
    #\   ${varCurBuilding}               Get Value    ${varSheetName}    ${varRow}   CURBUILDING
    #\   ${varCurMoo}                    Get Value    ${varSheetName}    ${varRow}   CURMOO
    #\   ${varCurSoi}                    Get Value    ${varSheetName}    ${varRow}   CURSOI
    #\   ${varCurStreet}                 Get Value    ${varSheetName}    ${varRow}   CURSTREET
    #\   ${varCurProvince}               Get Value    ${varSheetName}    ${varRow}   CURPROVINCE
    #\   ${varCurDistrict}               Get Value    ${varSheetName}    ${varRow}   CURDISTRICT
    #\   ${varCurSubDistrict}            Get Value    ${varSheetName}    ${varRow}   CURSUBDISTRICT
    #\   ${varWorkNo}                    Get Value    ${varSheetName}    ${varRow}   WORKNO
    #\   ${varWorkBuilding}              Get Value    ${varSheetName}    ${varRow}   WORKBUILDING
    #\   ${varWorkMoo}                   Get Value    ${varSheetName}    ${varRow}   WORKMOO
    #\   ${varWorkSoi}                   Get Value    ${varSheetName}    ${varRow}   WORKSOI
    #\   ${varWorkStreet}                Get Value    ${varSheetName}    ${varRow}   WORKSTREET
    #\   ${varWorkProvince}              Get Value    ${varSheetName}    ${varRow}   WORKPROVINCE
    #\   ${varWorkDistrict}              Get Value    ${varSheetName}    ${varRow}   WORKDISTRICT
    #\   ${varWorkSubDistrict}           Get Value    ${varSheetName}    ${varRow}   WORKSUBDISTRICT
    \   ${varCurEmail}                  Get Value    ${varSheetName}    ${varRow}   CUREMAIL
    \   ${varPosition}                  Get Value    ${varSheetName}    ${varRow}   POSITION
    \   ${varWorkType}                  Get Value    ${varSheetName}    ${varRow}   WORKTYPE
    \   ${varBusinessType}              Get Value    ${varSheetName}    ${varRow}   BUSINESSTYPE
    \   ${varSalaryPerYear}             Get Value    ${varSheetName}    ${varRow}   SALARYPERYEAR
    \   ${varReceivePolicy}             Get Value    ${varSheetName}    ${varRow}   RECEIVEPOLICY
    \   ${varTitleBeneficiary}          Get Value    ${varSheetName}    ${varRow}   TITLEBENEFICIARY
    \   ${varNameBeneficiary}           Get Value    ${varSheetName}    ${varRow}   NAMEBENEFICIARY
    \   ${varSurnameBeneficiary}        Get Value    ${varSheetName}    ${varRow}   SURNAMEBENEFICIARY
    \   ${varRelationBeneficiary}       Get Value    ${varSheetName}    ${varRow}   RELATIONBENEFICIARY
    \   ${varAgeBeneficiary}            Get Value    ${varSheetName}    ${varRow}   AGEBENEFICIARY
    \   ${varCardNoBenef}               Get Value    ${varSheetName}    ${varRow}   CARDNOBENEF
    \   ${varReceiptTemp}               Get Value    ${varSheetName}    ${varRow}   RECEIPTTEMP
    \   ${varReceiptDate}               Get Value    ${varSheetName}    ${varRow}   RECEIPTDATE
    \   Go To                           ${varURLFullProposal}
    \   Select From List                         id=branchListOwner           ${varBranchOwner}
    \   Input Text                               id=requestIdCriteria         ${varREQNO}
    \   Click Element                            id=bSearch
    \   Wait Until Page Contains Element         xpath=//div[3]/div[2]/div[2]/div[2]/div/div/div[2]/div/div/table/tbody[2]/tr[1]/td[3]/a
    \   Click link                               xpath=//div[3]/div[2]/div[2]/div[2]/div/div/div[2]/div/div/table/tbody[2]/tr[1]/td[3]/a
    \   Wait Until Element Is Visible            id=newcase-panel
#Tab ผุ้เอาประกัน/ตัวแทน
    \   Select From List                         id=religion                  ${varReligion}
    \   Select From List                         id=marryStatus               ${varMarryStatus}
    \   Input Text                               id=homeNo                    ${varHomeNo}
    \   Input Text                               id=building                  ${varBuilding}
    \   Input Text                               id=moo                       ${varMoo}
    \   Input Text                               id=soi                       ${varSoi}
    \   Input Text                               id=street                    ${varStreet}
    \   Select From List                         id=province                  ${varProvince}
    \   Select From List                         id=district                  ${varDistrict}
    \   Select From List                         id=subdistrict               ${varSubDistrict}
    #\   Input Text                               id=curHomeNo                        ${varCurHomeNo}
    #\   Input Text                               id=curBuilding                      ${varCurBuilding}
    #\   Input Text                               id=curMoo                           ${varCurMoo}
    #\   Input Text                               id=curSoi                           ${varCurSoi}
    #\   Input Text                               id=curStreet                        ${varCurStreet}
    #\   Select From List                         id=curProvince                      ${varCurProvince}
    #\   Select From List                         id=curDistrict                      ${varCurDistrict}
    #\   Select From List                         id=curSubdistrict                   ${varCurSubDistrict}
    #\   Input Text                               id=workNo                           ${varWorkNo}
    #\   Input Text                               id=workBuilding                     ${varWorkBuilding}
    #\   Input Text                               id=workMoo                          ${varWorkMoo}
    #\   Input Text                               id=workSoi                          ${varWorkSoi}
    #\   Input Text                               id=workStreet                       ${varWorkStreet}
    #\   Select From List                         id=workProvince                     ${varWorkProvince}
    #\   Select From List                         id=workDistrict                     ${varWorkDistrict}
    #\   Select From List                         id=workSubdistrict                  ${varWorkSubDistrict}
    \   Click Element                            id=listCusCurrentAddress
    \   Select From List                         id=listCusCurrentAddress           1
    \   Click Element                            id=listCusWorkAddress
    \   Select From List                         id=listCusWorkAddress              3
    \   Input Text                               id=curEmail                        ${varCurEmail}
    \   Click Element                            id=rAddressWork
    \   Select From List by Value                id=listCusCurrentAddress           1
    \   Input Text                               id=position                        ${varPosition}
    \   Input Text                               id=workType                        ${varWorkType}
    \   Input Text                               id=businessType                    ${varBusinessType}
    \   Input Text                               id=salaryPerYear                   ${varSalaryPerYear}
    \   Click element                            xpath=//div[2]/div/form/div[2]/div/ul/li[2]/a
#Tab แบบประกัน
    \   Click Element                            id=payTypeCash
    \   Click element                            xpath=//div[3]/div[3]/div/div[2]/div/form/div[2]/div/div/div[2]/div/table[2]/tbody/tr[6]/td/fieldset/table/tbody/tr[2]/td[2]/div[1]/input
    \   Wait Until Page Contains Element         id=cReturnPost
    \   Click Element                            id=cReturnPost
    \   Select From List                         id=receivePolicyCode                   ${varReceivePolicy}
    \   Select From List by Value                id=branchListReceivePolicyCode         ${varBranchOwner}
    \   ${varTotalPremium}=    Get Text          id=totalPremium
    \   Click element                            xpath=//div[3]/div[3]/div/div[2]/div/form/div[2]/div/ul/li[3]/a
#Tab ผู้รับผลประโยชน์
    \   Select From List                         id=titleBeneficiary           ${varTitleBeneficiary}
    \   Input Text                               id=nameBeneficiary            ${varNameBeneficiary}
    \   Input Text                               id=surnameBeneficiary         ${varSurnameBeneficiary}
    \   Select From List                         id=relationBeneficiary        ${varRelationBeneficiary}
    \   Input Text                               id=ageBeneficiary             ${varAgeBeneficiary}
    \   Select From List                         id=cardTypeBenef              เลขประจำตัว 13 หลัก
    \   Input Text                               id=cardNoBenef                ${varCardNoBenef}
    \   Input Text                               id=perBeneficiary             100
    \   Select From List                         id=listBenefAddress           ที่อยู่เดียวกันกับ ทะเบียนบ้านผู้เอาประกัน
    \   Click Button                             id=bAddBenef
    \   Click element                            xpath=//div[3]/div[3]/div/div[2]/div/form/div[2]/div/ul/li[5]/a
#Tab เอกสารประกอบการขอเอาประกัน
    \   Click Element                            id=cc1
    \   Click Element                            id=cc16
    #\   Click Element                            id=cc3
    #\   Click Element                            id=cc4
    #\   Click Element                            id=cc5
    #\   Click Element                            id=cc6
    #\   Click Element                            id=cc7
    #\   Click Element                            id=cc8
    #\   Click Element                            id=cc11
    \   Click Element                            id=dCheck1Checkbox
    \   Click Element                            id=dCheck2Checkbox
    \   Click Element                            id=dCheck3Checkbox
    \   Click Element                            id=dCheck4Checkbox
    \   Click Element                            id=dCheck5Checkbox
    \   Click element                            xpath=//div[3]/div[3]/div/div[2]/div/form/div[2]/div/ul/li[6]/a
#Tab ใบรับเงินชั่วคราว
    \   Click element                        xpath=//div[3]/div[3]/div/div[2]/div/form/div[2]/div/div/div[6]/div/table/tr[1]/td[2]/input
    \   Input Text                           xpath=//div[3]/div[3]/div/div[2]/div/form/div[2]/div/div/div[6]/div/table/tr[1]/td[2]/input      ${varReceiptTemp}
    \   Wait Until Page Contains Element     xpath=//div[3]/div[3]/div/div[2]/div/form/div[2]/div/div/div[6]/div/table/tr[4]/td[2]/input
    \   Click element                        xpath=//div[3]/div[3]/div/div[2]/div/form/div[2]/div/div/div[6]/div/table/tr[4]/td[2]/input
    \   Input Text                           xpath=//div[3]/div[3]/div/div[2]/div/form/div[2]/div/div/div[6]/div/table/tr[4]/td[2]/input      ${varReceiptDate}
    \   Wait Until Page Contains Element     xpath=//div[3]/div[3]/div/div[2]/div/form/div[2]/div/div/div[6]/div/table/tr[7]/td[2]/input
    \   Click element                        xpath=//div[3]/div[3]/div/div[2]/div/form/div[2]/div/div/div[6]/div/table/tr[7]/td[2]/input
    \   Input Text                           xpath=//div[3]/div[3]/div/div[2]/div/form/div[2]/div/div/div[6]/div/table/tr[7]/td[2]/input      ${varTotalPremium}
    \   Click element                        xpath=//div[3]/div[3]/div/div[3]/span/button[1]
    \   Click element                        xpath=//div[3]/div[9]/div/div[3]/span/button[1]
    \   Confirm Action
    \   Sleep                                    3
    \   click link                               link=พิมพ์ใบสรุป
    \   Sleep                                    9
    \   Click link                               link=พิจารณา
    \   Sleep                                    3
    \   Click element                            xpath=//div[3]/div[12]/div/div[3]/span/button[1]
    #\   ${policyNo}=  Get Text                  id=netSummaryApprovePolicy
    \   Sleep                                    2
    \   click element                            xpath=//div[3]/div[13]/div/div[3]/span/button
    \   Run Keyword If   "${varRow}"=="${varRows}"   Exit For Loop

Screen Verify Proposal
    Open Excel                        ${varExcelFile}
    ${varRows}    Get Row Count       ${varSheetName}
    : For    ${varRow}    In Range    1    ${varRows}
    \   ${varBranchOwner}           Get Value    ${varSheetName}    ${varRow}   BRANCHOWNER
    \   ${varAgent}                 Get Value    ${varSheetName}    ${varRow}   AGENT
    \   Go To               ${varURLVerifyProposal}
    \   Select From List                         id=branchListOwner                     ${varBranchOwner}
    \   Input Text                               id=agentCriteria                       ${varAgent}
    \   Sleep                                    1
    \   Press Key                                id=agentCriteria                   \\13
    \   Click Element                            id=bSearchPrintPolicy
    \   Sleep                                    2
    \   click link                               xpath=//div[10]/div[2]/div[2]/div[2]/div/div/div[2]/div/div/table/tbody[2]/tr[1]/td[1]/a
    \   Click element                            xpath=//div[10]/div[3]/div[2]/div[2]/div/table/tbody/tr[1]/td[1]/button
    \   Click element                            xpath=//div[10]/div[9]/div/div[3]/span/button[1]
    \   Click element                            xpath=//div[10]/div[10]/div/div[2]/div/form/table[3]/tbody/tr[2]/td[2]/Select
    \   Click element                            xpath=//div[10]/div[10]/div/div[2]/div/form/table[3]/tbody/tr[2]/td[2]/Select/option[2]
    \   Click element                            xpath=//div[10]/div[10]/div/div[2]/div/form/table[3]/tbody/tr[3]/td[2]/Select
    \   Click element                            xpath=//div[10]/div[10]/div/div[2]/div/form/table[3]/tbody/tr[3]/td[2]/Select/option[5]
    \   Wait Until Page Contains Element         id=approveBySameRemarkOtherValue
    \   Input Text                               id=approveBySameRemarkOtherValue       เลือกอื่นๆ
    \   Wait Until Page Contains Element         xpath=//div[10]/div[10]/div/div[3]/span/button[3]
    \   Click element                            xpath=//div[10]/div[10]/div/div[3]/span/button[3]
    \   sleep                                    4
    \   Click element                            xpath=//div[10]/div[14]/div/div[3]/span/button
    \   Sleep                                    2
#ตารางกรมธรรม์
    \   Click element                            xpath=//div[10]/div[3]/div[2]/div[2]/div/table/tbody/tr/td[1]/button
    \   Sleep                                    2
#ตาราง CV
    \   Click element                            xpath=//div[10]/div[3]/div[2]/div[2]/div/table/tbody/tr[2]/td[1]/button
    \   Sleep                                    2
#สรุปเอกสารประกอบ
    \   Click element                            xpath=//div[10]/div[3]/div[2]/div[2]/div/table/tbody/tr[3]/td[1]/button
    \   Sleep                                    2
#ค่าชดเชยรายวัน
    \   Click element                            xpath=//div[10]/div[3]/div[2]/div[2]/div/table/tbody/tr[4]/td[1]/button
    \   Sleep                                    2
#หนังสือตอบรับ
    \   Click element                            xpath=//div[10]/div[3]/div[2]/div[2]/div/table/tbody/tr[5]/td[1]/button
    \   Sleep                                    2
    \   Wait Until Element Is Visible            id=newcaseord-complete-print
    \   Click element                            xpath=//div[10]/div[12]/div/div[3]/span/button
    \   Sleep                                    1
    \   Run Keyword If             "${varRow}"=="${varRows}"   Exit For Loop

Screen Print Receipt
    Open Excel                        ${varExcelFile}
    ${varRows}    Get Row Count       ${varSheetName}
    : For    ${varRow}    In Range    1    ${varRows}
    \   ${varBranchOwner}           Get Value    ${varSheetName}    ${varRow}   BRANCHOWNER
    \   ${varAgent}                 Get Value    ${varSheetName}    ${varRow}   AGENT
    \   Go To                       ${varURLPrintReceipt}
    \   Select From List                         id=branchListOwner           ${varBranchOwner}
    \   Input Text                               id=agentOwnerCase            ${varAgent}
    \   Sleep                                    1
    \   Press Key                                id=agentOwnerCase      \\13
    \   Wait Until Element Is Visible            xpath=//div[2]/div[2]/div/div[2]/div/div/div[2]/div/div/table/tbody[2]/tr/td[1]
    \   Click Element                            id=btShowPanelSummary
    \   Wait Until Element Is Visible            id=summary-panel
    \   Click element                            xpath=//div[2]/div[4]/div/div[3]/span/button[1]
    \   Sleep                                    10
    \   Click element                            xpath=//div[2]/div[5]/div/div[3]/span/button
    \   Run Keyword If   "${varRow}"=="${varRow}"   Exit For Loop

Screen Send AS400
    Go To                    ${varURLTransferAS400}
    Click Element                            id=bTransferDataOceanSm
    Click Element                            xpath=//div[2]/div[10]/div/div[3]/span/button[1]
    Sleep                                    4
    Click Element                            xpath=//div[2]/div[12]/div/div[3]/span/button
    Click Element                            id=bTransfer
    Click Element                            xpath=//div[2]/div[9]/div/div[3]/span/button[1]
    Execute manual step    โอ้วแม่เจ้ามันเยี่ยมยอด    default_error=อ้าวยังไม่เสร็จหรือค่ะ

Get Value
    [Arguments]     ${varSheetName}    ${varRow}    ${varColumnName}
    ${varCols} =    Get Column Count    ${varSheetName}
    : For    ${x}    In Range    0    ${varCols}
    \   ${varHeader}    Read Cell Data By Coordinates    ${varSheetName}    ${x}    0
    \   Run Keyword If    "${varHeader}"=="${varColumnName}"    Set Test Variable    ${varCol}    ${x}
    log    ${varCol}

    ${varData}     Read Cell Data By Coordinates    ${varSheetName}    ${varCol}  ${varRow}
    [Return]    ${varData}
