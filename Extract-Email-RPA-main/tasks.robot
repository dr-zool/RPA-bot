*** Settings ***
Documentation     Read and extract from email with specific keyword and list those details in an excel file
Library           RPA.Email.ImapSmtp
...               smtp_server=smtp.gmail.com
...               smtp_port=587
Library           RPA.Excel.Files

*** Variables ***
${USERNAME}       gmail
${PASSWORD}       password
${EXCEL_FILE}     ./output/output.xlsx
${KEYWORD}        SUBJECT "Coursera"

*** Tasks ***
List emails
    Authorize    account=${USERNAME}    password=${PASSWORD}
    @{emails}    List Messages    ${KEYWORD}
    FOR    ${email}    IN    @{EMAILS}
        Log    ${email}[Subject]
        Log    ${email}[From]
        Log    ${email}[Date]
        Log    ${email}[Delivered-To]
        Log    ${email}[Received]
        Log    ${email}[Has-Attachments]
        Log    ${email}[uid]
    END

Write to Excel
    @{emails}    List Messages    ${KEYWORD}
    Create Workbook    ${EXCEL_FILE}
    FOR    ${email}    IN    @{EMAILS}
        &{row}=    Create Dictionary
        ...    Email ID    ${email}[uid]
        ...    From    ${email}[From]
        ...    Date    ${email}[Date]
        ...    Has Attachments?    ${email}[Has-Attachments]
        # ...    Received    ${email}[Received]
        Append Rows to Worksheet    ${row}    header=${TRUE}
    END
    Save Workbook

*** Keywords ***
