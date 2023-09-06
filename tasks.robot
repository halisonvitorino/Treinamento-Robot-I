*** Settings ***
Documentation       Insert the sales data for the week and export it as a PDF.

Library             RPA.Browser.Selenium    auto_close=${FALSE}
Library             RPA.HTTP
Library             RPA.Excel.Files
Library             RPA.PDF

*** Variables ***
${DOWNLOAD_PATH}=   ${OUTPUT DIR}${/}resources/SalesData.xlsx

*** Tasks ***
Insert the sales data for the week and export it as a PDF
    Open the intranet website
    Log in
    Download the Excel file
    Fill the form using the data from the Excel file
    Collect the results
    Export the table as PDF
    [Teardown]    Log out and close the browser

*** Keywords ***
Open the intranet website
    Open Chrome Browser    https://robotsparebinindustries.com/
    Maximize Browser Window

Log in
    Input Text        username    maria
    Input Password    password    thoushallnotpass
    Submit Form
    Wait Until Page Contains Element    id:sales-form

Download the Excel file
    Download    https://robotsparebinindustries.com/SalesData.xlsx    overwrite=True    verify=false    target_file=${DOWNLOAD_PATH}

Fill the form using the data from the Excel file
    Open Workbook     ${OUTPUT DIR}${/}resources/SalesData.xlsx
    ${sales_reps}=    Read Worksheet As Table    header=True
    Close Workbook
    FOR    ${sales_rep}    IN    @{sales_reps}
        Wait Until Page Contains Element           firstname
        Fill and submit the form for one person    ${sales_rep}
    END

Fill and submit the form for one person
    [Arguments]                                ${sales_rep}
    Input Text                   firstname     ${sales_rep}[First Name]
    Input Text                   lastname      ${sales_rep}[Last Name]
    Select From List By Value    salestarget   ${sales_rep}[Sales Target]
    Input Text                   salesresult   ${sales_rep}[Sales]
    Submit Form

Collect the results
    Screenshot    css:div.sales-summary    ${OUTPUT_DIR}${/}output/sales_summary.png    

Export the table as PDF
    Wait Until Element Is Visible    id:sales-results
    ${sales_results_html} =    Get Element Attribute    id:sales-results    outerHTML    
    Html To Pdf    ${sales_results_html}    ${OUTPUT_DIR}${/}output/sales_summary.pdf

Log out and close the browser
    Click Button    Log out
    Close Browser

