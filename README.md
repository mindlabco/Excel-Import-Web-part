# Excel Import Web part
# Description
It is a custom web part developed in SharePoint Framework (SPFx). We have used Bootstrap, jQuery, third party css, script and custom script. We use this web part to get all the details from the excel sheet and save the values into respective list. We have three list over here i.e. *Students, Teachers, Classes*.

This web part imports the excel sheet (Use Excel format which is given in this repostiory)get the values from Excel File and then based on the selection of the user from web part i.e. Whether the entries should be done in either of the list, it creates the entry in the respective list.

# How to use
To use the web part follow the below steps:-
1) Clone or Download the web part solution
2) Install all the list STPs (which is available inside the repository) in your site (Keep the name same as it is, do not change the name of the list)
3) Enter the values in the excel sheet. Use the excel sheet which is provided in this repository. (Except the dummy data, do not change anything in the excel sheet)
4) Close the excel sheet
5) Navigate to the cloned repository folder
6) Open your system's terminal on that folder
7) In your terminal, Navigate to the web part folder inside the cloned repository
8) Now run *npm install* command to install all the npm packages

# Output

Below Screenshot is the output of this web part

![Image of web part](https://github.com/mindlabco/Excel-Import-Web-part/blob/master/Excel-Import.png)
