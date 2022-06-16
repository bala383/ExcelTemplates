# ExcelTemplates
Generic excel file upload and download using OpenXml

ExcelTemplateReader is a service.

Method Name: 'ExportListToExcel'  
 -generic list<T> any type of list . 
 -sheetname we can speicify what ever we can. 
 - columns is option parameter. if you want to include specific columns only we can do by passing properity name as ',' seperater.
  
Returns the binary format of excel file .

Method Name:  'ReadExcelFile'  
 -IformFile as first parameter  . 
   
Returns the list<T> 
