# dynamicDocUpdate
This js script dynamically updates information in a View spreadsheet based on a number of Master spreadheets (all done in new google sheets).</br> 
The View/Master spreadsheet <a href ="https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet-app#openbyidid">id </a> (KEY) is listed in the URL:
</br>  </a> //docs.google.com/spreadsheets/d/abc1234567/edit#gid=0 the "abc1234567" is the id (KEY)
</br>
The function Update Calendar:
</br> 1) runs a query to import selected columns from the Master sheet, and 
</br> 2) displayes the data in a specified format in View sheet

