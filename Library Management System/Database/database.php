<?php

// Creating the connection to the excel file on the path
$xlFile = realpath('https://drivinc-my.sharepoint.com/personal/roxana_boiciuc_driv_com/Documents/Roxana%20Boiciuc/Personal/Library%20Management%20System/Database/google_books_1299.xls');
$xlDirectory = dirname($xlFile);
$connectionString = odbc_connect("Driver={Microsoft Excel Driver (*.xls)};DriverId=790;Dbq=$xlFile;DefaultDir=$xlDir" , '', '');

// creating the select query to indicate what I want to select from the excel sheet
$sqlQuery = "SELECT author, title FROM google_books_1299";

//code to perform the query on the database
$results = odbc_exec($connectionString, $sqlQuery);

//creating variabes for the fields I want to display on the website
while(odbc_fetch_row($results))
{ $output1 = odbc_result($results, 1);
  $output2 = odbc_result($results, 2);
}

//display the results on the website with print
print(“<html><head>Excel Data</head><body>”);
print("$output1 ");
print("$output2 ");
print(“</body></html>”);

//closing the connection
odbc_close($connectionString);

?>
