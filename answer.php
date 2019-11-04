
<?php
session_start();
require_once('SimpleXLSX.php');
include_once("connection.php");

    
    $db = new dbObj();
    $connString =  $db->getConnstring();
    //choose name
    $filename="test.xlsx";

    $sql = "SELECT * from essai  ";// to be modified 

    $queryRecords = mysqli_query($connString, $sql) or die("error");
    $array = array();
    while( $row = mysqli_fetch_assoc($queryRecords) ) { 
            
            
// create an array that correspond to your table data 
          $array[] = array(
        'nom' => $row['nom'],
        'description' => $row['description']
        

    );
        }


    
//attribute to each column header a type see the guide of simplexslx
    $header = array(
  ''=>'string',//text
  ''=>'string'//text 
);
    
  $title = array("Your title");
 

//affect table arrays to new table named rows 
$rows = $array;
//create excel styles
$styles1 = array( 'height'=>'30px','font'=>'Arial','font-size'=>15,'font-style'=>'bold', 'fill'=>'#eee', 'halign'=>'center', 'border'=>'left,right,top,bottom');
$styles4 = array( 'width'=>'200px','font'=>'Arial','font-size'=>10,'font-style'=>'bold', 'fill'=>'#81F8D1', 'halign'=>'center', 'border'=>'left,right,top,bottom');

    // Send file headers
header('Content-Type: application/vnd.ms-excel');
header('Content-Disposition: attachment; filename='.basename($filename));

 $writer = new XLSXWriter();
$writer->writeSheetHeader('Sheet1', $header,$col_options = ['widths'=>[20,70,30,20],['suppress_row'=>true]]);// write header
 $writer->writeSheetRow('Sheet1', $title,$styles1);    //write title
//write data
foreach($rows as $row)
    $writer->writeSheetRow('Sheet1', $row,$styles5,$col_options = ['widths'=>[20,70,30,20]]);

$writer->writeToFile($filename);

    readfile($filename);
?>
