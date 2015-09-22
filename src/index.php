<?php
require_once dirname(__FILE__) . '/PHPExcel/IOFactory.php';
require_once dirname(__FILE__) . '/functions.php';

ini_set('memory_limit', '64M');

$objExcel = null;
if (isset($_FILES['excelFile'])){
	$objExcel = PHPExcel_IOFactory::load($_FILES['excelFile']['tmp_name']); //PHPExcel
}

?><!DOCTYPE html> 
<html>
	<head>
		<title>Excel to Wiki</title>
		<meta content="text/html; charset=UTF-8" http-equiv="Content-Type" />
		<link href="index.css" rel="stylesheet" type="text/css" media="screen" />
	</head>
	<body>
<?php 
if($objExcel == null){
?>
	<div id="form">
		<span class="leyend">Import xsl xlsx or ods file</span>
		<form action="" method="post" enctype="multipart/form-data">
			<input class="file" type="file" name="excelFile">
			<span class="type">
				<label class="label" for="type">Convert to</label>
				<select  name="type">
					<option value="html">HTML</option>
					<option value="wiki">Wiki-table</option>
				</select>
			</span>
			<span class="type">
				<input type="checkbox" name="ib" value="true"/>
				<label class="label" for="ib">Ignore borders</label>
			</span>
			<span class="type">
				
				<input type="checkbox" name="if" value="true"/>
				<label class="label" for="if">Ignore font</label>
			</span>			
			<input class="submit" type="submit" value="Convert">
		</form>
	</div>
<?php 
} else { 
	$type = $_POST['type']=='html'?1:2;
	
	$borders = (isset($_POST['ib']) && $_POST['ib']==true)?false:true;
	$font 	 = (isset($_POST['if']) && $_POST['if']==true)?false:true;
	
	$sheets = $objExcel->getAllSheets(); // PHPExcel_Worksheet[]
	$i=0;
	echo '<div class="container">';
	foreach($sheets as $sheet){
		$title   = $sheet->getTitle();
		$cellIds = $sheet->getCellCollection(true); // PHPExcel_Cell[]
		$table   = null;
		
		// Extract data
		foreach($cellIds as $cellId){
			$cell = $sheet->getCell($cellId);
			$row  = $cell->getRow();
			$col  = translateCol($cell->getColumn());
			$table[$row][$col] = extractData($cell);
		}
		
		$objExcel = null; // clear memory
		
		// Print result
		echo "<div id=\"sheet-{$i}\" class=\"sheet\"><h1>$title</h1>";
		switch($type){
			case 1: renderHTML($table,$borders,$font); break;
			case 2: renderMediaWiki($table,$borders,$font); break;
		}
		echo "</div>";
	}
	echo '</div>';
}	
?>
	</body>
</html>