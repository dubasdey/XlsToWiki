<?php
require_once dirname(__FILE__) . '/PHPExcel/IOFactory.php';

function translateCol($inCol){
	$array = array("A"=>0,"B"=>1,"C"=>2,"D"=>3,"E"=>4,"F"=>5,"G"=>6,"H"=>7,"I"=>8,"J"=>9,"K"=>10,"L"=>11,"M"=>12,"N"=>13,"O"=>14,"P"=>15,"Q"=>16,"R"=>17,"S"=>18,"T"=>19,"U"=>20,"V"=>21,"W"=>22,"X"=>23,"Y"=>24,"Z"=>25);
	$pos = 0;
	foreach(str_split($inCol,1) as $part){ $pos+=$array[$part]+1; }
	return $pos;
}

function getMaxCol($table){
	$r = 0;
	foreach($table as $row=>$rowData){
		foreach($rowData as $col=>$colData){ if ($r<$col){ $r = $col; } }
	}
	return $r;
}

$objExcel = null;
if (isset($_FILES['excelFile'])){
	$objExcel = PHPExcel_IOFactory::load($_FILES['excelFile']['tmp_name']); //PHPExcel
}

function renderBorder($border,$prefix){
	if ($border['type'] !='none'){ 
		echo $prefix.':';
		if($border['type'] == 'thin'){ echo '1px solid'; }
		echo " #".$border['color'].";"; 
	}
}
function renderStyle($s){
	$b=$rows[$i]['style']['borders'];
	$f=$rows[$i]['style']['font'];
				
	// BG
	if(isset($s['background'])){echo "background-color:#".$s['background'].";";}
	
	// Font
	if(isset($f['name'])){echo "font-family:".$f['name'].";";}
	if(isset($f['color'])){echo "color:#".$f['color'].";";}
	if($f['bold']){echo "font-weight:bold;";}
	if($f['italic']){echo "font-style: italic;";}
	if($f['underline']!='none'){echo "text-decoration: underline;";}
	if($f['strike']){echo "text-decoration: line-through;";}
	
	//Border
	renderBorder($b['left'],"border-left");
	renderBorder($b['right'],"border-right");
	renderBorder($b['top'],"border-top");
	renderBorder($b['bottom'],"border-bottom");
}

function renderHTML($title,$table){
	$max = getMaxCol($table);
	echo "<h2>$title</h2>";
	echo "<table style=\"border-collapse:collapse;\">";
	foreach($table as $rows){
		echo "<tr>";
		for($i=0;$i<=$max;$i++){
			if (isset($rows[$i])){
				echo '<td style="';
				renderStyle($rows[$i]['style']);
				echo '">'.$rows[$i]['data']."</td>";
			} else {
				echo "<td></td>";
			}
		}
		echo "</tr>";
	}
	echo "</table>";
}

function renderMediaWiki($title,$table){
	$max = getMaxCol($table);
	echo "{|\r\n";
	echo "|+".$title."\r\n";
	foreach($table as $rows){
		echo "|-\r\n";
		for($i=0;$i<=$max;$i++){
			if (isset($rows[$i])){
				echo "|".' style="';
				renderStyle($rows[$i]['style']);
				echo '" '."| ".$rows[$i]['data'];
			} else {
				echo "||";
			}
		}
	}
	echo "|}\r\n";
}
?>
<html>
	<head>
		<title>Excel to Wiki</title>
	</head>
	<body>
<?php 
if($objExcel == null){
?>
	<span>Import xsl xlsx or ods file</span>
	<form action="" method="post" enctype="multipart/form-data">
		<input type="file" name="excelFile">
		<br/>
		<input type="submit" value="Upload" name="submit">
	</form>
<?php 
} else { 
	$sheets = $objExcel->getAllSheets(); // PHPExcel_Worksheet[]
	foreach($sheets as $sheet){
		$title = $sheet->getTitle();
		$cellIds = $sheet->getCellCollection(true); // PHPExcel_Cell[]
		$table = null;
		// Extract data
		foreach($cellIds as $cellId){
			$cell = $sheet->getCell($cellId);
			$row  = $cell->getRow();
			$col  = translateCol($cell->getColumn());
			if($cell->isFormula()){
				$table[$row][$col]['data'] 			= $cell->getOldCalculatedValue();
			}else{
				$table[$row][$col]['data'] 			= $cell->getValue();
			}
			$table[$row][$col]['data_format'] 	= $cell->getFormattedValue();
			$table[$row][$col]['type']  		= $cell->getDataType();
			
			// Font
			$table[$row][$col]['style']['font']['name'] 		= $cell->getStyle()->getFont()->getName();
			$table[$row][$col]['style']['font']['bold'] 		= $cell->getStyle()->getFont()->getBold();
			$table[$row][$col]['style']['font']['italic'] 		= $cell->getStyle()->getFont()->getItalic();
			$table[$row][$col]['style']['font']['strike'] 		= $cell->getStyle()->getFont()->getStrikethrough();
			$table[$row][$col]['style']['font']['underline']	= $cell->getStyle()->getFont()->getUnderline();
			$table[$row][$col]['style']['font']['color'] 		= $cell->getStyle()->getFont()->getColor()->getRGB();
			
			if($cell->getStyle()->getFill()->getFillType() != 'none'){
				$table[$row][$col]['style']['background'] = $cell->getStyle()->getFill()->getStartColor()->getRGB();
			}
			
			$table[$row][$col]['style']['borders']['left']['color'] = $cell->getStyle()->getBorders()->getLeft()->getColor()->getRGB();
			$table[$row][$col]['style']['borders']['left']['type'] = $cell->getStyle()->getBorders()->getLeft()->getBorderStyle();
			$table[$row][$col]['style']['borders']['top']['color'] = $cell->getStyle()->getBorders()->getTop()->getColor()->getRGB();
			$table[$row][$col]['style']['borders']['top']['type'] = $cell->getStyle()->getBorders()->getTop()->getBorderStyle();
			$table[$row][$col]['style']['borders']['right']['color'] = $cell->getStyle()->getBorders()->getRight()->getColor()->getRGB();
			$table[$row][$col]['style']['borders']['right']['type'] = $cell->getStyle()->getBorders()->getRight()->getBorderStyle();
			$table[$row][$col]['style']['borders']['bottom']['color'] = $cell->getStyle()->getBorders()->getBottom()->getColor()->getRGB();
			$table[$row][$col]['style']['borders']['bottom']['type'] = $cell->getStyle()->getBorders()->getBottom()->getBorderStyle();	
		}
		renderHTML($title,$table);
		echo "<pre>";
		renderMediaWiki($title,$table);
		echo "<pre>";
	}
}	
?>
	</body>
</html>