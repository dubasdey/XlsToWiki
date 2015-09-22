<?

function translateCol($inCol){
	$array = array("A"=>0,"B"=>1,"C"=>2,"D"=>3,"E"=>4,"F"=>5,"G"=>6,"H"=>7,"I"=>8,"J"=>9,"K"=>10,"L"=>11,"M"=>12,"N"=>13,"O"=>14,"P"=>15,"Q"=>16,"R"=>17,"S"=>18,"T"=>19,"U"=>20,"V"=>21,"W"=>22,"X"=>23,"Y"=>24,"Z"=>25);
	$pos = 0;
	foreach(str_split($inCol,1) as $part){ $pos+=$array[$part]+1; }
	return $pos;
}

function getMaxCol($table){
	$r = 0;
	if(count($table)>0){
		foreach($table as $row=>$rowData){
			foreach($rowData as $col=>$colData){ if ($r<$col){ $r = $col; } }
		}
	}
	return $r;
}

function extractData($cell){
	
	if($cell->isFormula()){
		$rowData['data'] 			= $cell->getOldCalculatedValue();
	}else{
		$rowData['data'] 			= $cell->getValue();
	}
	
	$rowData['data_format'] 		= $cell->getFormattedValue();
	$rowData['type']  				= $cell->getDataType();
	
	// Font
	$rowData['style']['font']['name'] 		= $cell->getStyle()->getFont()->getName();
	$rowData['style']['font']['bold'] 		= $cell->getStyle()->getFont()->getBold();
	$rowData['style']['font']['italic'] 	= $cell->getStyle()->getFont()->getItalic();
	$rowData['style']['font']['strike'] 	= $cell->getStyle()->getFont()->getStrikethrough();
	
	$rowData['style']['font']['underline']	= $cell->getStyle()->getFont()->getUnderline() !='none'?true:false;
	
	$rowData['style']['font']['color'] 		= $cell->getStyle()->getFont()->getColor()->getRGB();
	
	if($cell->getStyle()->getFill()->getFillType() != 'none'){
		$rowData['style']['background'] = $cell->getStyle()->getFill()->getStartColor()->getRGB();
	}
	
	$rowData['style']['borders']['left']['color'] 	= $cell->getStyle()->getBorders()->getLeft()->getColor()->getRGB();
	$rowData['style']['borders']['left']['type']	= $cell->getStyle()->getBorders()->getLeft()->getBorderStyle();
	$rowData['style']['borders']['top']['color'] 	= $cell->getStyle()->getBorders()->getTop()->getColor()->getRGB();
	$rowData['style']['borders']['top']['type'] 	= $cell->getStyle()->getBorders()->getTop()->getBorderStyle();
	$rowData['style']['borders']['right']['color'] 	= $cell->getStyle()->getBorders()->getRight()->getColor()->getRGB();
	$rowData['style']['borders']['right']['type'] 	= $cell->getStyle()->getBorders()->getRight()->getBorderStyle();
	$rowData['style']['borders']['bottom']['color'] = $cell->getStyle()->getBorders()->getBottom()->getColor()->getRGB();
	$rowData['style']['borders']['bottom']['type'] 	= $cell->getStyle()->getBorders()->getBottom()->getBorderStyle();		
	return $rowData;
}

function getBorder($border,$prefix){
	$result="";
	if ($border['type'] !='none'){ 
		$result.=$prefix.':';
		if($border['type'] == 'thin'){ $result.= '1px solid'; }
		$result.=' #'.$border['color'].';'; 
	}
	return $result;
}

function getStyle($s,$borders,$font){
	$style="";
	
	// Background
	if(isset($s['background'])){$style.="background-color:#".$s['background'].";";}
	
	// Font
	$f=$s['font'];
	if($font){
		if(isset($f['name'])){$style.= "font-family:".$f['name'].";";}
	}
	
	if(isset($f['color'])){$style.= "color:#".$f['color'].";";}
	if($f['bold']){$style.= "font-weight:bold;";}
	if($f['italic']){$style.= "font-style: italic;";}
	if($f['underline']){$style.= "text-decoration: underline;";}
	if($f['strike']){$style.= "text-decoration: line-through;";}
	
	//Border
	if($borders){
		if($s['borders']['left'] === $s['borders']['right'] && $s['borders']['right'] === $s['borders']['top'] && $s['borders']['top'] === $s['borders']['bottom']){
			$style.=getBorder($s['borders']['left'],"border");
		}else{
			$style.=getBorder($s['borders']['right'],"border-right");
			$style.=getBorder($s['borders']['top'],"border-top");
			$style.=getBorder($s['borders']['bottom'],"border-bottom");	
		}
	}
	return $style;
}

function renderHTML($table,$borders,$font){
	$max = getMaxCol($table);
	if($max>0){
		echo "<table style=\"border-collapse:collapse;\">";
		foreach($table as $rows){
			echo "<tr>";
			for($i=0;$i<=$max;$i++){
				if (isset($rows[$i])){
					echo '<td style="'.getStyle($rows[$i]['style'],$borders,$font).'">'.$rows[$i]['data'].'</td>';
				} else {
					echo "<td></td>";
				}
			}
			echo "</tr>";
		}
		echo "</table>";
	}else{
		echo '<span classs="warn">Empty sheet</span>';
	}
}

function renderMediaWiki($table,$borders,$font){
	$max = getMaxCol($table);
	if($max>0){
		echo "<pre>{|\r\n";
		foreach($table as $rows){
			echo "|-\r\n";
			for($i=0;$i<=$max;$i++){
				if (isset($rows[$i])){
					echo '| style="'.getStyle($rows[$i]['style'],$borders,$font).'"|'.$rows[$i]['data'];
				} else {
					echo '||';
				}
			}
		}
		echo "|}\r\n</pre>";
	}else{
		echo '<span classs="warn">Empty sheet</span>';
	}	
}

?>