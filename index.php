<?
header('Content-Type: text/html; charset=utf-8');
error_reporting(E_ALL ^ E_NOTICE);
echo '<pre>';

function ss_parse( $filename ) {
	$dom = DOMDocument::load( $filename );
	$Worksheets = $dom->getElementsByTagName( 'Worksheet' );
	foreach($Worksheets as $Worksheet) {
		$rows = $Worksheet->getElementsByTagName( 'Row' );
		$row_index = 1;
		$sheetname = $Worksheet->getAttribute( 'ss:Name' );
		foreach ($rows as $row)	{
			
			$rind = $row->getAttribute( 'ss:Index' );
			if ( $rind != null ) {
				$row_index = $rind;
			}
			
			$cells = $row->getElementsByTagName( 'Cell' );
			$index = 1;
			foreach( $cells as $cell ) {
				
				$cind = $cell->getAttribute( 'ss:Index' );
				if ( $cind != null ) {
					$index = $cind;
				}

				$formula = trim($cell->getAttribute( 'ss:Formula' ), ' =');
				
				if( $datas = $cell->getElementsByTagName('Data') ) {
					$xd =  array( 
						'value' => $datas->item(0)->nodeValue,
						'formula' => $formula,
					);
				}else{
					$xd =  array( 
						'value' => $cell->nodeValue,
						'formula' => $formula,
					);	
				}
				
				if( strlen( $xd['value'] ) || strlen( $xd['formula'] ) ) {
					$spreadsheet_data[$sheetname][ $row_index ][ $index ] = $xd;
				}
				
				$index += 1;
			}

			$row_index++;
		}
	}
	return $spreadsheet_data;
}

$spreadsheet_data = ss_parse( 'test2.xml' );

foreach( $spreadsheet_data as $sheetname => &$sheet ) {

	foreach( $sheet as $row_index => &$column ) {

		foreach( $column as $col_index => &$col_value ) {
			if( !strlen( $col_value['expanded'] ) ){
				if( strlen($col_value['formula']) ) {
					$col_value['expanded']=  expand_eq( $col_value['formula'], $row_index, $col_index, $sheetname );
				}else{
					//$col_value['expanded'] = " #{$sheetname}_{$row_index}_{$col_index}# ";
				}
			}

			if( $col_value['expanded'] ) {
				echo $sheetname . ':'  .base_xls($col_index) . '' . $row_index . ' <strong>Forumla :' . '</strong>:<div style="max-height: 200px; overflow: auto; background: #eee">' . $col_value['expanded'] . /*':' . $col_value['formula'] . */ '</div><br />';
				echo '<strong>Returns</strong>: ' . eval( 'return ' . $col_value['expanded'] . ';' ) . '<br /><br />';
				
				
				flush();
			}

		}
	}

function auto_test( $sheet, $col, $row ) {
	global $spreadsheet_data;

	if( !is_numeric($col) ) { $col = base_xls_rev( $col ); }
	
	$expanded = expand_eq( $spreadsheet_data[$sheet][$row][$col]['formula'], $row, $col, $sheet );
	extract( $GLOBALS['xbob'] );
	$expected = $spreadsheet_data[$sheet][$row][$col]['value'];
	
	echo '<div ';
	if( ($result = eval( 'return ' . $expanded . ';' ) ) == $expected ) {
		echo 'style="background: #a0b96a">';
	}else{
		echo 'style="background: #b96a6a">';
	}
	echo '<pre style="display:inline">';
	echo $sheet . '!' . base_xls( $col ) . $row .  ':	<small>EXP:<em>'.$expected.'</em>	CALC:<em>' . $result . '</em></small>';
	echo '</pre></div>';
}

class XML_XLS {
	static function X_IF( $bool, $a, $b = 0 ) {
		if( $bool ) {
			return $a;
		}else{
			return $b;	
		}
	}

	static function X_MAX() {
		return max( func_get_args() );
	}

	static function X_MIN() {
		return min( func_get_args() );
	}

	static function X_OR() {
		for ($i = 0;$i < func_num_args();$i++) {
			if( func_get_arg($i) ) return true;
		}
		return false;
	}

	static function X_AND( $a, $b ) {
		return $a && $b;
	}

	static function X_CONCATENATE() {
		$j = '';
		for ($i = 0;$i < func_num_args();$i++) {
			$j .= func_get_arg($i);
		}
		return $j;
	}

	static function X_MID( $text, $start, $end ) {
		return substr( $text, $start - 1, $end );
	}

	static function X_ISEVEN( $x ) {
		return !( $x & 1 );
	}
	
	static function X_ISODD( $x ) {
		return !self::X_ISEVEN( $x );
	}
	
	static function X_SUM() {
		$j = 0;
		for ($i = 0;$i < func_num_args();$i++) {
			$j += func_get_arg($i);
		}
		return $j;
	}
	
	static function X_NOT( $x ) {
		return !$x;
	}
	
}



function expand_eq( $formula, $row_index, $col_index, $sheet, $depth = 0 ) {
	global $spreadsheet_data;
	
	$expanded_formula = $formula;
	
	$RANGE = '/(((?:(?P<sheet>[A-Z]{1,})!|\Z(?P<sheet2>[A-Z ()]+)\Z!)?R((\[(?P<rowrel>-?\d+)\])|(?P<rowabs>\d+))?C((\[(?P<colrel>-?\d+)\])|(?P<colabs>\d+))?):(R((\[(?P<rowrel2>-?\d+)\])|(?P<rowabs2>\d+))?C((\[(?P<colrel2>-?\d+)\])|(?P<colabs2>\d+))?))/si';
	
	preg_match_all($RANGE, $expanded_formula, $matches);
	$expanded_formula = preg_replace($RANGE, '///\1///', $expanded_formula);

	foreach( $matches[0] as $index => &$match ) {
		
		if( strlen( $matches['sheet'][$index] ) > 0 ) {
			$cur_sheet = $matches['sheet'][$index];
		}elseif( strlen( $matches['sheet2'][$index] ) > 0 ) {
			$cur_sheet = $matches['sheet2'][$index];
		}else{
			$cur_sheet = $sheet;	
		}
		
		if( $matches['rowrel'][$index] ) {
			$cur_row = (int)$row_index + (int)$matches['rowrel'][$index];		
		}elseif( $matches['rowabs'][$index] ) {
			$cur_row = (int)$matches['rowabs'][$index];
		}else{
			$cur_row = (int)$row_index;
		}
		
		if( $matches['colrel'][$index] ) {
			$cur_col = (int)$col_index + (int)$matches['colrel'][$index];		
		}elseif( $matches['colabs'][$index] ) {
			$cur_col =  (int)$matches['colabs'][$index];
		}else{
			$cur_col = (int)$col_index;
		}
		
		if( $matches['rowrel2'][$index] ) {
			$cur_row2 = (int)$row_index + (int)$matches['rowrel2'][$index];		
		}elseif( $matches['rowabs2'][$index] ) {
			$cur_row2 = (int)$matches['rowabs2'][$index];
		}else{
			$cur_row2 = (int)$row_index;
		}
		
		if( $matches['colrel2'][$index] ) {
			$cur_col2 = (int)$col_index + (int)$matches['colrel2'][$index];
		}elseif( $matches['colabs2'][$index] ) {
			$cur_col2 =  (int)$matches['colabs2'][$index];
		}else{
			$cur_col2 = (int)$col_index;
		}
		
		$finals = array();
		
		for( $range_col = $cur_col; $range_col <= $cur_col2; $range_col++ ) {
			for( $range_row = $cur_row; $range_row <= $cur_row2; $range_row++ ) {
				
				$cur_selected =& $spreadsheet_data[ $cur_sheet ][ $range_row ][ $range_col ];
				
				if( strlen($cur_selected['expanded']) ) {
					if( strlen($cur_selected['formula']) == 0 ) {
						$temp = true;
					}
				}elseif( strlen($cur_selected['formula']) ){
					$cur_selected['expanded'] = expand_eq( $cur_selected['formula'], $range_row, $range_col, $cur_sheet, $depth + 1 );
				}else{
					$cur_selected['expanded'] = " \$".sheet_clean($cur_sheet)."_{$range_row}_{$range_col} ";
					$GLOBALS[sheet_clean($cur_sheet)."_{$range_row}_{$range_col}"] = $cur_selected['value'];
					$GLOBALS['xbob'][sheet_clean($cur_sheet)."_{$range_row}_{$range_col}"] = $cur_selected['value'];
					$temp = true;
				}
				
				$finals[] = $cur_selected['expanded'];
				
			}
		}
				
		$expanded_formula = str_replace( "///{$match}///", PHP_EOL . str_repeat( "\t", $depth) . ' ( /* RANGE « */ ' . $cur_selected[ 'expanded' ] . ' /* » RANGE */ ) ' . PHP_EOL, $expanded_formula );
		
	}

	// --------------------------------------------------------------------
	
	//LITTERAL REPLACMENT / EXPANSION
	$LITTERAL = '/(?<!:)((?:(?P<sheet>[A-Z]{1,})!|\Z(?P<sheet2>[A-Z ()]+)\Z!)?R((\[(?P<rowrel>-?\d+)\])|(?P<rowabs>\d+))?C((\[(?P<colrel>-?\d+)\])|(?P<colabs>\d+))?)(?!:)/si';
	
	preg_match_all($LITTERAL, $expanded_formula, $matches);
	$expanded_formula = preg_replace($LITTERAL, '((\1))', $expanded_formula);

	foreach( $matches[0] as $index => &$match ) {
	
		if( strlen( $matches['sheet'][$index] ) > 0 ) {
			$cur_sheet = $matches['sheet'][$index];
		}elseif( strlen( $matches['sheet2'][$index] ) > 0 ) {
			$cur_sheet = $matches['sheet2'][$index];
		}else{
			$cur_sheet = $sheet;	
		}
		
		if( $matches['rowrel'][$index] ) {
			$cur_row = (int)$row_index + (int)$matches['rowrel'][$index];		
		}elseif( $matches['rowabs'][$index] ) {
			$cur_row = (int)$matches['rowabs'][$index];
		}else{
			$cur_row = (int)$row_index;
		}
		
		if( $matches['colrel'][$index] ) {
			$cur_col = (int)$col_index + (int)$matches['colrel'][$index];		
		}elseif( $matches['colabs'][$index] ) {
			$cur_col =  (int)$matches['colabs'][$index];
		}else{
			$cur_col = (int)$col_index;
		}
		
		$cur_selected =& $spreadsheet_data[ $cur_sheet ][ $cur_row ][ $cur_col ];
		
		$temp = false;
		
		if( strlen($cur_selected['expanded']) ) {
			if( strlen($cur_selected['formula']) == 0 ) {
				$temp = true;
			}
		}elseif( strlen($cur_selected['formula']) ){
			$cur_selected['expanded'] = expand_eq( $cur_selected['formula'], $cur_row, $cur_col, $cur_sheet, $depth + 1 );
		}else{
			$cur_selected['expanded'] = " \$".sheet_clean($cur_sheet)."_{$cur_row}_{$cur_col} ";
			$GLOBALS[sheet_clean($cur_sheet)."_{$cur_row}_{$cur_col}"] = $cur_selected['value'];
			$GLOBALS['xbob'][sheet_clean($cur_sheet)."_{$cur_row}_{$cur_col}"] = $cur_selected['value'];
			$temp = true;
		}
		
		$xls_cellname = sheet_clean($cur_sheet). "!" . base_xls( $cur_col ) . $cur_row;
		$posname = $xls_cellname . ' ' . $depth . ( $temp ? ' value: ' . $cur_selected['value'] : '') . ';';

		$expanded_formula = str_replace( "(({$match}))", PHP_EOL . str_repeat( "\t", $depth) . ' ( /* '. $posname .' « */ ' . $cur_selected[ 'expanded' ] . ' /* » '. $xls_cellname .' */ ) ' . PHP_EOL, $expanded_formula );

	}
	
	//Special PI handling
	$expanded_formula = preg_replace('/PI\(\)/i', pi(), $expanded_formula);
	
	//Functions
	$expanded_formula = preg_replace('/([A-Z]{1,})\(/six', ' XML_XLS::X_\1 ( ', $expanded_formula);
	$expanded_formula = preg_replace('/(?<![=])=(?![=])/six', '==', $expanded_formula);
	
	//Power Expansion
	$expanded_formula .= ' '; //lazy fix for overflow issue.
	
	$x = 0;
	while( $x = strpos($expanded_formula, '^', $x + 1 ) ) {
		$base = get_local_exp_part( $expanded_formula, $x, false, $data_b );
		$exp  = get_local_exp_part( $expanded_formula, $x, true, $data_e );
		$expanded_formula = substr( $expanded_formula, 0, $data_b['end'] ) . ' pow ( ' . $base . ' , ' . $exp . ' ) ' . substr( $expanded_formula, $data_e['end'] + 1 );
	}

	return $expanded_formula;

}

function base_xls( $number ) {
	$str = base_convert($number - 1, 10, 26);
	$str = strtr( $str, '0123456789abcdefghijklmnopq', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ');
	for( $i = 0; $i <= strlen($str) - 2; $i++ ) {
		$str[ $i ] = chr(ord( $str[$i]) - 1 );
	}
	return $str;
}

function base_xls_rev( $letter ) {
	return strpos('ABCDEFGHIJKLMNOPQRSTUVWXYZ', strtoupper( $letter ) ) + 1;
}

function sheet_clean( $str ) {
	return preg_replace('/[^A-Z]/six', 'X', $str);	
}

function get_local_exp_part( $equat, $init_pos, $exp = false, &$data = null ) {
	static $index = 0;
	$index++;
	$part = '';
	$data = false;
	$open_paren = 0;
	
	for( $i = 1; $i <= 10000; $i++) {

		$j = ( $exp ? 0 - $i : $i );

		if( !$data && $equat[$init_pos - $j] != ' ' ) {
			$data = array( 'pos' => $init_pos - $j, 'char' => $equat[$init_pos - $j], 'index' => $index );
		}

		if( $data ) {
			if( $exp ) {
				$part .= $equat[$init_pos - $j];
			}else{
				$part = $equat[$init_pos - $j] . $part;
			}
		}

		if( $data ) {
			if( $data['char'] == ($exp ? '(' : ')') ) {

				if( $equat[$init_pos - $j] == ($exp ? '(' : ')') ) {
					$open_paren++;
				}elseif( $equat[$init_pos - $j] == (!$exp ? '(' : ')') ) {
					$open_paren--;
				}

				if( $open_paren == 0 ) {
					$data['end'] = $init_pos - $j;
					break;
				}

			}else{				
				if( preg_match('/[^a-zA-Z0-9_\-$\.]/i', $equat[$init_pos - $j] ) || $equat[$init_pos - $j] == '' ) {
					if( $exp ) {
						if( $equat[$init_pos - $j] != '' ) {
							$part = substr( $part, 0, -1 );
						}
						$data['end'] = $init_pos - $j - 1;
					}else{
						if( $equat[$init_pos - $j] != '' ) {
							$part = substr( $part, 1 );
						}
						$data['end'] = $init_pos - $j + 1;
					}
					break;
					
				}
			}
		}
	}
	
	return $part;

}