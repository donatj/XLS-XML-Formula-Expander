<?php
/** 
* 
* XXLS
* 
* A PHP based Excel 2003 XML Parser / Formula Evaluator
* 
* @author Jesse G. Donat <donatj@gmail.com>
* @license http://opensource.org/licenses/mit-license.php
* @version .8
* 
*/

class XXLS {

	private $sheet_data = array();
	public $debug = false;

	/**
	* @param string $filename
	* @return XXLS
	*/
	function __construct( $filename ) {
		$this->sheet_data = $this->ss_parse( $filename );
	}

	/**
	* Process an Excel 2003 XML File into an Array
	* 
	* @param string $filename
	*/
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
						$xd['value'] = $datas->item(0)->nodeValue;
					}else{
						$xd['value'] = $cell->nodeValue;
					}
					
					$xd['formula'] = $formula ? $formula : null;

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
	
	/**
	* Evaluate an Excel Formula
	* 
	* @param string $sheet
	* @param string|int $col
	* @param int $row
	* @return mixed
	*/
	public function evaluate( $sheet, $col, $row ) {
		if( !is_numeric($col) ) { $col = self::base_xls_rev( $col ); }
		
		$expanded = $this->expand_eq( $this->sheet_data[$sheet][$row][$col]['formula'], $row, $col, $sheet );
		extract( $GLOBALS['xbob'] );

		return eval( 'return ' . $expanded . ';' );		
	}
	/**
	* Test a cells value either automatically or by expected value
	* 
	* @param string $sheet
	* @param string|int $col
	* @param int $row
	* @param mixed $expected
	*/
	public function auto_test( $sheet, $col, $row, $expected = null ) {
		if( !is_numeric($col) ) { $col = self::base_xls_rev( $col ); }
		
		if( $expected === null ) {
			$expected = $this->sheet_data[$sheet][$row][$col]['value'];
		}
		$result = $this->evaluate( $sheet, $col, $row );


		echo '<div ';
		if( $result == $expected ) {
			echo 'style="background: #a0b96a">';
		}else{
			echo 'style="background: #b96a6a">';
			$err = true;
		}

		echo '<pre style="display:inline">';
		echo $sheet . '!' . self::base_xls( $col ) . $row .  ':	<small>EXP:<em>'.$expected.'</em>	CALC:<em>' . $result . '</em></small>';
		echo '</pre>';
		if( $err /*|| !strlen( $result )*/ ) {
			echo '<div style="border: 1px solid #aaa; max-height: 200px; overflow: auto; background: #eee"><pre>' . $expanded . '</pre></div>';
		}
		echo '</div>' . PHP_EOL;
		flush();
	}

	/**
	* Fully expands an Excel formula
	* 
	* @access private
	* 
	* @param string $formula
	* @param int $row_index
	* @param int $col_index
	* @param string $sheet
	* @param int $depth used for recursion
	* @return string
	*/
	private function expand_eq( $formula, $row_index, $col_index, $sheet, $depth = 0 ) {
		
		if( strlen( $this->sheet_data[ $sheet ][ $row_index ][ $col_index ]['expanded'] ) ) {
			return $this->sheet_data[ $sheet ][ $row_index ][ $col_index ]['expanded'];
		}
		
		$expanded_formula = $formula;

		$expanded_formula = self::ms_string( $expanded_formula );

		$RANGE = '/(((?:(?P<sheet>[A-Z]{1,})!|\'(?P<sheet2>[A-Z ()]+)\'!)?R((\[(?P<rowrel>-?\d+)\])|(?P<rowabs>\d+))?C((\[(?P<colrel>-?\d+)\])|(?P<colabs>\d+))?):(R((\[(?P<rowrel2>-?\d+)\])|(?P<rowabs2>\d+))?C((\[(?P<colrel2>-?\d+)\])|(?P<colabs2>\d+))?))/si';

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

					$cur_selected =& $this->sheet_data[ $cur_sheet ][ $range_row ][ $range_col ];

					if( strlen($cur_selected['expanded']) ) {
						//
					}elseif( strlen($cur_selected['formula']) ){
						$cur_selected['expanded'] = $this->expand_eq( $cur_selected['formula'], $range_row, $range_col, $cur_sheet, $depth + 1 );
					}else{
						$cur_selected['expanded'] = " \$".self::sheet_clean($cur_sheet)."_{$range_row}_{$range_col} ";
						$GLOBALS[self::sheet_clean($cur_sheet)."_{$range_row}_{$range_col}"] = $cur_selected['value'];
						$GLOBALS['xbob'][self::sheet_clean($cur_sheet)."_{$range_row}_{$range_col}"] = $cur_selected['value'];
					}

					$finals[] = $cur_selected['expanded'];

				}
			}

			$xls_cellname = self::sheet_clean($cur_sheet). "!" . self::base_xls( $cur_col ) . $cur_row . ':' . self::base_xls( $cur_col2 ) . $cur_row2;

			$expanded_formula = str_replace( "///{$match}///", PHP_EOL . str_repeat( "\t", $depth) . ($this->debug ? ' /* RANGE '.$xls_cellname.' « */ ' : '') . implode(' , ', $finals ) . ($this->debug ? ' /* » RANGE */ ' : '') . PHP_EOL, $expanded_formula );

		}

		// --------------------------------------------------------------------

		//LITTERAL REPLACMENT / EXPANSION
		$LITTERAL = '/(?<!:)((?:(?P<sheet>[A-Z]{1,})!|\'(?P<sheet2>[A-Z ()]+)\'!)?R((\[(?P<rowrel>-?\d+)\])|(?P<rowabs>\d+))?C((\[(?P<colrel>-?\d+)\])|(?P<colabs>\d+))?)(?!:)/si';

		preg_match_all($LITTERAL, $expanded_formula, $matches);
		$expanded_formula = preg_replace($LITTERAL, '///\1///', $expanded_formula);

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

			$cur_selected =& $this->sheet_data[ $cur_sheet ][ $cur_row ][ $cur_col ];

			$temp = false;

			if( strlen($cur_selected['expanded']) ) {
				if( strlen($cur_selected['formula']) == 0 ) {
					$temp = true;
				}
			}elseif( strlen($cur_selected['formula']) ){
				$cur_selected['expanded'] = $this->expand_eq( $cur_selected['formula'], $cur_row, $cur_col, $cur_sheet, $depth + 1 );
			}else{
				$cur_selected['expanded'] = " \$".self::sheet_clean($cur_sheet)."_{$cur_row}_{$cur_col} ";
				$GLOBALS[self::sheet_clean($cur_sheet)."_{$cur_row}_{$cur_col}"] = $cur_selected['value'];
				$GLOBALS['xbob'][self::sheet_clean($cur_sheet)."_{$cur_row}_{$cur_col}"] = $cur_selected['value'];
				$temp = true;
			}

			$xls_cellname = self::sheet_clean($cur_sheet). "!" . self::base_xls( $cur_col ) . $cur_row;
			$posname = $xls_cellname . ' ' . $depth . ( $temp ? ' value: ' . $cur_selected['value'] : '') . ';';

			$expanded_formula = str_replace( "///{$match}///", PHP_EOL . str_repeat( "\t", $depth) . ($this->debug ? ' ( /* '. $posname .' « */ ' : '') . $cur_selected[ 'expanded' ] . ($this->debug ? ' /* » '. $xls_cellname .' */ ) ' : '') . PHP_EOL, $expanded_formula );

		}

		//Special PI handling
		$expanded_formula = preg_replace('/PI\(\)/i', pi(), $expanded_formula);

		//Functions
		$expanded_formula = preg_replace('/([A-Z]{1,})\(/six', ' XXLS_METHODS::X_\1 ( ', $expanded_formula);
		$expanded_formula = preg_replace('/(?<![=])=(?![=])/six', '==', $expanded_formula);

		//Power Expansion
		$expanded_formula .= ' '; //lazy fix for overflow issue.

		$x = 0;
		while( $x = strpos($expanded_formula, '^', $x + 1 ) ) {
			$base = self::get_local_exp_part( $expanded_formula, $x, false, $data_b );
			$exp  = self::get_local_exp_part( $expanded_formula, $x, true, $data_e );
			$expanded_formula = substr( $expanded_formula, 0, $data_b['end'] ) . ' pow ( ' . $base . ' , ' . $exp . ' ) ' . substr( $expanded_formula, $data_e['end'] + 1 );
		}
		
		$this->sheet_data[ $sheet ][ $row_index ][ $col_index ]['expanded'] = $expanded_formula;

		file_put_contents( 'cache/' . md5( json_encode( array(  $row_index, $col_index, $sheet ) ) ) . '.php', '<?' . PHP_EOL . $expanded_formula );

		return $expanded_formula;

	}

	/**
	* Converts Base10 to BaseExcelColumn
	* 
	* @param int $number
	* @return string
	*/
	static function base_xls( $number ) {
		$str = base_convert($number - 1, 10, 26);
		$str = strtr( $str, '0123456789abcdefghijklmnopq', 'ABCDEFGHIJKLMNOPQRSTUVWXYZ');
		for( $i = 0; $i <= strlen($str) - 2; $i++ ) {
			$str[ $i ] = chr(ord( $str[$i]) - 1 );
		}
		return $str;
	}

	/**
	* Converts BaseExcelColumn to Base10
	* 
	* @param string $letter
	* @return int
	*/
	static function base_xls_rev( $letter ) {
		return strpos('ABCDEFGHIJKLMNOPQRSTUVWXYZ', strtoupper( $letter ) ) + 1;
	}

	/**
	* Removes non-alpha-numeric characters from sheetnames
	* 
	* @param string $str
	* @return string
	*/
	static private function sheet_clean( $str ) {
		return preg_replace('/[^A-Z]/six', 'X', $str);	
	}

	/**
	* Finds parts of a carrot (^) style exponent, half at a time.
	* 
	* @param string $equat The equation to search for
	* @param int $init_pos Initial positon to begin searching for ^
	* @param bool $exp If true exponent, false base
	* @param mixed $data by reference information about the found return
	* @return string
	*/
	static private function get_local_exp_part( $equat, $init_pos, $exp = false, &$data = null ) {
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

	/**
	* Converts Microsoft style string, eg "10""" to C style "10\""
	* 
	* @param string $formula
	* @return string
	*/
	static private function ms_string( $formula ) {
		if( strpos($formula,'""') ) {
			for( $i = 0; $i <= strlen( $formula ); $i++ ) {
				if( $str_init && $formula[$i] == '"' ) {
					if( $formula[$i + 1] != '"' ) {
						$str_init = false;
					}else{
						$str .= '\\"';
						$i += 2;
					}
				}else{
					if( !$str_init && $formula[$i] == '"' ) {
						$str_init = true;
					}
				}
				$str .= $formula[$i];
			}		
			return $str;	
		}
		return $formula;
	}

}

/**
* Class of static reimplimentation of Excel methods
*/
class XXLS_METHODS {
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

	static function X_ROUND( $val, $precision = 0 ) {
		return round( $val, $precision );
	}

	static function X_ROUNDDOWN( $val, $precision = 0 ) {
		$x = pow( 10, $precision );
		return floor( $val * $x ) / $x;
	}

	static function X_VALUE( $val ) {
		return (double)trim($val);
	}

}
