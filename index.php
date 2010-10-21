<?
$references = array();

$dom = DOMDocument::load( 'tests.xml' );
$Worksheets = $dom->getElementsByTagName( 'Worksheet' );
foreach($Worksheets as $Worksheet) {
	$rows = $Worksheet->getElementsByTagName( 'Row' );
	$row_index = 1;
	$sheetname = $Worksheet->getAttribute( 'ss:Name' );
	foreach ($rows as $row)	{
		$cells = $row->getElementsByTagName( 'Cell' );
		$index = 1;
		foreach( $cells as $cell ) {
			$ind = $cell->getAttribute( 'ss:Index' );
			if ( $ind != null ) {
				//echo "Picked up a special index! Woo";
				$index = $ind;
			}

			$formula = trim($cell->getAttribute( 'ss:Formula' ), ' =');

			$spreadsheet_data[$sheetname][ $row_index ][ $index ] = array( 
			'value' => $cell->nodeValue,
			'formula' => $formula,
			);

			$index += 1;
		}

		$row_index++;
	}
}

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
		}
	}
}

function expand_eq( $formula, $row_index, $col_index, $sheet ) {
	//$expanded_formula = $formula;
	global $spreadsheet_data;
	preg_match_all('/(?<!:)(?:(?P<sheet>[A-Z]{1,})!|\'(?P<sheet2>[A-Z ]{1,})\'!|)R\[?(?P<row>-?\d{0,})\]?C\[?(?P<cell>-?\d{0,})\]?(?!:)/six', $formula, $matches);
	$expanded_formula = preg_replace('/((?<!:)(?:(?P<sheet>[A-Z]{1,})!|\'(?P<sheet2>[A-Z ]{1,})\'!|)R\[?(?P<row>-?\d{0,})\]?C\[?(?P<cell>-?\d{0,})\]?(?!:))/six', '((\1))', $formula);
	//print_r( $matches );
	foreach( $matches[0] as $index => &$match ) {
		if( strlen( $matches['sheet'][$index] ) > 0 ) {
			$cur_sheet = $matches['sheet'][$index];
		}elseif( strlen( $matches['sheet2'][$index] ) > 0 ) {
			$cur_sheet = $matches['sheet2'][$index];
		}else{
			$cur_sheet = $sheet;	
		}

		$cur_row = $row_index + $matches['row'][$index];
		$cur_col = $col_index + $matches['cell'][$index];

		$cur_selected =& $spreadsheet_data[ $cur_sheet ][ $cur_row ][ $cur_col ];

		if( strlen($cur_selected['expanded']) ) {

		}elseif( strlen($cur_selected['formula']) ){
			$cur_selected['expanded'] = expand_eq( $cur_selected['formula'], $cur_row, $cur_col, $cur_sheet );
		}else{
			$cur_selected['expanded'] = " #{$cur_sheet}_{$cur_row}_{$cur_col}# ";
		}

		$expanded_formula = str_replace( "(({$match}))", ' ( ' . $cur_selected[ 'expanded' ] . ' ) ', $expanded_formula );

	}

	return $expanded_formula;

}

echo '<pre>';
print_r( $spreadsheet_data );