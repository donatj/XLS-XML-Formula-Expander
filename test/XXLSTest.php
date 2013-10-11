<?php

class XXLSTest extends PHPUnit_Framework_TestCase {

	public function testMatch() {

		ini_set('memory_limit', '2048M');

		$xls = new XXLS(__DIR__ . '/ss_test.xml');
//		$xls->debug = true;

		$sheet_data = $xls->getSheetData();


//		$testData = $xls->celltest( 'Inputs', 1, 56 );
//		$this->assertTrue( $testData['passing'],  "\n--\n--\n--\n" . var_export($testData, true) );






		$n = 0;
		foreach( $sheet_data as $sheet => $cols ) {
			foreach( $cols as $col => $rows ) {
				foreach( $rows as $row => $data ) {
					if( $data['formula'] ) {
						$testData = $xls->celltest($sheet, $row, $col);
						$this->assertTrue( $testData['passing'], var_export($testData, true) );
					}
				}
			}
		}



	}


}