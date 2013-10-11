<?php

class XXLSTest extends PHPUnit_Framework_TestCase {

	public function testMatch() {

//		ini_set('memory_limit', '2048M');

		$xls = new XXLS(__DIR__ . '/ss_test.xml');

		$sheet_data  = $xls->getSheetData();

		foreach( $sheet_data as $sheet => $cols ) {
			foreach( $cols as $col => $rows ) {
				foreach( $rows as $row => $data ) {
					if( isset($data['formula']) && $data['formula'] ) {
						$testData = $xls->celltest($sheet, $row, $col);
						$this->assertTrue($testData['passing'], var_export($testData, true));
					}
				}
			}
		}


	}


}