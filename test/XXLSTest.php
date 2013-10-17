<?php

class XXLSTest extends PHPUnit_Framework_TestCase {

	public function testMatch() {

//		ini_set('memory_limit', '2048M');

		$xls = new XXLS(__DIR__ . '/ss_test.xml');

		$sheet_data = $xls->getSheetData();

		foreach( $sheet_data as $sheet => $cols ) {
			foreach( $cols as $col => $rows ) {
				foreach( $rows as $row => $data ) {
					$testData = $xls->celltest($sheet, $row, $col);
					$this->assertTrue($testData['passing'], var_export($testData, true));
				}
			}
		}


	}

	public function testBase_xls_rev() {

		$this->assertEquals(1, XXLS::base_xls_rev('A'));
		$this->assertEquals(4, XXLS::base_xls_rev('D'));
		$this->assertEquals(4, XXLS::base_xls_rev('d'));
		$this->assertEquals(26, XXLS::base_xls_rev('Z'));
		$this->assertEquals(27, XXLS::base_xls_rev('AA'));
		$this->assertEquals(205, XXLS::base_xls_rev('GW'));
		$this->assertEquals(4670983, XXLS::base_xls_rev('JESSE'));

	}


}