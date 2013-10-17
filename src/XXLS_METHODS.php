<?php

/**
 * Class of static re-implimentation of Excel methods
 */
class XXLS_METHODS {
	static function X_IF( $bool, $a, $b = 0 ) {
		if( $bool ) {
			return $a;
		} else {
			return $b;
		}
	}

	static function X_MAX() {
		$data = self::array_flatten(func_get_args());

		return max($data);
	}

	static function X_MIN() {
		$data = self::array_flatten(func_get_args());

		return min($data);
	}

	static function X_OR() {
		$data = self::array_flatten(func_get_args());
		foreach( $data as $datum ) {
			if( $datum ) return true;
		}

		return false;
	}

	static function X_AND() {
		for( $i = 0; $i < func_num_args(); $i++ ) {
			if( !func_get_arg($i) ) return false;
		}

		return true;
	}

	static function X_CONCATENATE() {
		$data = self::array_flatten(func_get_args());
		$j    = '';
		foreach( $data as $k ) {
			$j .= $k;
		}

		return $j;
	}

	static function X_MID( $text, $start, $end ) {
		return substr($text, $start - 1, $end);
	}

	static function X_ISEVEN( $x ) {
		return !($x & 1);
	}

	static function X_ISODD( $x ) {
		return !self::X_ISEVEN($x);
	}

	static function X_SUM() {
		$data = self::array_flatten(func_get_args());

		return array_sum($data);
	}

	static function X_NOT( $x ) {
		return !$x;
	}

	static function X_ROUND( $val, $precision = 0 ) {
		return round($val, $precision);
	}

	static function X_ROUNDDOWN( $val, $precision = 0 ) {
		$x = pow(10, $precision);

		return floor($val * $x) / $x;
	}

	static function X_ROUNDUP( $val, $precision = 0 ) {
		return self::X_ROUND($val, $precision);
	}

	static function X_VALUE( $val ) {
		return (double)trim($val);
	}

	static function X_CEILING( $val, $sig = 1 ) {
		return ceil($val / $sig) * $sig;
	}

	static function X_FLOOR( $val, $sig = 1 ) {
		return floor($val / $sig) * $sig;
	}

	static function X_SQRT( $val ) {
		return sqrt($val);
	}

	static function X_VLOOKUP( $lookup_value, array $table_array, $col_index_num, $range_lookup = true ) {
		$leftmost = reset($table_array);

		$index = false;
		if( $range_lookup ) {
			$reverse = array_reverse($leftmost, true);
			foreach( $reverse as $rIndex => $val ) {
				if( strval($val) != '' && $val <= $lookup_value ) {
					$index = $rIndex;
					break;
				}
			}
		} else {
			$index = array_search($lookup_value, $leftmost);
		}

		if( $index === false ) {
			return null;
		}

		return $table_array[$col_index_num - 1][$index];
	}

	/**
	 * Given an array, find all the values recursively.
	 *
	 * @param  array $array             The Array to be Flattened
	 * @return array|NULL                The resulting array or NULL on failure
	 */
	private function array_flatten( $array ) {
		if( !is_array($array) ) return null;
		$it    = new RecursiveIteratorIterator(new RecursiveArrayIterator($array));
		$final = array();
		foreach( $it as $v ) {
			$final[] = $v;
		}

		return $final;
	}

}
