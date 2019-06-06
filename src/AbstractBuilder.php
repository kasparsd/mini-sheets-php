<?php

namespace kasparsd\MiniSheets;

abstract class AbstractBuilder {

	protected $rows = [];

	public function add_row( $fields ) {
		$this->rows[] = $fields;
	}

	public function add_rows( $rows ) {
		$this->rows = array_merge( $this->rows, $rows );
	}

	public function rows() {
		return $this->rows;
	}

	/**
	 * Return the created timestamp in W3CDTF format.
	 *
	 * @return string
	 */
	public function created() {
		return gmdate( 'Y-m-d\TH:i:s\Z' );
	}

	public function escape_xml( $string ) {
		return htmlspecialchars( $string, ENT_XML1 | ENT_COMPAT, 'UTF-8' );
	}

	abstract public function build();

}
