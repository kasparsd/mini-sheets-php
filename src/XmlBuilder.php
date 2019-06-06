<?php

namespace kasparsd\MiniSheets;

class XmlBuilder extends AbstractBuilder {

	protected $props = [];

	public function add_props( $props ) {
		$this->props = array_merge( $this->props, $props );
	}

	public function props() {
		return array_merge(
			$this->props,
			[
				'Created' => $this->created(),
			]
		);
	}

	public function build() {
		$props_fields = [];

		foreach ( $this->props() as $prop_key => $prop_value ) {
			$props_fields[] = sprintf(
				'<%1$s>%2$s</%1$s>',
				$this->escape_xml( $prop_key ),
				$this->escape_xml( $prop_value )
			);
		}

		return sprintf(
			'<?xml version="1.0"?>
			<?mso-application progid="Excel.Sheet"?>
			<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
				xmlns:o="urn:schemas-microsoft-com:office:office"
				xmlns:x="urn:schemas-microsoft-com:office:excel"
				xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
				xmlns:html="http://www.w3.org/TR/REC-html40">
			<DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">
			 	%s
			</DocumentProperties>
			<Worksheet ss:Name="Sheet1">
				<Table>%s</Table>
			</Worksheet>
			</Workbook>',
			implode( "\n", $props_fields ),
			implode( "\n", $this->rows_formatted() )
		);
	}

	protected function rows_formatted() {
		$rows = [];

		foreach ( $this->rows() as $row ) {
			$cells = [];

			foreach ( $row as $field_value ) {
				$field_type = 'String';

				if ( is_numeric( $field_value ) ) {
					$field_type = 'Number';
				}

				$cells[] = sprintf(
					'<Cell><Data ss:Type="%s">%s</Data></Cell>',
					$this->escape_xml( $field_type ),
					$this->escape_xml( $field_value )
				);
			}

			$rows[] = sprintf(
				'<Row>%s</Row>',
				implode( "\n", $cells )
			);
		}

		return $rows;
	}

}
