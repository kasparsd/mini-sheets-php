<?php

namespace kasparsd\MiniSheets;

class XlsxBuilder extends AbstractBuilder {

	protected $shared_strings = [];

	protected $zipper;

	public function __construct( \ZipArchive $zipper ) {
		$this->zipper = $zipper;
	}

	public function build() {
		$this->zipper->addEmptyDir( 'docProps' );
		$this->zipper->addFromString( 'docProps/app.xml', $this->app_xml() );
		$this->zipper->addFromString( 'docProps/core.xml', $this->core_xml() );

		$this->zipper->addEmptyDir( '_rels' );
		$this->zipper->addFromString( '_rels/.rels', $this->rels_xml() );

		$this->zipper->addEmptyDir( 'xl/worksheets' );
		$this->zipper->addFromString( 'xl/worksheets/sheet1.xml', $this->sheet_xml() );
		$this->zipper->addFromString( 'xl/workbook.xml', $this->workbook_xml() );
		$this->zipper->addFromString( 'xl/sharedStrings.xml', $this->shared_strings_xml() );

		$this->zipper->addEmptyDir( 'xl/_rels' );
		$this->zipper->addFromString( 'xl/_rels/workbook.xml.rels', self::workbook_rels_xml() );

		$this->zipper->addFromString( '[Content_Types].xml', $this->content_types_xml() );

		return $this->zipper;
	}

	public function xlsx_get_shared_string_no( $string ) {
		static $string_pos = [];

		if ( isset( $this->shared_strings[ $string ] ) ) {
			$this->shared_strings[ $string ] += 1;
		} else {
			$this->shared_strings[ $string ] = 1;
		}

		if ( ! isset( $string_pos[ $string ] ) ) {
			$string_pos[ $string ] = array_search( $string, array_keys( $this->shared_strings ), true );
		}

		return $string_pos[ $string ];
	}

	public function xlsx_cell_name( $row_no, $column_no ) {
		$n = $column_no;

		for ( $r = ''; $n >= 0; $n = intval( $n / 26 ) - 1 ) {
			$r = chr( $n % 26 + 0x41 ) . $r;
		}

		return $r . ( $row_no + 1 );
	}

	public function sheet_xml() {
		$rows = [];

		foreach ( $this->rows as $row_no => $row ) {
			$cells = [];
			$row = array_values( $row );

			foreach ( $row as $col_no => $field_value ) {
				$field_type = 's';

				if ( is_numeric( $field_value ) ) {
					$field_type = 'n';
				}

				$field_value_no = $this->xlsx_get_shared_string_no( $field_value );

				$cells[] = sprintf(
					'<c r="%s" t="%s"><v>%d</v></c>',
					$this->escape_xml( $this->xlsx_cell_name( $row_no, $col_no ) ),
					$this->escape_xml( $field_type ),
					$this->escape_xml( $field_value_no )
				);
			}

			$rows[] = sprintf(
				'<row r="%s">
					%s
				</row>',
				$this->escape_xml( $row_no + 1 ),
				implode( "\n", $cells )
			);
		}

		return sprintf(
			'<?xml version="1.0" encoding="utf-8" standalone="yes"?>
			<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
				<sheetData>
					%s
				</sheetData>
			</worksheet>',
			implode( "\n", $rows )
		);
	}

	public function shared_strings_xml() {
		$shared_strings = [];

		foreach ( $this->shared_strings as $string => $string_count ) {
			$shared_strings[] = sprintf(
				'<si><t>%s</t></si>',
				$this->escape_xml( $string )
			);
		}

		return sprintf(
			'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
			<sst count="%d" uniqueCount="%d" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
				%s
			</sst>',
			array_sum( $this->shared_strings ),
			count( $this->shared_strings ),
			implode( "\n", $shared_strings )
		);
	}


	public function workbook_xml() {
		return sprintf(
			'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
			<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
			<sheets>
				<sheet name="Sheet1" sheetId="1" r:id="rId1" />
			</sheets>
			</workbook>'
		);
	}

	public function content_types_xml() {
		return sprintf(
			'<?xml version="1.0" encoding="UTF-8"?>
			<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
				<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
				<Default Extension="xml" ContentType="application/xml"/>
				<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
				<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
				<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
				<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
				<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
			</Types>'
		);
	}

	public function workbook_rels_xml() {
		return sprintf(
			'<?xml version="1.0" encoding="UTF-8"?>
			<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
				<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
				<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
			</Relationships>'
		);
	}

	public function app_xml() {
		return sprintf(
			'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
			<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
				<Application>MiniSheets</Application>
			</Properties>'
		);
	}

	public function core_xml() {
		return sprintf(
			'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
			<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
				<dcterms:created xsi:type="dcterms:W3CDTF">%s</dcterms:created>
				<dc:creator>MiniSheets</dc:creator>
			</cp:coreProperties>',
			$this->escape_xml( $this->created() )
		);
	}

	public function rels_xml() {
		return sprintf(
			'<?xml version="1.0" encoding="UTF-8"?>
			<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
				<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
				<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
				<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
			</Relationships>'
		);
	}

}
