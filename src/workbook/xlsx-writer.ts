import { strToU8, zipSync } from 'fflate';

export interface PaymentSourceSummary {
	paymentSource: string;
	totalCents: number;
}

export interface CardSummary {
	card: string;
	totalCents: number;
	sources: PaymentSourceSummary[];
}

export function createExpenseSummaryWorkbook(cards: CardSummary[]): ArrayBuffer {
	const worksheet = createWorksheet(cards);
	const archive = zipSync({
		'[Content_Types].xml': strToU8(contentTypesXml()),
		'_rels/.rels': strToU8(rootRelationshipsXml()),
		'xl/workbook.xml': strToU8(workbookXml()),
		'xl/_rels/workbook.xml.rels': strToU8(workbookRelationshipsXml()),
		'xl/styles.xml': strToU8(stylesXml()),
		'xl/worksheets/sheet1.xml': strToU8(worksheet),
	});
	return archive.buffer;
}

function createWorksheet(cards: CardSummary[]): string {
	const rows = [
		row(1, 22, [textCell('A1', 'Expense summary', 1)]),
		row(2, 20, [textCell('A2', 'Card / payment source', 2), textCell('B2', 'Amount', 2)]),
	];

	let rowNumber = 3;
	for (const card of cards) {
		rows.push(
			row(rowNumber, 20, [
				textCell(`A${rowNumber}`, `#${card.card}`, 3),
				numberCell(`B${rowNumber}`, card.totalCents, 4),
			]),
		);
		rowNumber += 1;

		for (const source of card.sources) {
			rows.push(
				row(rowNumber, 15, [
					textCell(`A${rowNumber}`, `#${source.paymentSource}`, 5),
					numberCell(`B${rowNumber}`, source.totalCents, 6),
				]),
			);
			rowNumber += 1;
		}
	}

	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
	<sheetViews><sheetView workbookViewId="0" showGridLines="0"><pane ySplit="2" topLeftCell="A3" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
	<sheetFormatPr defaultRowHeight="15"/>
	<cols><col min="1" max="1" width="30" customWidth="1"/><col min="2" max="2" width="15" customWidth="1"/></cols>
	<sheetData>${rows.join('')}</sheetData>
	<mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>
</worksheet>`;
}

function row(number: number, height: number, cells: string[]): string {
	return `<row r="${number}" ht="${height}" customHeight="1">${cells.join('')}</row>`;
}

function textCell(reference: string, value: string, style: number): string {
	return `<c r="${reference}" s="${style}" t="inlineStr"><is><t>${escapeXml(value)}</t></is></c>`;
}

function numberCell(reference: string, cents: number, style: number): string {
	return `<c r="${reference}" s="${style}"><v>${formatCents(cents)}</v></c>`;
}

function formatCents(cents: number): string {
	const absolute = Math.abs(cents);
	const dollars = Math.trunc(absolute / 100);
	const remainingCents = String(absolute % 100).padStart(2, '0');
	return `${cents < 0 ? '-' : ''}${dollars}.${remainingCents}`;
}

function escapeXml(value: string): string {
	return value
		.replaceAll('&', '&amp;')
		.replaceAll('<', '&lt;')
		.replaceAll('>', '&gt;')
		.replaceAll('"', '&quot;')
		.replaceAll("'", '&apos;');
}

function contentTypesXml(): string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
	<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
	<Default Extension="xml" ContentType="application/xml"/>
	<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
	<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
	<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`;
}

function rootRelationshipsXml(): string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
	<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`;
}

function workbookXml(): string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
	<sheets><sheet name="Summary" sheetId="1" r:id="rId1"/></sheets>
</workbook>`;
}

function workbookRelationshipsXml(): string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
	<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
	<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`;
}

function stylesXml(): string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
	<numFmts count="1"><numFmt numFmtId="164" formatCode="$#,##0.00;[Red]-$#,##0.00"/></numFmts>
	<fonts count="3">
		<font><sz val="11"/><color rgb="FF000000"/><name val="Calibri"/><family val="2"/></font>
		<font><b/><sz val="14"/><color rgb="FFFFFFFF"/><name val="Aptos"/><family val="2"/></font>
		<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>
	</fonts>
	<fills count="7">
		<fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>
		<fill><patternFill patternType="solid"><fgColor rgb="FF202020"/><bgColor indexed="64"/></patternFill></fill>
		<fill><patternFill patternType="solid"><fgColor rgb="FF303030"/><bgColor indexed="64"/></patternFill></fill>
		<fill><patternFill patternType="solid"><fgColor rgb="FF888888"/><bgColor indexed="64"/></patternFill></fill>
		<fill><patternFill patternType="solid"><fgColor rgb="FF193A64"/><bgColor indexed="64"/></patternFill></fill>
		<fill><patternFill patternType="solid"><fgColor rgb="FFFFFFFF"/><bgColor indexed="64"/></patternFill></fill>
	</fills>
	<borders count="2">
		<border><left/><right/><top/><bottom/><diagonal/></border>
		<border><left style="thin"><color rgb="FF000000"/></left><right style="thin"><color rgb="FF000000"/></right><top style="thin"><color rgb="FF000000"/></top><bottom style="thin"><color rgb="FF000000"/></bottom><diagonal/></border>
	</borders>
	<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
	<cellXfs count="7">
		<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
		<xf numFmtId="0" fontId="1" fillId="2" borderId="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
		<xf numFmtId="0" fontId="2" fillId="3" borderId="1" applyFont="1" applyFill="1" applyBorder="1"/>
		<xf numFmtId="0" fontId="2" fillId="4" borderId="1" applyFont="1" applyFill="1" applyBorder="1"/>
		<xf numFmtId="164" fontId="2" fillId="5" borderId="1" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right"/></xf>
		<xf numFmtId="0" fontId="0" fillId="6" borderId="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment vertical="center"/></xf>
		<xf numFmtId="164" fontId="0" fillId="6" borderId="1" applyNumberFormat="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="right" vertical="center"/></xf>
	</cellXfs>
	<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>`;
}
