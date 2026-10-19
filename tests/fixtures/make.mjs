// Сборка синтетических тестовых файлов. Данные выдуманы.
// Запуск: node tests/fixtures/make.mjs  (нужен soffice для файлов из html и csv)
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { execFileSync } from 'node:child_process';
import { buildZip } from '../../src/zip.mjs';

const dir = path.dirname(new URL(import.meta.url).pathname);
const NS = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14"`;
const XD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const r = (t, rpr = '') => `<w:r>${rpr ? `<w:rPr>${rpr}</w:rPr>` : ''}<w:t xml:space="preserve">${t}</w:t></w:r>`;
const p = (...runs) => `<w:p>${runs.join('')}</w:p>`;

const textbox = `<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="2" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="2400000" cy="600000"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="1" name="Надпись 1"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2400000" cy="600000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr><wps:txbx><w:txbxContent><w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="1000"/></w:tabs></w:pPr>${r('Ответственный: ')}${r('Абрамов Дми', '<w:b/>')}${r('трий')}</w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:pict><v:shape id="tb1" style="width:189pt;height:47pt" type="#_x0000_t202"><v:textbox><w:txbxContent><w:p>${r('Ответственный: Абрамов Дмитрий')}</w:p></w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r>`;

const documentXml = XD + `<w:document ${NS}><w:body>
${p(r('ДОГОВОР № 17 с '), r('ООО «Ро', '<w:b/>'), r('маш', '<w:i/>'), r('ка» ', '<w:u w:val="single"/><w:color w:val="C00000"/>'), r('на поставку'))}
${p(r('Представитель: '), r('Петров П'), r('ётр&#160;Ильич'), r(', соисполнитель: '), '<w:r><w:t>Кали­нин</w:t></w:r>', r(' и Кали'), '<w:r><w:softHyphen/></w:r>', r('нин Олег'))}
${p(r('Партнёр: Бета &amp; Ко, '), r('ИНН'), '<w:r><w:tab/></w:r>', r('7729654321'), r(', код '), r('Сидор'), '<w:r><w:tab/></w:r>', r('Сидоров'))}
${p(r('Подписант: '), '<w:del w:id="1" w:author="Иванов И.И." w:date="2026-09-01T10:00:00Z"><w:r><w:delText>Сорокин Па</w:delText></w:r><w:r><w:rPr><w:b/></w:rPr><w:delText>вел</w:delText></w:r></w:del>', '<w:ins w:id="2" w:author="Иванов И.И." w:date="2026-09-01T10:00:00Z"><w:r><w:t>Орлов </w:t></w:r><w:r><w:t>Олег</w:t></w:r></w:ins>', r(', генеральный директор'))}
${p(r('Почта: '), '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> HYPERLINK "mailto:m.kuznetsova@sev-terminal.example" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>', r('m.kuznetsova@', '<w:rStyle w:val="Hyperlink"/>'), r('sev-terminal.ru', '<w:rStyle w:val="Hyperlink"/>'), '<w:r><w:fldChar w:fldCharType="end"/></w:r>', r(' и '), '<w:hyperlink r:id="rIdLink">', r('сайт'), '</w:hyperlink>')}
${p(r('Сноска к паспорту'), '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r>', r(' и концевая'), '<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteReference w:id="1"/></w:r>')}
<w:p><w:commentRangeStart w:id="0"/>${r('Срок поставки 10 дней')}<w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>
<w:p>${r('Надпись: ')}${textbox}</w:p>
<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:sz="4"/><w:bottom w:val="single" w:sz="4"/><w:insideH w:val="single" w:sz="4"/><w:insideV w:val="single" w:sz="4"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="5000"/></w:tblGrid>
<w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${p(r('Контакт'))}</w:tc><w:tc><w:tcPr><w:tcW w:w="5000" w:type="dxa"/></w:tcPr>${p(r('Кузнецова Мари', '<w:b/>'), r('на Олеговна'))}</w:tc></w:tr>
<w:tr><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr>${p(r('Счёт'))}</w:tc><w:tc><w:tcPr><w:tcW w:w="5000" w:type="dxa"/></w:tcPr>${p(r('40702810938000012345'))}</w:tc></w:tr>
</w:tbl>
${p(r('Конец документа. Метка-ловушка [ОРГ_1] уже была в тексте.'))}
<w:sectPr><w:headerReference w:type="default" r:id="rIdH1"/><w:footerReference w:type="default" r:id="rIdF1"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1134" w:right="850" w:bottom="1134" w:left="1701" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>
</w:body></w:document>`;

const hdr = (tag, body) => XD + `<w:${tag} ${NS}>${body}</w:${tag}>`;
const files = {
  '[Content_Types].xml': XD + `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/><Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/><Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/><Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`,
  '_rels/.rels': XD + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>`,
  'word/_rels/document.xml.rels': XD + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdS" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rIdSet" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" Target="settings.xml"/><Relationship Id="rIdH1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="rIdF1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/><Relationship Id="rIdFn" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/><Relationship Id="rIdEn" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/><Relationship Id="rIdC" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/><Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://sev-terminal.ru/kuznetsova" TargetMode="External"/></Relationships>`,
  'word/document.xml': documentXml,
  'word/styles.xml': XD + `<w:styles ${NS}><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Liberation Serif" w:hAnsi="Liberation Serif" w:cs="Liberation Serif"/><w:sz w:val="24"/><w:lang w:val="ru-RU"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style></w:styles>`,
  'word/settings.xml': XD + `<w:settings ${NS}><w:footnotePr><w:footnote w:id="-1"/><w:footnote w:id="0"/></w:footnotePr><w:endnotePr><w:endnote w:id="-1"/><w:endnote w:id="0"/></w:endnotePr></w:settings>`,
  'word/header1.xml': hdr('hdr', p(r('ООО «Ромашка»', '<w:b/>'), r(' · конфиденциально'))),
  'word/footer1.xml': hdr('ftr', p(r('Тел. +7 (916) 245-'), r('18-73'))),
  'word/footnotes.xml': hdr('footnotes', `<w:footnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:footnote><w:footnote w:id="1">${p('<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r>', r(' Паспорт 45 18 № 603217'))}</w:footnote>`),
  'word/endnotes.xml': hdr('endnotes', `<w:endnote w:type="separator" w:id="-1"><w:p><w:r><w:separator/></w:r></w:p></w:endnote><w:endnote w:type="continuationSeparator" w:id="0"><w:p><w:r><w:continuationSeparator/></w:r></w:p></w:endnote><w:endnote w:id="1">${p('<w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:endnoteRef/></w:r>', r(' Связь: '), r('m.kuznetsova@sev-terminal.example'))}</w:endnote>`),
  'word/comments.xml': hdr('comments', `<w:comment w:id="0" w:author="Иванов И.И." w:date="2026-09-01T10:00:00Z" w:initials="ИИ">${p(r('Уточнить у ООО «Ромашка»'))}</w:comment>`),
  'docProps/core.xml': XD + `<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>Договор с ООО «Ромашка»</dc:title><dc:creator>Иванов Иван Иванович</dc:creator><cp:lastModifiedBy>Кузнецова Марина Олеговна</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">2026-09-01T10:00:00Z</dcterms:created></cp:coreProperties>`,
  'docProps/app.xml': XD + `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Microsoft Office Word</Application><Company>ООО «Ромашка»</Company><Manager>Сорокин Павел</Manager></Properties>`,
};

// xlsx: два листа, общие строки с форматированием, inlineStr, формулы, ИНН числом, примечание
const SS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const shared = ['Клиент', 'ООО «Ромашка»', 'ИНН', 'Сумма', 'Счёт', null, ' Иванов Иван ', 'Итого'];
const sst = shared.map(s => s == null
  ? '<si><r><rPr><b/><sz val="11"/><rFont val="Calibri"/></rPr><t>Кузне</t></r><r><rPr><i/><sz val="11"/><rFont val="Calibri"/></rPr><t xml:space="preserve">цова Марина </t></r><r><t>Олеговна</t></r></si>'
  : `<si><t${/^\s|\s$/.test(s) ? ' xml:space="preserve"' : ''}>${s.replace(/&/g, '&amp;')}</t></si>`).join('');
const sheet1 = XD + `<worksheet ${SS}><dimension ref="A1:F5"/><sheetData>
<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>2</v></c><c r="C1" t="s"><v>3</v></c><c r="D1" t="s"><v>4</v></c></row>
<row r="2"><c r="A2" t="s"><v>1</v></c><c r="B2" s="1"><v>7729654321</v></c><c r="C2"><v>1500</v></c><c r="D2"><v>4.0702810938000001E+19</v></c><c r="E2" t="str"><f>IF(A2="ООО «Ромашка»","да","нет")</f><v>да</v></c><c r="F2" t="str"><f>A2</f><v>ООО «Ромашка»</v></c></row>
<row r="3"><c r="A3" t="inlineStr"><is><t>Петров Пётр</t></is></c><c r="B3"><v>500100732259</v></c><c r="C3"><v>1700</v></c></row>
<row r="4"><c r="A4" t="inlineStr"><is><r><rPr><b/></rPr><t>Абрамов Дми</t></r><r><t>трий</t></r></is></c><c r="B4" t="s"><v>5</v></c><c r="C4"><f>SUM(C2:C3)</f><v>3200</v></c></row>
<row r="5"><c r="A5" t="s"><v>6</v></c><c r="B5" t="s"><v>7</v></c><c r="C5"><v>123456</v></c></row>
</sheetData><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/><headerFooter><oddHeader>&amp;LООО «Ромашка»&amp;RКонфиденциально</oddHeader></headerFooter><legacyDrawing r:id="rId1"/></worksheet>`;
const sheet2 = XD + `<worksheet ${SS}><sheetData>
<row r="1"><c r="A1" t="s"><v>7</v></c><c r="B1"><f>Клиенты!C4</f><v>3200</v></c></row>
<row r="2"><c r="A2" t="s"><v>2</v></c><c r="B2" s="1"><f>Клиенты!B2</f><v>7729654321</v></c></row>
<row r="3"><c r="A3" t="s"><v>4</v></c><c r="B3"><f>SUM(Клиенты!C2:C3)*2</f><v>6400</v></c></row>
</sheetData></worksheet>`;
const xfiles = {
  '[Content_Types].xml': XD + `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>`,
  '_rels/.rels': XD + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>`,
  'xl/workbook.xml': XD + `<workbook ${SS}><sheets><sheet name="Клиенты" sheetId="1" r:id="rId1"/><sheet name="Итоги" sheetId="2" r:id="rId2"/></sheets><definedNames><definedName name="Партнёр">"ООО «Ромашка»"</definedName><definedName name="Суммы">Клиенты!$C$2:$C$3</definedName></definedNames><calcPr calcId="191029"/></workbook>`,
  'xl/_rels/workbook.xml.rels': XD + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`,
  'xl/worksheets/sheet1.xml': sheet1,
  'xl/worksheets/_rels/sheet1.xml.rels': XD + `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" Target="../drawings/vmlDrawing1.vml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/></Relationships>`,
  'xl/worksheets/sheet2.xml': sheet2,
  'xl/sharedStrings.xml': XD + `<sst ${SS} count="${shared.length + 3}" uniqueCount="${shared.length}">${sst}</sst>`,
  'xl/styles.xml': XD + `<styleSheet ${SS}><numFmts count="1"><numFmt numFmtId="164" formatCode="0"/></numFmts><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`,
  'xl/comments1.xml': XD + `<comments ${SS}><authors><author>Иванов И.И.</author></authors><commentList><comment ref="B2" authorId="0"><text><r><rPr><b/><sz val="9"/></rPr><t>Иванов И.И.:</t></r><r><rPr><sz val="9"/></rPr><t xml:space="preserve"> проверить ИНН ООО «Ромашка»</t></r></text></comment></commentList></comments>`,
  'xl/drawings/vmlDrawing1.vml': `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype><v:shape id="_x0000_s1025" type="#_x0000_t202" style="position:absolute;margin-left:80pt;margin-top:2pt;width:120pt;height:60pt;z-index:1;visibility:hidden" fillcolor="#ffffe1"><v:textbox/><x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Row>1</x:Row><x:Column>1</x:Column></x:ClientData></v:shape></xml>`,
  'docProps/core.xml': files['docProps/core.xml'],
  'docProps/app.xml': XD + `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Microsoft Excel</Application><TitlesOfParts><vt:vector size="2" baseType="lpstr"><vt:lpstr>Клиенты</vt:lpstr><vt:lpstr>Итоги</vt:lpstr></vt:vector></TitlesOfParts><Company>ООО «Ромашка»</Company></Properties>`,
};

fs.writeFileSync(path.join(dir, 'synthetic.docx'), await buildZip(Object.entries(files)));
fs.writeFileSync(path.join(dir, 'synthetic.xlsx'), await buildZip(Object.entries(xfiles)));
console.log('собраны synthetic.docx и synthetic.xlsx');

// Файлы из LibreOffice: html → docx, csv → xlsx
const html = `<!doctype html><html><head><meta charset="utf-8"><title>Справка</title></head><body>
<h1>Справка для ООО «Ромашка»</h1>
<p>Выдана <b>Кузнецовой</b> Марине Олеговне, паспорт 45 18 № 603217, тел. +7 (916) 245-18-73.</p>
<table border="1"><tr><td>ФИО</td><td>Петров Пётр Ильич</td></tr><tr><td>ИНН</td><td>7729654321</td></tr></table>
<p>Почта для ответа: m.kuznetsova@sev-terminal.example</p></body></html>`;
const csv = `ФИО,ИНН,Сумма,Город\nПетров Пётр Ильич,7729654321,1500,Москва\nКузнецова Марина Олеговна,500100732259,1700,Тула\nАбрамов Дмитрий,7701234567,900,Москва\n`;
const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'maska-fx-'));
fs.writeFileSync(path.join(tmp, 'lo.html'), html);
fs.writeFileSync(path.join(tmp, 'lo.csv'), csv);
const prof = `-env:UserInstallation=file://${tmp}/profile`;
const lo = (args, from, to) => {
  try {
    execFileSync('soffice', [prof, '--headless', ...args, '--outdir', tmp, path.join(tmp, from)], { stdio: 'pipe', timeout: 120000 });
    fs.copyFileSync(path.join(tmp, to), path.join(dir, 'libreoffice.' + to.split('.').pop()));
    console.log('собран libreoffice.' + to.split('.').pop());
  } catch (e) { console.log(`soffice не собрал ${to} (для docx нужен модуль Writer):`, e.message.split('\n')[0]); }
};
lo(['--convert-to', 'docx:MS Word 2007 XML'], 'lo.html', 'lo.docx');
lo(['--infilter=CSV:44,34,76,1', '--convert-to', 'xlsx'], 'lo.csv', 'lo.xlsx');
fs.rmSync(tmp, { recursive: true, force: true });
