/* 시트 데이터 -> .xlsx (zip.js 필요)
 * wb = { sheets: [{ name, cols:[px], rows:[[cell|null]], merges:[[r1,c1,r2,c2]], images:[{f,r,c,dx,dy,w,h}] }] }
 * cell = { v, b, i, u, fg, bg, ha, va, wrap, bd, link }
 * imgUrl(name) -> 그림 URL
 */
function buildXlsx(wb, imgUrl) {
	var EMU = 9525;
	var NS = 'xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
	var REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";

	function colName(c) {
		var s = "";
		for (c++; c > 0; c = Math.floor((c - 1) / 26)) s = String.fromCharCode(65 + (c - 1) % 26) + s;
		return s;
	}
	function argb(hex) { return "FF" + String(hex).replace("#", "").toUpperCase(); }

	// ---- 서식 표 (중복 제거)
	var fonts = ['<font><sz val="10"/><name val="맑은 고딕"/><family val="3"/><charset val="129"/></font>'], fontIdx = { "": 0 };
	var fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'], fillIdx = { "": 0 };
	var xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'], xfIdx = { "": 0 };

	function styleOf(cell) {
		if (!cell) return 0;
		var fg = cell.fg || (cell.link ? "#0563C1" : ""), u = cell.u || (cell.link && !cell.fg);
		var fk = [cell.b ? 1 : 0, cell.i ? 1 : 0, u ? 1 : 0, fg].join("|");
		if (fk === "0|0|0|") fk = "";
		if (!(fk in fontIdx)) {
			fontIdx[fk] = fonts.length;
			fonts.push("<font>" + (cell.b ? "<b/>" : "") + (cell.i ? "<i/>" : "") + (u ? "<u/>" : "")
				+ '<sz val="10"/>' + (fg ? '<color rgb="' + argb(fg) + '"/>' : "") + '<name val="맑은 고딕"/><family val="3"/><charset val="129"/></font>');
		}
		var bk = cell.bg || "";
		if (!(bk in fillIdx)) {
			fillIdx[bk] = fills.length;
			fills.push('<fill><patternFill patternType="solid"><fgColor rgb="' + argb(bk) + '"/><bgColor indexed="64"/></patternFill></fill>');
		}
		var al = (cell.ha ? ' horizontal="' + (cell.ha === "center" ? "center" : cell.ha === "right" ? "right" : "left") + '"' : "")
			+ (cell.va ? ' vertical="' + (cell.va === "middle" ? "center" : cell.va === "top" ? "top" : "bottom") + '"' : "")
			+ (cell.wrap || /\n/.test(cell.v || "") ? ' wrapText="1"' : "");
		var xk = [fontIdx[fk], fillIdx[bk], cell.bd ? 1 : 0, al].join("|");
		if (xk === "0|0|0|") return 0;
		if (!(xk in xfIdx)) {
			xfIdx[xk] = xfs.length;
			xfs.push('<xf numFmtId="0" fontId="' + fontIdx[fk] + '" fillId="' + fillIdx[bk] + '" borderId="' + (cell.bd ? 1 : 0) + '" xfId="0"'
				+ (fontIdx[fk] ? ' applyFont="1"' : "") + (fillIdx[bk] ? ' applyFill="1"' : "") + (cell.bd ? ' applyBorder="1"' : "")
				+ (al ? ' applyAlignment="1"><alignment' + al + "/></xf>" : "/>"));
		}
		return xfIdx[xk];
	}

	var zip = new Zip();
	var ctypes = [], wbRels = [], sheetsXml = [], jobs = [], imgNo = 0;

	wb.sheets.forEach(function (sh, si) {
		var n = si + 1, rows = sh.rows || [], cols = sh.cols || [];
		var merged = {};
		(sh.merges || []).forEach(function (m) {
			for (var r = m[0]; r <= m[2]; r++) for (var c = m[1]; c <= m[3]; c++) if (r !== m[0] || c !== m[1]) merged[r + "," + c] = 1;
		});
		var links = [], rels = [], x = [];
		x.push('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet ' + NS + ">");
		x.push('<sheetViews><sheetView workbookViewId="0"' + (si === 0 ? ' tabSelected="1"' : "") + "/></sheetViews>");
		x.push('<sheetFormatPr defaultRowHeight="16.5"/>');
		if (cols.length) {
			x.push("<cols>");
			cols.forEach(function (px, c) {
				x.push('<col min="' + (c + 1) + '" max="' + (c + 1) + '" width="' + Math.max(0.5, (px - 5) / 7).toFixed(2) + '" customWidth="1"/>');
			});
			x.push("</cols>");
		}
		x.push("<sheetData>");
		rows.forEach(function (row, r) {
			var cells = [], lines = 1;
			(row || []).forEach(function (cell, c) {
				if (!cell) return;
				var s = merged[r + "," + c] && !cell.bd && !cell.bg ? 0 : styleOf(cell);
				var v = cell.v == null ? "" : String(cell.v);
				if (!v && !s) return;
				var ref = colName(c) + (r + 1);
				if (v) {
					cells.push('<c r="' + ref + '"' + (s ? ' s="' + s + '"' : "") + ' t="inlineStr"><is><t xml:space="preserve">' + xmlEsc(v) + "</t></is></c>");
					// 줄바꿈이 있는 셀은 행 높이를 늘려 줌 (Excel 은 열 때 자동 맞춤을 안 함)
					if (!merged[r + "," + c]) {
						var chars = Math.max(1, Math.floor(((cols[c] || 64) - 5) / 7));
						var ln = v.split("\n").reduce(function (s2, part) {
							return s2 + (cell.wrap ? Math.max(1, Math.ceil(part.replace(/[^\x00-\xff]/g, "aa").length / chars)) : 1);
						}, 0);
						var isMergeTop = (sh.merges || []).some(function (m) { return m[0] === r && m[1] === c && m[2] > m[0]; });
						if (!isMergeTop) lines = Math.max(lines, ln);
					}
				} else {
					cells.push('<c r="' + ref + '" s="' + s + '"/>');
				}
				if (cell.link) {
					var hash = cell.link.charAt(0) === "#";
					if (hash) links.push('<hyperlink ref="' + ref + '" location="' + xmlEsc(cell.link.slice(1)) + '"/>');
					else {
						var id = "rId" + (rels.length + 1);
						rels.push('<Relationship Id="' + id + '" Type="' + REL + 'hyperlink" Target="' + xmlEsc(cell.link) + '" TargetMode="External"/>');
						links.push('<hyperlink ref="' + ref + '" r:id="' + id + '"/>');
					}
				}
			});
			if (cells.length || lines > 1)
				x.push('<row r="' + (r + 1) + '"' + (lines > 1 ? ' ht="' + (lines * 16.5) + '" customHeight="1"' : "") + ">" + cells.join("") + "</row>");
		});
		x.push("</sheetData>");
		if ((sh.merges || []).length) {
			x.push('<mergeCells count="' + sh.merges.length + '">');
			sh.merges.forEach(function (m) { x.push('<mergeCell ref="' + colName(m[1]) + (m[0] + 1) + ":" + colName(m[3]) + (m[2] + 1) + '"/>'); });
			x.push("</mergeCells>");
		}
		if (links.length) x.push("<hyperlinks>" + links.join("") + "</hyperlinks>");
		x.push('<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>');

		var imgs = sh.images || [];
		if (imgs.length) {
			var did = "rId" + (rels.length + 1);
			rels.push('<Relationship Id="' + did + '" Type="' + REL + 'drawing" Target="../drawings/drawing' + n + '.xml"/>');
			x.push('<drawing r:id="' + did + '"/>');
			ctypes.push('<Override PartName="/xl/drawings/drawing' + n + '.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>');
			var d = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'];
			var drels = [];
			imgs.forEach(function (im, k) {
				var ext = (im.f.split(".").pop() || "png").toLowerCase().replace("jpg", "jpeg");
				var media = "image" + (++imgNo) + "." + ext;
				jobs.push(fetchBytes(imgUrl(im.f)).then(function (b) { zip.add("xl/media/" + media, b); }));
				drels.push('<Relationship Id="rId' + (k + 1) + '" Type="' + REL + 'image" Target="../media/' + media + '"/>');
				var cx = Math.round(im.w * EMU), cy = Math.round(im.h * EMU);
				d.push('<xdr:oneCellAnchor><xdr:from><xdr:col>' + im.c + "</xdr:col><xdr:colOff>" + Math.round((im.dx || 0) * EMU) + "</xdr:colOff><xdr:row>" + im.r
					+ "</xdr:row><xdr:rowOff>" + Math.round((im.dy || 0) * EMU) + '</xdr:rowOff></xdr:from><xdr:ext cx="' + cx + '" cy="' + cy + '"/>'
					+ '<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="' + (k + 2) + '" name="그림 ' + (k + 1) + '"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr>'
					+ '<xdr:blipFill><a:blip r:embed="rId' + (k + 1) + '"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>'
					+ '<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>');
			});
			d.push("</xdr:wsDr>");
			zip.add("xl/drawings/drawing" + n + ".xml", d.join(""));
			zip.add("xl/drawings/_rels/drawing" + n + ".xml.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + drels.join("") + "</Relationships>");
		}
		x.push("</worksheet>");
		zip.add("xl/worksheets/sheet" + n + ".xml", x.join(""));
		if (rels.length)
			zip.add("xl/worksheets/_rels/sheet" + n + ".xml.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + rels.join("") + "</Relationships>");

		ctypes.push('<Override PartName="/xl/worksheets/sheet' + n + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>');
		wbRels.push('<Relationship Id="rId' + n + '" Type="' + REL + 'worksheet" Target="worksheets/sheet' + n + '.xml"/>');
		// 시트 이름: 31자, []:*?/\ 불가, 중복 불가
		var nm = String(sh.name || "Sheet" + n).replace(/[\[\]:*?\/\\]/g, "_").slice(0, 31) || "Sheet" + n;
		sheetsXml.push('<sheet name="' + xmlEsc(nm) + '" sheetId="' + n + '" r:id="rId' + n + '"/>');
	});

	var nSheets = wb.sheets.length;
	wbRels.push('<Relationship Id="rId' + (nSheets + 1) + '" Type="' + REL + 'styles" Target="styles.xml"/>');
	zip.add("xl/workbook.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook ' + NS + "><sheets>" + sheetsXml.join("") + "</sheets></workbook>");
	zip.add("xl/_rels/workbook.xml.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + wbRels.join("") + "</Relationships>");
	zip.add("xl/styles.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
		+ '<fonts count="' + fonts.length + '">' + fonts.join("") + "</fonts>"
		+ '<fills count="' + fills.length + '">' + fills.join("") + "</fills>"
		+ '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>'
		+ '<border><left style="thin"><color indexed="64"/></left><right style="thin"><color indexed="64"/></right><top style="thin"><color indexed="64"/></top><bottom style="thin"><color indexed="64"/></bottom><diagonal/></border></borders>'
		+ '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
		+ '<cellXfs count="' + xfs.length + '">' + xfs.join("") + "</cellXfs>"
		+ '<cellStyles count="1"><cellStyle name="표준" xfId="0" builtinId="0"/></cellStyles></styleSheet>');
	zip.add("_rels/.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="' + REL + 'officeDocument" Target="xl/workbook.xml"/></Relationships>');
	zip.add("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
		+ '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
		+ '<Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="gif" ContentType="image/gif"/>'
		+ '<Default Extension="webp" ContentType="image/webp"/><Default Extension="bmp" ContentType="image/bmp"/>'
		+ '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
		+ '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
		+ ctypes.join("") + "</Types>");

	return Promise.all(jobs).then(function () {
		return zip.blob("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
	});
}
