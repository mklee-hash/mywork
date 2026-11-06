/* 편집기 HTML -> .docx (zip.js 필요)
 * 제목(h1~h3, p.title), 문단, 굵게/기울임/밑줄/취소선/글자색/배경색, 링크, 번호·글머리 목록(중첩, start),
 * 표(colspan/rowspan/셀 배경), 그림, 줄바꿈을 옮김. 제목 번호(1. / 1.1)는 Word 제목 스타일 번호로 붙음.
 */
function buildDocx(root) {
	var REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/";
	var EMU = 9525, MAX_W = 640;   // 본문 폭(px) ≒ 17cm
	var BLOCK = /^(P|DIV|H[1-6]|LI|BLOCKQUOTE|PRE|SECTION|ARTICLE|HEADER|FOOTER|FIGURE|FIGCAPTION|ADDRESS|DL|DT|DD)$/;
	var rels = [], media = [], nums = [], jobs = [], picNo = 0;

	function rid() { return "rId" + (rels.length + 10); }
	function hex(c) {
		if (!c) return null;
		c = String(c).trim();
		var m = /^#([0-9a-f]{3}|[0-9a-f]{6})$/i.exec(c);
		if (m) return (m[1].length === 3 ? m[1].replace(/./g, "$&$&") : m[1]).toUpperCase();
		m = /^rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*([\d.]+))?\)$/.exec(c);
		if (m && (m[4] == null || +m[4] > 0)) return [m[1], m[2], m[3]].map(function (n) { return ("0" + (+n).toString(16)).slice(-2); }).join("").toUpperCase();
		var named = { yellow: "FFFF00", red: "FF0000", blue: "0000FF", green: "00FF00", cyan: "00FFFF", magenta: "FF00FF", black: "000000", white: "FFFFFF", gray: "808080", lightgray: "D3D3D3", darkblue: "000080", darkred: "800000", darkgreen: "008000", darkyellow: "808000", darkcyan: "008080", darkmagenta: "800080" };
		return named[c.toLowerCase()] || null;
	}

	// ---- 문단 만들기
	function Out() { this.xml = []; this.p = null; this.pendingNum = null; this.indent = 0; }
	Out.prototype.open = function (pPr, keep) {
		this.flush();
		this.p = { pPr: pPr || {}, runs: [], keep: !!keep };
	};
	Out.prototype.run = function (x) {
		if (!this.p) this.p = { pPr: {}, runs: [], keep: false };
		this.p.runs.push(x);
	};
	Out.prototype.flush = function () {
		var p = this.p;
		this.p = null;
		if (!p || (!p.runs.length && !p.keep)) return;
		var pr = "", o = p.pPr;
		if (this.pendingNum && !o.style) { o.num = this.pendingNum; this.pendingNum = null; }
		else if (this.indent && !o.style) o.ind = this.indent;
		if (o.style) pr += '<w:pStyle w:val="' + o.style + '"/>';
		if (o.num) pr += '<w:numPr><w:ilvl w:val="' + o.num.lvl + '"/><w:numId w:val="' + o.num.id + '"/></w:numPr>';
		if (o.hr) pr += '<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="CBD5E1"/></w:pBdr>';
		if (o.ind && !o.num) pr += '<w:ind w:left="' + o.ind + '"/>';
		if (o.jc) pr += '<w:jc w:val="' + o.jc + '"/>';
		this.xml.push("<w:p>" + (pr ? "<w:pPr>" + pr + "</w:pPr>" : "") + p.runs.join("") + "</w:p>");
	};
	Out.prototype.raw = function (x) { this.flush(); this.xml.push(x); };

	function rPr(f) {
		var s = "";
		if (f.href) s += '<w:rStyle w:val="Hyperlink"/>';
		if (f.b) s += "<w:b/>";
		if (f.i) s += "<w:i/>";
		if (f.s) s += "<w:strike/>";
		if (f.color) s += '<w:color w:val="' + f.color + '"/>';
		if (f.u && !f.href) s += '<w:u w:val="single"/>';
		if (f.bg) s += '<w:shd w:val="clear" w:color="auto" w:fill="' + f.bg + '"/>';
		return s ? "<w:rPr>" + s + "</w:rPr>" : "";
	}
	function textRun(text, f) {
		var parts = text.split("\t"), x = "";
		parts.forEach(function (t, i) {
			if (i) x += "<w:tab/>";
			if (t) x += '<w:t xml:space="preserve">' + xmlEsc(t) + "</w:t>";
		});
		x = "<w:r>" + rPr(f) + x + "</w:r>";
		if (f.href) {
			var id = rid();
			rels.push('<Relationship Id="' + id + '" Type="' + REL + 'hyperlink" Target="' + xmlEsc(f.href) + '" TargetMode="External"/>');
			x = '<w:hyperlink r:id="' + id + '" w:history="1">' + x + "</w:hyperlink>";
		}
		return x;
	}

	function imageRun(img) {
		var src = img.getAttribute("src") || "";
		if (!src) return "";
		var w = +img.getAttribute("width") || img.naturalWidth || img.width || 300;
		var h = +img.getAttribute("height") || (img.naturalWidth ? Math.round(w * img.naturalHeight / img.naturalWidth) : img.height) || 200;
		if (w > MAX_W) { h = Math.round(h * MAX_W / w); w = MAX_W; }
		var id = rid(), n = ++picNo, file = "image" + n;
		var m = { id: id, file: file, ext: "png" };
		media.push(m);
		rels.push(m);   // 확장자가 정해진 뒤 문자열로 바꿈
		jobs.push((src.indexOf("data:") === 0 ? Promise.resolve(dataUrlBytes(src)) : fetchBytes(img.src)).then(function (buf) {
			var b = new Uint8Array(buf);
			m.ext = b[0] === 0xFF && b[1] === 0xD8 ? "jpeg" : b[0] === 0x47 && b[1] === 0x49 ? "gif" : b[0] === 0x42 && b[1] === 0x4D ? "bmp" : "png";
			m.bytes = b;
		}));
		var cx = w * EMU, cy = h * EMU;
		return '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="' + cx + '" cy="' + cy + '"/>'
			+ '<wp:docPr id="' + n + '" name="그림 ' + n + '"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>'
			+ '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>'
			+ '<pic:nvPicPr><pic:cNvPr id="' + n + '" name="' + file + '"/><pic:cNvPicPr/></pic:nvPicPr>'
			+ '<pic:blipFill><a:blip r:embed="' + id + '"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>'
			+ '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>'
			+ "</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>";
	}

	function fmtOf(el, f) {
		var g = {}, k;
		for (k in f) g[k] = f[k];
		var t = el.tagName, st = el.style || {};
		if (t === "B" || t === "STRONG" || /^(bold|[6-9]00)$/.test(st.fontWeight)) g.b = 1;
		if (st.fontWeight === "normal" || st.fontWeight === "400") delete g.b;
		if (t === "I" || t === "EM" || st.fontStyle === "italic") g.i = 1;
		if (t === "U" || /underline/.test(st.textDecoration || st.textDecorationLine || "")) g.u = 1;
		if (t === "S" || t === "STRIKE" || t === "DEL" || /line-through/.test(st.textDecoration || st.textDecorationLine || "")) g.s = 1;
		if (t === "FONT" && el.getAttribute("color")) g.color = hex(el.getAttribute("color")) || g.color;
		if (st.color) g.color = hex(st.color) || g.color;
		if (t === "MARK") g.bg = "FFFF00";
		if (st.backgroundColor) g.bg = hex(st.backgroundColor) || g.bg;
		if (t === "A" && el.getAttribute("href") && !/^#/.test(el.getAttribute("href")) && !/^javascript:/i.test(el.getAttribute("href"))) g.href = el.href;
		return g;
	}
	function alignOf(el) {
		var a = (el.style && el.style.textAlign) || el.getAttribute("align") || "";
		return a === "center" ? "center" : a === "right" ? "right" : a === "justify" ? "both" : "";
	}
	function lastInBlock(node) {
		for (var n = node.nextSibling; n; n = n.nextSibling) {
			if (n.nodeType === 1 && !BLOCK.test(n.tagName)) return false;
			if (n.nodeType === 3 && /\S/.test(n.nodeValue)) return false;
			if (n.nodeType === 1) return true;
		}
		return true;
	}

	function walk(node, out, f) {
		for (var n = node.firstChild; n; n = n.nextSibling) {
			if (n.nodeType === 3) {
				var t = n.nodeValue.replace(/[\r\n]+/g, " ").replace(/ {2,}/g, " ");
				if (!out.p || !out.p.runs.length) t = t.replace(/^ +/, "");
				if (t) out.run(textRun(t, f));
				continue;
			}
			if (n.nodeType !== 1) continue;
			var tag = n.tagName;
			if (/^(SCRIPT|STYLE|META|LINK|TEMPLATE)$/.test(tag)) continue;
			if (tag === "BR") { if (!lastInBlock(n)) out.run("<w:r><w:br/></w:r>"); continue; }
			if (tag === "IMG") { out.run(imageRun(n)); continue; }
			if (tag === "HR") { out.open({ hr: 1 }, true); out.flush(); continue; }
			if (tag === "TABLE") { table(n, out, f); continue; }
			if (tag === "OL" || tag === "UL") { list(n, out, f); continue; }
			if (BLOCK.test(tag)) {
				var o = { jc: alignOf(n) };
				if (/^H[1-3]$/.test(tag)) o.style = "Heading" + tag.charAt(1);
				else if (/^H[4-6]$/.test(tag)) o.style = "Heading3";
				else if (n.classList.contains("title")) o.style = "Title";
				out.open(o, tag !== "DIV" && tag !== "LI");
				walk(n, out, fmtOf(n, f));
				out.flush();
				continue;
			}
			walk(n, out, fmtOf(n, f));
		}
	}

	// 목록 하나 = Word 번호 하나 (start 반영, 다음 목록과 번호가 섞이지 않음)
	function list(el, out, f) {
		out.flush();
		var depth = out.depth || 0, ordered = el.tagName === "OL";
		var type = ordered ? (el.classList.contains("circle") ? "circle" : el.getAttribute("type") || "1") : "bullet";
		var id = nums.length + 2;
		nums.push({ id: id, type: type, start: +el.getAttribute("start") || 1, lvl: depth });
		var saveIndent = out.indent;
		for (var li = el.firstElementChild; li; li = li.nextElementSibling) {
			if (li.tagName === "OL" || li.tagName === "UL") { out.depth = depth + 1; list(li, out, f); out.depth = depth; continue; }
			if (li.tagName !== "LI") continue;
			out.pendingNum = { id: id, lvl: Math.min(depth, 8) };
			out.indent = 425 * (depth + 1) + 360;
			out.depth = depth + 1;
			// 첫 내용이 문단(p 등)이면 그 문단이 번호를 받음
			var first = li.firstChild;
			while (first && first.nodeType === 3 && !/\S/.test(first.nodeValue)) first = first.nextSibling;
			out.open({ jc: alignOf(li) }, !(first && first.nodeType === 1 && BLOCK.test(first.tagName)));
			walk(li, out, fmtOf(li, f));
			out.flush();
			if (out.pendingNum) { out.open({}, true); out.flush(); }   // 빈 항목
			out.depth = depth;
		}
		out.indent = saveIndent;
	}

	function table(el, out, f) {
		out.flush();
		var rows = Array.prototype.filter.call(el.querySelectorAll("tr"), function (tr) { return tr.closest("table") === el; });
		var grid = [], ncol = 0;
		rows.forEach(function (tr, r) {
			grid[r] = grid[r] || [];
			var c = 0;
			Array.prototype.forEach.call(tr.children, function (td) {
				if (!/^T[DH]$/.test(td.tagName)) return;
				while (grid[r][c]) c++;
				var cs = Math.max(1, +td.getAttribute("colspan") || 1), rs = Math.max(1, +td.getAttribute("rowspan") || 1);
				for (var i = 0; i < rs; i++) {
					grid[r + i] = grid[r + i] || [];
					for (var j = 0; j < cs; j++) grid[r + i][c + j] = { td: td, top: i === 0, left: j === 0, cs: cs };
				}
				c += cs;
				ncol = Math.max(ncol, c);
			});
		});
		var colW = Math.floor(9638 / Math.max(1, ncol));
		var x = ['<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr><w:tblGrid>'];
		for (var c = 0; c < ncol; c++) x.push('<w:gridCol w:w="' + colW + '"/>');
		x.push("</w:tblGrid>");
		grid.forEach(function (row) {
			x.push("<w:tr>");
			for (var c = 0; c < ncol; c++) {
				var g = row[c];
				if (!g) { x.push('<w:tc><w:tcPr><w:tcW w:w="' + colW + '" w:type="dxa"/></w:tcPr><w:p/></w:tc>'); continue; }
				if (!g.left) continue;
				var pr = '<w:tcW w:w="' + colW * g.cs + '" w:type="dxa"/>' + (g.cs > 1 ? '<w:gridSpan w:val="' + g.cs + '"/>' : "");
				var rs = +g.td.getAttribute("rowspan") || 1;
				if (rs > 1) pr += g.top ? '<w:vMerge w:val="restart"/>' : "<w:vMerge/>";
				var bg = hex(g.td.style.backgroundColor) || hex(g.td.getAttribute("bgcolor"));
				if (bg) pr += '<w:shd w:val="clear" w:color="auto" w:fill="' + bg + '"/>';
				var inner = "<w:p/>";
				if (g.top) {
					var o = new Out();
					walk(g.td, o, fmtOf(g.td, f));
					o.flush();
					if (o.xml.length && /<\/w:p>$|<w:p\/>$/.test(o.xml[o.xml.length - 1])) inner = o.xml.join("");
					else inner = o.xml.join("") + "<w:p/>";
				}
				x.push("<w:tc><w:tcPr>" + pr + "</w:tcPr>" + inner + "</w:tc>");
			}
			x.push("</w:tr>");
		});
		x.push("</w:tbl>");
		out.raw(x.join(""));
	}

	var body = new Out();
	walk(root, body, {});
	body.flush();

	// ---- numbering.xml
	function lvlXml(type, i, start) {
		var fmt = { "1": "decimal", "A": "upperLetter", "a": "lowerLetter", "I": "upperRoman", "i": "lowerRoman", circle: "decimalEnclosedCircle", bullet: "bullet" }[type] || "decimal";
		var text = fmt === "bullet" ? ["•", "◦", "▪"][i % 3] : fmt === "decimalEnclosedCircle" ? "%" + (i + 1) : "%" + (i + 1) + ".";
		var left = 425 * (i + 1) + 360;
		return '<w:lvl w:ilvl="' + i + '"><w:start w:val="' + (start || 1) + '"/><w:numFmt w:val="' + fmt + '"/><w:lvlText w:val="' + text + '"/><w:lvlJc w:val="left"/>'
			+ '<w:pPr><w:ind w:left="' + left + '" w:hanging="360"/></w:pPr>' + (fmt === "bullet" ? '<w:rPr><w:rFonts w:ascii="맑은 고딕" w:hAnsi="맑은 고딕"/></w:rPr>' : "") + "</w:lvl>";
	}
	var absXml = ['<w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="multilevel"/>'
		+ '<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="Heading1"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="425" w:hanging="425"/></w:pPr></w:lvl>'
		+ '<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="Heading2"/><w:lvlText w:val="%1.%2"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="567" w:hanging="567"/></w:pPr></w:lvl>'
		+ "</w:abstractNum>"];
	var numXml = ['<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>'];
	nums.forEach(function (n, k) {
		var a = ['<w:abstractNum w:abstractNumId="' + (k + 1) + '"><w:multiLevelType w:val="hybridMultilevel"/>'];
		for (var i = 0; i < 9; i++) a.push(lvlXml(n.type, i, i === n.lvl ? n.start : 1));
		absXml.push(a.join("") + "</w:abstractNum>");
		numXml.push('<w:num w:numId="' + n.id + '"><w:abstractNumId w:val="' + (k + 1) + '"/></w:num>');
	});

	var W_NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
	var docXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document ' + W_NS
		+ ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">'
		+ "<w:body>" + body.xml.join("")
		+ '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1418" w:right="1134" w:bottom="1418" w:left="1134" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr>'
		+ "</w:body></w:document>";

	var font = '<w:rFonts w:ascii="맑은 고딕" w:eastAsia="맑은 고딕" w:hAnsi="맑은 고딕" w:cs="맑은 고딕"/>';
	function pStyle(id, name, pPr, rPr2, extra) {
		return '<w:style w:type="paragraph" w:styleId="' + id + '"><w:name w:val="' + name + '"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/>' + (extra || "") + "<w:qFormat/>"
			+ (pPr ? "<w:pPr>" + pPr + "</w:pPr>" : "") + (rPr2 ? "<w:rPr>" + rPr2 + "</w:rPr>" : "") + "</w:style>";
	}
	var stylesXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:styles ' + W_NS + ">"
		+ "<w:docDefaults><w:rPrDefault><w:rPr>" + font + '<w:sz w:val="20"/><w:szCs w:val="20"/><w:lang w:val="en-US" w:eastAsia="ko-KR"/></w:rPr></w:rPrDefault>'
		+ '<w:pPrDefault><w:pPr><w:spacing w:after="60" w:line="288" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
		+ '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>'
		+ pStyle("Title", "Title", '<w:pBdr><w:bottom w:val="single" w:sz="12" w:space="4" w:color="0F172A"/></w:pBdr><w:spacing w:after="300"/>', '<w:b/><w:sz w:val="44"/>')
		+ pStyle("Heading1", "heading 1", '<w:keepNext/><w:numPr><w:numId w:val="1"/></w:numPr><w:spacing w:before="480" w:after="160"/><w:outlineLvl w:val="0"/>', '<w:b/><w:sz w:val="32"/>')
		+ pStyle("Heading2", "heading 2", '<w:keepNext/><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr><w:spacing w:before="300" w:after="120"/><w:outlineLvl w:val="1"/>', '<w:b/><w:sz w:val="26"/>')
		+ pStyle("Heading3", "heading 3", '<w:keepNext/><w:spacing w:before="200" w:after="80"/><w:outlineLvl w:val="2"/>', '<w:b/><w:sz w:val="22"/>')
		+ '<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/><w:uiPriority w:val="1"/><w:semiHidden/></w:style>'
		+ '<w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:basedOn w:val="DefaultParagraphFont"/><w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr></w:style>'
		+ '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:semiHidden/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>'
		+ '<w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:basedOn w:val="TableNormal"/><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:tblPr><w:tblBorders>'
		+ '<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>'
		+ '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr></w:style>'
		+ "</w:styles>";

	return Promise.all(jobs).then(function () {
		var zip = new Zip();
		var relXml = rels.map(function (r) {
			return typeof r === "string" ? r : '<Relationship Id="' + r.id + '" Type="' + REL + 'image" Target="media/' + r.file + "." + r.ext + '"/>';
		});
		relXml.push('<Relationship Id="rId1" Type="' + REL + 'styles" Target="styles.xml"/>');
		relXml.push('<Relationship Id="rId2" Type="' + REL + 'numbering" Target="numbering.xml"/>');
		zip.add("[Content_Types].xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
			+ '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
			+ '<Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/><Default Extension="gif" ContentType="image/gif"/><Default Extension="bmp" ContentType="image/bmp"/>'
			+ '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
			+ '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
			+ '<Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/>'
			+ "</Types>");
		zip.add("_rels/.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="' + REL + 'officeDocument" Target="word/document.xml"/></Relationships>');
		zip.add("word/document.xml", docXml);
		zip.add("word/styles.xml", stylesXml);
		zip.add("word/numbering.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:numbering ' + W_NS + ">" + absXml.join("") + numXml.join("") + "</w:numbering>");
		zip.add("word/_rels/document.xml.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + relXml.join("") + "</Relationships>");
		media.forEach(function (m) { zip.add("word/media/" + m.file + "." + m.ext, m.bytes); });
		return zip.blob("application/vnd.openxmlformats-officedocument.wordprocessingml.document");
	});
}
