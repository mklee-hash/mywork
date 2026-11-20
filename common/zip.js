/* 최소 ZIP 작성기 (무압축 STORE) - xlsx/docx 다운로드용 */
var Zip = (function () {
	var CRC_TABLE = (function () {
		var t = new Uint32Array(256);
		for (var n = 0; n < 256; n++) {
			var c = n;
			for (var k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
			t[n] = c >>> 0;
		}
		return t;
	})();

	function crc32(bytes) {
		var c = 0xFFFFFFFF;
		for (var i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xFF] ^ (c >>> 8);
		return (c ^ 0xFFFFFFFF) >>> 0;
	}

	function Zip() { this.files = []; }

	// data: 문자열(UTF-8 로 저장) 또는 Uint8Array / ArrayBuffer
	Zip.prototype.add = function (name, data) {
		var bytes = typeof data === "string" ? new TextEncoder().encode(data) : new Uint8Array(data);
		this.files.push({ name: new TextEncoder().encode(name), data: bytes, crc: crc32(bytes) });
		return this;
	};

	Zip.prototype.blob = function (mime) {
		var d = new Date();
		var time = (d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1);
		var date = ((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate();
		var parts = [], central = [], offset = 0;
		this.files.forEach(function (f) {
			var h = new DataView(new ArrayBuffer(30));
			h.setUint32(0, 0x04034b50, true);
			h.setUint16(4, 20, true);
			h.setUint16(6, 0x0800, true);          // 파일명 UTF-8
			h.setUint16(8, 0, true);               // STORE
			h.setUint16(10, time, true);
			h.setUint16(12, date, true);
			h.setUint32(14, f.crc, true);
			h.setUint32(18, f.data.length, true);
			h.setUint32(22, f.data.length, true);
			h.setUint16(26, f.name.length, true);
			parts.push(h.buffer, f.name, f.data);

			var c = new DataView(new ArrayBuffer(46));
			c.setUint32(0, 0x02014b50, true);
			c.setUint16(4, 20, true);
			c.setUint16(6, 20, true);
			c.setUint16(8, 0x0800, true);
			c.setUint16(10, 0, true);
			c.setUint16(12, time, true);
			c.setUint16(14, date, true);
			c.setUint32(16, f.crc, true);
			c.setUint32(20, f.data.length, true);
			c.setUint32(24, f.data.length, true);
			c.setUint16(28, f.name.length, true);
			c.setUint32(42, offset, true);
			central.push(c.buffer, f.name);
			offset += 30 + f.name.length + f.data.length;
		});
		var size = central.reduce(function (s, b) { return s + (b.byteLength || b.length); }, 0);
		var e = new DataView(new ArrayBuffer(22));
		e.setUint32(0, 0x06054b50, true);
		e.setUint16(8, this.files.length, true);
		e.setUint16(10, this.files.length, true);
		e.setUint32(12, size, true);
		e.setUint32(16, offset, true);
		return new Blob(parts.concat(central, [e.buffer]), { type: mime || "application/zip" });
	};

	return Zip;
})();

/* 공통 유틸 */
function xmlEsc(s) {
	return String(s == null ? "" : s)
		.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g, "")
		.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function downloadBlob(blob, name) {
	var a = document.createElement("a");
	a.href = URL.createObjectURL(blob);
	a.download = name;
	document.body.appendChild(a);
	a.click();
	setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
}

// 원본 파일 이름(확장자 제외)
function sourceBase(src) {
	return String(src || "download").split(/[\\\/]/).pop().replace(/\.[^.]+$/, "");
}

function fetchBytes(url) {
	return fetch(url, { cache: "no-store" }).then(function (r) {
		if (!r.ok) throw new Error(url + " " + r.status);
		return r.arrayBuffer();
	});
}

function dataUrlBytes(url) {
	var b64 = url.slice(url.indexOf(",") + 1);
	var bin = atob(b64), out = new Uint8Array(bin.length);
	for (var i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
	return out;
}
