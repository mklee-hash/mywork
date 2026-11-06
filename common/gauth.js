/* Google 로그인 공용 (MY Work 로그인 화면 common/login.jsp · GitHub 보기 사이트 static.js)
 *   - Google Identity Services 토큰 (openid email + drive.file): 로그인한 계정 확인, 보기 키 파일 읽기/쓰기
 *   - 보기 키 = 내 Google Drive 의 "MY Work 보기 키 (지우지 마세요).txt"
 *     drive.file 권한이라 이 앱이 만든 파일만 보이고 읽을 수 있음. PC 의 MY Work 에 로그인할 때 맞춰 둠
 *   - 팝업이 막히지 않도록 prepare() 로 미리 준비하고, 버튼을 누른 그 순간 token() 을 부름
 */
window.MW_GAUTH = (function () {
	var KEYFILE = "MY Work 보기 키 (지우지 마세요).txt";
	var SCOPES = "openid email https://www.googleapis.com/auth/drive.file";
	var gisP = null, client = null, pending = null;

	function loadGis() {
		if (!gisP) gisP = new Promise(function (ok, no) {
			var s = document.createElement("script");
			s.src = "https://accounts.google.com/gsi/client";
			s.onload = ok;
			s.onerror = function () { gisP = null; no(new Error("Google 로그인 스크립트를 불러오지 못했습니다.")); };
			document.head.appendChild(s);
		});
		return gisP;
	}
	function done(err, tok) {
		var p = pending;
		pending = null;
		if (p) err ? p.no(err) : p.ok(tok);
	}
	function prepare(clientId, email) {
		return loadGis().then(function () {
			client = google.accounts.oauth2.initTokenClient({
				client_id: clientId, scope: SCOPES, hint: email,
				callback: function (r) { r.error ? done(new Error(r.error_description || r.error)) : done(null, r.access_token); },
				error_callback: function (e) { done(new Error(e && e.type === "popup_closed" ? "로그인 창을 닫았습니다." : (e && (e.message || e.type)) || "로그인하지 못했습니다.")); },
			});
		});
	}
	function token() {
		return new Promise(function (ok, no) {
			if (!client) return no(new Error("Google 로그인을 준비하는 중입니다. 잠시 뒤 다시 눌러 주세요."));
			pending = { ok: ok, no: no };
			client.requestAccessToken();
		});
	}
	function api(tok, url, opt) {
		opt = opt || {};
		opt.headers = Object.assign({ Authorization: "Bearer " + tok }, opt.headers || {});
		return fetch(url, opt).then(function (r) {
			if (!r.ok) return r.text().then(function (t) { throw new Error("Google " + r.status + ": " + t.slice(0, 200)); });
			return r;
		});
	}
	function email(tok) {
		return api(tok, "https://www.googleapis.com/oauth2/v3/userinfo").then(function (r) { return r.json(); }).then(function (j) { return String(j.email || "").toLowerCase(); });
	}
	function findKey(tok) {
		var q = encodeURIComponent("name='" + KEYFILE + "' and trashed=false");
		return api(tok, "https://www.googleapis.com/drive/v3/files?spaces=drive&fields=files(id)&q=" + q)
			.then(function (r) { return r.json(); }).then(function (j) { return (j.files || [])[0] || null; });
	}
	function readKey(tok) {
		return findKey(tok).then(function (f) {
			return f ? api(tok, "https://www.googleapis.com/drive/v3/files/" + f.id + "?alt=media").then(function (r) { return r.text(); }).then(function (t) { return t.trim(); }) : null;
		});
	}
	function writeKey(tok, keyB64) {
		return findKey(tok).then(function (f) {
			if (f) return api(tok, "https://www.googleapis.com/upload/drive/v3/files/" + f.id + "?uploadType=media", { method: "PATCH", headers: { "Content-Type": "text/plain" }, body: keyB64 });
			var b = "mwkey" + Date.now();
			var meta = { name: KEYFILE, mimeType: "text/plain", description: "MY Work GitHub 보기 사이트를 여는 키. 지워도 PC 의 MY Work 에 다시 로그인하면 새로 만들어집니다." };
			var body = "--" + b + "\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n" + JSON.stringify(meta) + "\r\n--" + b + "\r\nContent-Type: text/plain\r\n\r\n" + keyB64 + "\r\n--" + b + "--";
			return api(tok, "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart", { method: "POST", headers: { "Content-Type": "multipart/related; boundary=" + b }, body: body });
		});
	}
	return { prepare: prepare, token: token, email: email, readKey: readKey, writeKey: writeKey };
})();
