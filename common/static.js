/* GitHub 보기 사이트 전용 (common/publish.jsp 가 각 HTML 의 <head> 맨 앞에 넣음)
 *   - data.jsp 요청을 암호화된 파일 d/<이름>.bin 읽기로 바꿈 (AES-GCM, 키 = PBKDF2-SHA256(비밀번호))
 *   - 파일 이름 = SHA-256(salt + "|" + 요청 키) 앞 32자리, 요청 키 = "space=..&doc=..[&img=..]" 또는 "space=work&list=1|cats=1"
 *   - 수정/저장 버튼은 숨김 (보기 전용)
 *   - 여는 방법: Google 로그인(mklee@forcs.com, 키는 내 Drive 의 보기 키 파일 — gauth.js) 또는 비상용 비밀번호
 *   - 비밀번호는 저장하지 않고, 풀린 키만 이 브라우저(sessionStorage, [기억] 시 localStorage)에 둠
 */
(function () {
	var me = document.currentScript;
	var ROOT = me.getAttribute("data-root") || "";
	var KEY = "mkwork.key";
	var origFetch = window.fetch.bind(window);
	window.MK_STATIC = true;
	window.MK_EXT = ".html";

	// 보기 전용: 수정 도구 숨김
	var css = "#bSave,#tools,#bHist,#imgbar,#pEdit,#pFoot,#bNew,#bPub,.rw,[data-add]{display:none!important}.card li .star{pointer-events:none}"
		+ "#tTitle,#tCat,#tStatus{pointer-events:none;background:#f8fafc}"
		+ "#mkLock{position:fixed;inset:0;z-index:1000;display:flex;align-items:center;justify-content:center;background:#0f172a;font-family:'맑은 고딕','Malgun Gothic',sans-serif}"
		+ "#mkLock form{width:360px;max-width:calc(100% - 32px);padding:26px 24px;background:#fff;border-radius:12px;font-size:14px;color:#0f172a}"
		+ "#mkLock h2{margin:0 0 6px;font-size:20px}#mkLock p{margin:0 0 14px;color:#64748b;font-size:13px;line-height:1.6}"
		+ "#mkLock input[type=password]{width:100%;height:36px;padding:0 10px;border:1px solid #cbd5e1;border-radius:6px;font:inherit;box-sizing:border-box}"
		+ "#mkLock label{display:flex;gap:6px;align-items:center;margin:10px 0 14px;font-size:13px;color:#475569}"
		+ "#mkLock button{width:100%;height:36px;border:0;border-radius:6px;background:#0d9488;color:#fff;font:inherit;cursor:pointer}"
		+ "#mkLock .err{min-height:18px;margin-top:8px;color:#b91c1c;font-size:13px}";
	var st = document.createElement("style");
	st.textContent = css;
	document.head.appendChild(st);
	document.addEventListener("DOMContentLoaded", function () {
		var ed = document.getElementById("ed");
		if (ed) ed.setAttribute("contenteditable", "false");
		["tTitle", "tCat", "tStatus"].forEach(function (id) { var el = document.getElementById(id); if (el) el.disabled = true; });
		// 탭 아이콘: 보기 사이트는 보라색 가방 (내 PC 는 청록색)
		var ic = document.querySelector('link[rel="icon"]');
		if (!ic) { ic = document.createElement("link"); ic.rel = "icon"; ic.type = "image/svg+xml"; document.head.appendChild(ic); }
		ic.href = ROOT + "common/favicon-view.svg";
	});

	function b64(buf) { var s = "", a = new Uint8Array(buf); for (var i = 0; i < a.length; i++) s += String.fromCharCode(a[i]); return btoa(s); }
	function unb64(s) { var b = atob(s), a = new Uint8Array(b.length); for (var i = 0; i < b.length; i++) a[i] = b.charCodeAt(i); return a; }
	function hex(buf) { return Array.prototype.map.call(new Uint8Array(buf), function (x) { return (x < 16 ? "0" : "") + x.toString(16); }).join(""); }
	function store() { try { return localStorage.getItem(KEY) || sessionStorage.getItem(KEY); } catch (e) { return null; } }

	var metaP = origFetch(ROOT + "meta.json", { cache: "no-store" }).then(function (r) { return r.json(); });
	metaP.then(function (m) {
		window.MK_META = m;
		var el = document.getElementById("pubAt");
		if (el) el.textContent = "올린 시각 " + m.publishedAt;
	});

	// 기억한 키가 지금 올라간 내용과 맞는지 meta.check 로 확인 (PC 에서 비밀번호를 바꾸면 예전 키는 버림)
	var keyP = null;
	function getKey() {
		if (!keyP) {
			var raw = store();
			keyP = !raw ? Promise.reject(new Error("잠겨 있습니다. 비밀번호를 넣어 주세요."))
				: crypto.subtle.importKey("raw", unb64(raw), "AES-GCM", false, ["decrypt"]).then(function (key) {
					return metaP.then(function (m) { return decrypt(key, unb64(m.check)); }).then(function () { return key; }, function () {
						try { localStorage.removeItem(KEY); sessionStorage.removeItem(KEY); } catch (e) {}
						throw new Error("비밀번호가 바뀌었습니다. 새 비밀번호를 넣어 주세요.");
					});
				});
		}
		return keyP;
	}
	function decrypt(key, bytes) {
		return crypto.subtle.decrypt({ name: "AES-GCM", iv: bytes.slice(0, 12) }, key, bytes.slice(12));
	}

	// data.jsp?... -> 요청 키 (data.jsp 와 같은 규칙: space 가 없으면 server)
	function reqKey(url) {
		var u = new URL(url, location.href), q = u.searchParams;
		var k = "space=" + (q.get("space") || "server");
		if (q.get("list") != null) return k + "&list=1";
		if (q.get("cats") != null) return k + "&cats=1";
		k += "&doc=" + (q.get("doc") || "");
		if (q.get("img") != null) k += "&img=" + q.get("img");
		return k;
	}
	function load(url) {
		var k = reqKey(url);
		return Promise.all([metaP, getKey()]).then(function (x) {
			return crypto.subtle.digest("SHA-256", new TextEncoder().encode(x[0].salt + "|" + k)).then(function (h) {
				return origFetch(ROOT + "d/" + hex(h).slice(0, 32) + ".bin");
			}).then(function (r) {
				if (!r.ok) throw new Error("없는 항목입니다.");
				return r.arrayBuffer();
			}).then(function (buf) { return decrypt(x[1], new Uint8Array(buf)); });
		}).then(function (plain) {
			var img = /&img=(.+)$/.exec(k);
			if (img) {
				var ext = img[1].split(".").pop().toLowerCase();
				return new Response(plain, { status: 200, headers: { "Content-Type": "image/" + (ext === "jpg" ? "jpeg" : ext) } });
			}
			return new Response(plain, { status: 200, headers: { "Content-Type": "application/json; charset=UTF-8", "X-Ver": "" } });
		}, function (e) {
			return new Response(JSON.stringify({ error: e.message }), { status: 404, headers: { "Content-Type": "application/json" } });
		});
	}

	window.fetch = function (input, init) {
		var url = typeof input === "string" ? input : input.url;
		if (!/data\.jsp\?/.test(url)) return origFetch(input, init);
		var q = new URL(url, location.href).searchParams;
		if (init && init.method && init.method !== "GET") return Promise.resolve(new Response(JSON.stringify({ error: "보기 전용입니다. 수정은 내 PC 에서 하세요." }), { status: 403 }));
		if (q.get("backups") != null) return Promise.resolve(new Response("[]", { status: 200 }));
		return load(url);
	};

	// <img src="data.jsp?..."> 는 fetch 를 거치지 않으므로 풀어서 blob 주소로 바꿈
	var blobs = {};
	function fixImg(im) {
		var s = im.getAttribute("src");
		if (!s || !/data\.jsp\?/.test(s)) return;
		im.setAttribute("data-src", s);
		im.removeAttribute("src");
		var k = reqKey(s);
		(blobs[k] = blobs[k] || load(s).then(function (r) { return r.ok ? r.blob() : null; }).then(function (b) { return b && URL.createObjectURL(b); }))
			.then(function (u) { if (u) im.src = u; });
	}
	new MutationObserver(function (ms) {
		ms.forEach(function (m) {
			if (m.type === "attributes") { if (m.target.tagName === "IMG") fixImg(m.target); return; }
			m.addedNodes.forEach(function (n) {
				if (n.nodeType !== 1) return;
				if (n.tagName === "IMG") fixImg(n);
				else n.querySelectorAll && n.querySelectorAll("img").forEach(fixImg);
			});
		});
	}).observe(document.documentElement, { childList: true, subtree: true, attributes: true, attributeFilter: ["src"] });
	// window.open(imgUrl(...)) 같은 새 창 그림 열기
	var origOpen = window.open;
	window.open = function (u) {
		if (typeof u === "string" && /data\.jsp\?/.test(u)) {
			var w = origOpen.call(window, "", "_blank");
			load(u).then(function (r) { return r.blob(); }).then(function (b) { w.location.href = URL.createObjectURL(b); });
			return w;
		}
		return origOpen.apply(window, arguments);
	};

	// 잠금 / 풀기 (맨 바깥 창에서만 비밀번호를 받음)
	window.MK_LOCK = function () {
		try { localStorage.removeItem(KEY); sessionStorage.removeItem(KEY); } catch (e) {}
		location.reload();
	};
	if (window.top === window) {
		document.addEventListener("DOMContentLoaded", function () {
			if (!store()) showLock("");
			else getKey().catch(function (e) { showLock(e.message); });
		});
	}
	// 잠금 화면: Google 로그인(mklee@forcs.com) → 내 Drive 의 보기 키로 열기. 비밀번호는 비상용으로 접어 둠
	function keepStore() { return document.getElementById("mkKeep").checked ? localStorage : sessionStorage; }
	function useRaw(rawB64) {
		return crypto.subtle.importKey("raw", unb64(rawB64), "AES-GCM", true, ["decrypt"]).then(function (key) {
			return metaP.then(function (m) { return decrypt(key, unb64(m.check)); }).then(function () { return rawB64; });
		});
	}
	function showLock(why) {
		var d = document.createElement("div");
		d.id = "mkLock";
		d.innerHTML = '<form><h2>MY Work</h2><p id="mkWhy">' + (why || "보기 전용 사이트입니다.") + "</p>"
			+ '<button type="button" id="mkG" style="display:none;background:#fff;color:#0f172a;border:1px solid #cbd5e1;font-weight:bold">Google 계정으로 열기</button>'
			+ '<label><input type="checkbox" id="mkKeep"> 이 브라우저에서 기억</label>'
			+ '<details id="mkPwBox"><summary style="cursor:pointer;color:#64748b;font-size:13px;margin-bottom:8px">비밀번호로 열기</summary>'
			+ '<input type="password" id="mkPw" placeholder="비밀번호" autocomplete="current-password">'
			+ '<button type="submit" style="margin-top:8px">열기</button></details><div class="err" id="mkErr"></div></form>';
		document.body.appendChild(d);
		var err = document.getElementById("mkErr"), gBtn = document.getElementById("mkG");
		metaP.then(function (m) {
			if (!m.google) { document.getElementById("mkPwBox").open = true; document.getElementById("mkWhy").textContent = why || "보기 전용 사이트입니다. 비밀번호를 넣어야 내용이 보입니다."; return; }
			document.getElementById("mkWhy").innerHTML = (why ? why + "<br>" : "") + "<b>" + m.google.email + "</b> 구글 계정으로 여세요.";
			var s = document.createElement("script");
			s.src = ROOT + "common/gauth.js";
			s.onload = function () { MW_GAUTH.prepare(m.google.clientId, m.google.email).then(function () { gBtn.style.display = ""; }, function (e) { err.textContent = e.message; }); };
			document.head.appendChild(s);
			gBtn.onclick = function () {
				var tok;
				err.textContent = "Google 확인 중…";
				MW_GAUTH.token().then(function (t) {
					tok = t;
					return MW_GAUTH.email(t);
				}).then(function (email) {
					if (email !== m.google.email.toLowerCase()) throw new Error(email + " 계정은 열 수 없습니다. " + m.google.email + " 으로 로그인해 주세요.");
					return MW_GAUTH.readKey(tok);
				}).then(function (raw) {
					if (!raw) throw new Error("Google Drive 에 보기 키가 없습니다. PC 의 MY Work 에 한 번 로그인하면 연결됩니다.");
					return useRaw(raw).catch(function () { throw new Error("보기 키가 예전 것입니다. PC 의 MY Work 에 다시 로그인하면 새 키로 맞춰집니다."); });
				}).then(function (raw) {
					keepStore().setItem(KEY, raw);
					location.reload();
				}).catch(function (e) { err.textContent = e.message; });
			};
		});
		d.querySelector("form").onsubmit = function (e) {
			e.preventDefault();
			var pw = document.getElementById("mkPw").value;
			err.textContent = "확인 중…";
			metaP.then(function (m) {
				return crypto.subtle.importKey("raw", new TextEncoder().encode(pw), "PBKDF2", false, ["deriveKey"]).then(function (base) {
					return crypto.subtle.deriveKey({ name: "PBKDF2", salt: unb64(m.salt), iterations: m.iter, hash: "SHA-256" }, base, { name: "AES-GCM", length: 256 }, true, ["decrypt"]);
				}).then(function (key) {
					return decrypt(key, unb64(m.check)).then(function () { return crypto.subtle.exportKey("raw", key); });
				});
			}).then(function (raw) {
				keepStore().setItem(KEY, b64(raw));
				location.reload();
			}).catch(function () { err.textContent = "비밀번호가 맞지 않습니다."; });
		};
	}
})();
