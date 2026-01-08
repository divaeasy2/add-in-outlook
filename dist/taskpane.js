/******/ (function() { // webpackBootstrap
/******/ 	var __webpack_modules__ = ({

/***/ "./src/taskpane/taskpane.css":
/*!***********************************!*\
  !*** ./src/taskpane/taskpane.css ***!
  \***********************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {

"use strict";
module.exports = __webpack_require__.p + "280a39d1350777bdbdd0.css";

/***/ })

/******/ 	});
/************************************************************************/
/******/ 	// The module cache
/******/ 	var __webpack_module_cache__ = {};
/******/ 	
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/ 		// Check if module is in cache
/******/ 		var cachedModule = __webpack_module_cache__[moduleId];
/******/ 		if (cachedModule !== undefined) {
/******/ 			return cachedModule.exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = __webpack_module_cache__[moduleId] = {
/******/ 			// no module.id needed
/******/ 			// no module.loaded needed
/******/ 			exports: {}
/******/ 		};
/******/ 	
/******/ 		// Execute the module function
/******/ 		__webpack_modules__[moduleId](module, module.exports, __webpack_require__);
/******/ 	
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/ 	
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = __webpack_modules__;
/******/ 	
/************************************************************************/
/******/ 	/* webpack/runtime/global */
/******/ 	!function() {
/******/ 		__webpack_require__.g = (function() {
/******/ 			if (typeof globalThis === 'object') return globalThis;
/******/ 			try {
/******/ 				return this || new Function('return this')();
/******/ 			} catch (e) {
/******/ 				if (typeof window === 'object') return window;
/******/ 			}
/******/ 		})();
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/hasOwnProperty shorthand */
/******/ 	!function() {
/******/ 		__webpack_require__.o = function(obj, prop) { return Object.prototype.hasOwnProperty.call(obj, prop); }
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/make namespace object */
/******/ 	!function() {
/******/ 		// define __esModule on exports
/******/ 		__webpack_require__.r = function(exports) {
/******/ 			if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 				Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 			}
/******/ 			Object.defineProperty(exports, '__esModule', { value: true });
/******/ 		};
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/publicPath */
/******/ 	!function() {
/******/ 		var scriptUrl;
/******/ 		if (__webpack_require__.g.importScripts) scriptUrl = __webpack_require__.g.location + "";
/******/ 		var document = __webpack_require__.g.document;
/******/ 		if (!scriptUrl && document) {
/******/ 			if (document.currentScript && document.currentScript.tagName.toUpperCase() === 'SCRIPT')
/******/ 				scriptUrl = document.currentScript.src;
/******/ 			if (!scriptUrl) {
/******/ 				var scripts = document.getElementsByTagName("script");
/******/ 				if(scripts.length) {
/******/ 					var i = scripts.length - 1;
/******/ 					while (i > -1 && (!scriptUrl || !/^http(s?):/.test(scriptUrl))) scriptUrl = scripts[i--].src;
/******/ 				}
/******/ 			}
/******/ 		}
/******/ 		// When supporting browsers where an automatic publicPath is not supported you must specify an output.publicPath manually via configuration
/******/ 		// or pass an empty string ("") and set the __webpack_public_path__ variable from your code to use your own logic.
/******/ 		if (!scriptUrl) throw new Error("Automatic publicPath is not supported in this browser");
/******/ 		scriptUrl = scriptUrl.replace(/^blob:/, "").replace(/#.*$/, "").replace(/\?.*$/, "").replace(/\/[^\/]+$/, "/");
/******/ 		__webpack_require__.p = scriptUrl;
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/jsonp chunk loading */
/******/ 	!function() {
/******/ 		__webpack_require__.b = (typeof document !== 'undefined' && document.baseURI) || self.location.href;
/******/ 		
/******/ 		// object to store loaded and loading chunks
/******/ 		// undefined = chunk not loaded, null = chunk preloaded/prefetched
/******/ 		// [resolve, reject, Promise] = chunk loading, 0 = chunk loaded
/******/ 		var installedChunks = {
/******/ 			"taskpane": 0
/******/ 		};
/******/ 		
/******/ 		// no chunk on demand loading
/******/ 		
/******/ 		// no prefetching
/******/ 		
/******/ 		// no preloaded
/******/ 		
/******/ 		// no HMR
/******/ 		
/******/ 		// no HMR manifest
/******/ 		
/******/ 		// no on chunks loaded
/******/ 		
/******/ 		// no jsonp function
/******/ 	}();
/******/ 	
/************************************************************************/
var __webpack_exports__ = {};
// This entry needs to be wrapped in an IIFE because it needs to be isolated against other entry modules.
!function() {
/*!**********************************!*\
  !*** ./src/taskpane/taskpane.js ***!
  \**********************************/
function _regenerator() { /*! regenerator-runtime -- Copyright (c) 2014-present, Facebook, Inc. -- license (MIT): https://github.com/babel/babel/blob/main/packages/babel-helpers/LICENSE */ var e, t, r = "function" == typeof Symbol ? Symbol : {}, n = r.iterator || "@@iterator", o = r.toStringTag || "@@toStringTag"; function i(r, n, o, i) { var c = n && n.prototype instanceof Generator ? n : Generator, u = Object.create(c.prototype); return _regeneratorDefine2(u, "_invoke", function (r, n, o) { var i, c, u, f = 0, p = o || [], y = !1, G = { p: 0, n: 0, v: e, a: d, f: d.bind(e, 4), d: function d(t, r) { return i = t, c = 0, u = e, G.n = r, a; } }; function d(r, n) { for (c = r, u = n, t = 0; !y && f && !o && t < p.length; t++) { var o, i = p[t], d = G.p, l = i[2]; r > 3 ? (o = l === n) && (u = i[(c = i[4]) ? 5 : (c = 3, 3)], i[4] = i[5] = e) : i[0] <= d && ((o = r < 2 && d < i[1]) ? (c = 0, G.v = n, G.n = i[1]) : d < l && (o = r < 3 || i[0] > n || n > l) && (i[4] = r, i[5] = n, G.n = l, c = 0)); } if (o || r > 1) return a; throw y = !0, n; } return function (o, p, l) { if (f > 1) throw TypeError("Generator is already running"); for (y && 1 === p && d(p, l), c = p, u = l; (t = c < 2 ? e : u) || !y;) { i || (c ? c < 3 ? (c > 1 && (G.n = -1), d(c, u)) : G.n = u : G.v = u); try { if (f = 2, i) { if (c || (o = "next"), t = i[o]) { if (!(t = t.call(i, u))) throw TypeError("iterator result is not an object"); if (!t.done) return t; u = t.value, c < 2 && (c = 0); } else 1 === c && (t = i.return) && t.call(i), c < 2 && (u = TypeError("The iterator does not provide a '" + o + "' method"), c = 1); i = e; } else if ((t = (y = G.n < 0) ? u : r.call(n, G)) !== a) break; } catch (t) { i = e, c = 1, u = t; } finally { f = 1; } } return { value: t, done: y }; }; }(r, o, i), !0), u; } var a = {}; function Generator() {} function GeneratorFunction() {} function GeneratorFunctionPrototype() {} t = Object.getPrototypeOf; var c = [][n] ? t(t([][n]())) : (_regeneratorDefine2(t = {}, n, function () { return this; }), t), u = GeneratorFunctionPrototype.prototype = Generator.prototype = Object.create(c); function f(e) { return Object.setPrototypeOf ? Object.setPrototypeOf(e, GeneratorFunctionPrototype) : (e.__proto__ = GeneratorFunctionPrototype, _regeneratorDefine2(e, o, "GeneratorFunction")), e.prototype = Object.create(u), e; } return GeneratorFunction.prototype = GeneratorFunctionPrototype, _regeneratorDefine2(u, "constructor", GeneratorFunctionPrototype), _regeneratorDefine2(GeneratorFunctionPrototype, "constructor", GeneratorFunction), GeneratorFunction.displayName = "GeneratorFunction", _regeneratorDefine2(GeneratorFunctionPrototype, o, "GeneratorFunction"), _regeneratorDefine2(u), _regeneratorDefine2(u, o, "Generator"), _regeneratorDefine2(u, n, function () { return this; }), _regeneratorDefine2(u, "toString", function () { return "[object Generator]"; }), (_regenerator = function _regenerator() { return { w: i, m: f }; })(); }
function _regeneratorDefine2(e, r, n, t) { var i = Object.defineProperty; try { i({}, "", {}); } catch (e) { i = 0; } _regeneratorDefine2 = function _regeneratorDefine(e, r, n, t) { function o(r, n) { _regeneratorDefine2(e, r, function (e) { return this._invoke(r, n, e); }); } r ? i ? i(e, r, { value: n, enumerable: !t, configurable: !t, writable: !t }) : e[r] = n : (o("next", 0), o("throw", 1), o("return", 2)); }, _regeneratorDefine2(e, r, n, t); }
function asyncGeneratorStep(n, t, e, r, o, a, c) { try { var i = n[a](c), u = i.value; } catch (n) { return void e(n); } i.done ? t(u) : Promise.resolve(u).then(r, o); }
function _asyncToGenerator(n) { return function () { var t = this, e = arguments; return new Promise(function (r, o) { var a = n.apply(t, e); function _next(n) { asyncGeneratorStep(a, r, o, _next, _throw, "next", n); } function _throw(n) { asyncGeneratorStep(a, r, o, _next, _throw, "throw", n); } _next(void 0); }); }; }
/* global Office */

var cachedPayload = null;
var MAX_EMAIL_SIZE = 500 * 1024; // 500 KB
var statusTimeoutId = null;
var allEvents = []; // Store all events for filtering

Office.onReady(function () {
  var displayName = Office.context.mailbox.userProfile.displayName;
  document.getElementById("user").innerText = displayName;
  document.getElementById("userEmail").innerText = Office.context.mailbox.userProfile.emailAddress;

  // Set avatar with user initials
  setAvatarInitials(displayName);

  // SAV event - show options modal
  document.getElementById("btnSav").onclick = function () {
    return showSavOptions();
  };

  // Options modal buttons
  document.getElementById("btnSavNew").onclick = function () {
    return sendSavNew();
  };
  document.getElementById("btnSavLinked").onclick = function () {
    return loadChildEventsForSav();
  };
  document.getElementById("btnSavCancel").onclick = function () {
    return hideSavOptions();
  };

  // Other event types
  document.getElementById("btnComm").onclick = function () {
    return send("2");
  };
  document.getElementById("btnDDP").onclick = function () {
    return send("3");
  };
  document.getElementById("btnCDE").onclick = function () {
    return send("4");
  };
  document.getElementById("btnDDI").onclick = function () {
    return send("5");
  };

  // Child events modal buttons
  document.getElementById("confirmEvt").onclick = function () {
    return confirmLinkedEvent();
  };
  document.getElementById("AnnuleEvt").onclick = function () {
    return returnFromLinkedEvents();
  };

  // Search input for events
  document.getElementById("eventSearchInput").addEventListener("input", function (e) {
    return filterEvents(e.target.value);
  });

  // Confirmation modal buttons
  document.getElementById("btnConfirmLinked").onclick = function () {
    return sendWithLinkedEvent();
  };
  document.getElementById("btnCancelLinked").onclick = function () {
    return returnFromConfirmation();
  };
  prepareEmail();
});

/* ======================
   STATUS
====================== */

function setAvatarInitials(displayName) {
  var avatar = document.getElementById("avatar");
  if (!displayName) {
    avatar.innerText = "?";
    return;
  }
  var names = displayName.trim().split(/\s+/);
  var initials = "";
  if (names.length >= 2) {
    // Get first letter of first name and last name
    initials = (names[0][0] + names[names.length - 1][0]).toUpperCase();
  } else if (names.length === 1) {
    // Single name: show first letter twice or just once
    initials = names[0][0].toUpperCase();
  }
  avatar.innerText = initials;
}
function showStatus(msg) {
  var type = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : "info";
  var el = document.getElementById("status");
  el.className = "status ".concat(type);
  el.innerText = msg;
  el.style.display = "block";

  // Clear previous timeout if exists
  if (statusTimeoutId) {
    clearTimeout(statusTimeoutId);
  }

  // Auto-hide status message after 3.5 seconds
  statusTimeoutId = setTimeout(function () {
    el.style.display = "none";
    statusTimeoutId = null;
  }, 3500);
}
function showChildHint() {
  var msg = arguments.length > 0 && arguments[0] !== undefined ? arguments[0] : "";
  var hint = document.getElementById("savHint");
  if (!hint) return;
  if (!msg) {
    hint.style.display = "none";
    hint.innerText = "";
  } else {
    hint.innerText = msg;
    hint.style.display = "block";
  }
}

/* ======================
   PREPARE EMAIL
====================== */

function prepareEmail() {
  var item = Office.context.mailbox.item;
  var user = Office.context.mailbox.userProfile.emailAddress;
  item.body.getAsync(Office.CoercionType.Text, function (res) {
    var _item$from;
    if (res.status !== Office.AsyncResultStatus.Succeeded) {
      showStatus("❌ Impossible de lire l'email", "error");
      return;
    }
    cachedPayload = {
      evenement: {
        type: "",
        utilisateur: user,
        tiers: ((_item$from = item.from) === null || _item$from === void 0 ? void 0 : _item$from.emailAddress) || "",
        lib: item.subject || "",
        pj: "",
        evt_lie: ""
      }
    };
    document.getElementById("btnSav").disabled = false;
    document.getElementById("btnComm").disabled = false;
    showStatus("✅ Email prêt pour traitement", "info");
  });
}

/* ======================
   BUILD EMAIL (.eml)
====================== */

function buildEmailBase64(item, bodyText) {
  var _item$from2;
  var bodyBase64 = btoa(unescape(encodeURIComponent(bodyText)));
  var eml = "From: ".concat(((_item$from2 = item.from) === null || _item$from2 === void 0 ? void 0 : _item$from2.emailAddress) || "", "\nTo: ").concat(Office.context.mailbox.userProfile.emailAddress, "\nSubject: ").concat(item.subject || "", "\nDate: ").concat(new Date().toUTCString(), "\nMIME-Version: 1.0\nContent-Type: text/plain; charset=UTF-8\nContent-Transfer-Encoding: base64\n\n").concat(bodyBase64);
  var size = new Blob([eml]).size;
  if (size > MAX_EMAIL_SIZE) return null;
  return btoa(unescape(encodeURIComponent(eml)));
}

/* ======================
   PARSER API
====================== */

function parseWeirdApiResponse(raw) {
  var n1;
  try {
    n1 = JSON.parse(raw);
  } catch (_unused) {
    return {
      ok: false,
      error: "N1 n'est pas JSON",
      raw: raw
    };
  }
  var n2 = n1.raw || n1.response || raw;
  var cleaned = n2.replace(/\\"/g, '"').replace(/"{/g, '{').replace(/}"/g, '}').replace(/""result":/g, '"result":').replace(/"result":"result":/g, '"result":').replace(/"result":""/g, '"result":').replace(/"result":\s*"({)/g, '"result":$1').trim();
  debugLog("🔧 Nettoyé:\n" + cleaned);
  var n3;
  try {
    n3 = JSON.parse(cleaned);
  } catch (_unused2) {
    return {
      ok: false,
      error: "❌ Impossible de parser N2 → JSON",
      cleaned: cleaned
    };
  }
  var events = n3.Evenements || n3.evenements || n3.response && n3.response.Evenements;
  if (!events) return {
    ok: false,
    error: "❌ Aucun évènement trouvé",
    json: n3
  };
  return {
    ok: true,
    count: events.length,
    events: events
  };
}
function debugLog(msg) {
  var box = document.getElementById("debug");
  // box.style.display = "block";
  box.innerText += "\n" + msg;
}

/* ======================
   SAV EVENT WORKFLOW
====================== */

function showSavOptions() {
  document.getElementById("savOptionsModal").style.display = "block";
}
function hideSavOptions() {
  document.getElementById("savOptionsModal").style.display = "none";
  showStatus("");
  document.getElementById("status").style.display = "none";
}
function sendSavNew() {
  return _sendSavNew.apply(this, arguments);
}
function _sendSavNew() {
  _sendSavNew = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee() {
    return _regenerator().w(function (_context) {
      while (1) switch (_context.n) {
        case 0:
          hideSavOptions();
          _context.n = 1;
          return send("1");
        case 1:
          return _context.a(2);
      }
    }, _callee);
  }));
  return _sendSavNew.apply(this, arguments);
}
function loadChildEventsForSav() {
  return _loadChildEventsForSav.apply(this, arguments);
}
function _loadChildEventsForSav() {
  _loadChildEventsForSav = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee2() {
    var payload, res, parsed, popup, select, evtCount, searchInput;
    return _regenerator().w(function (_context2) {
      while (1) switch (_context2.n) {
        case 0:
          document.getElementById("status").style.display = "block";
          if (cachedPayload) {
            _context2.n = 1;
            break;
          }
          return _context2.a(2, showStatus("⚠️ Aucun email prêt", "error"));
        case 1:
          showStatus("⏳ Vérification des évènements ...", "info");
          payload = {
            evenement: {
              utilisateur: cachedPayload.evenement.utilisateur,
              tiers: cachedPayload.evenement.tiers
            }
          };
          _context2.n = 2;
          return fetch("https://maisondelarose.org/proxy/proxy_child.php", {
            method: "POST",
            headers: {
              "Content-Type": "application/json; charset=UTF-8"
            },
            body: JSON.stringify(payload)
          }).then(function (r) {
            return r.text();
          }).catch(function () {
            return null;
          });
        case 2:
          res = _context2.v;
          if (res) {
            _context2.n = 3;
            break;
          }
          return _context2.a(2, showStatus("❌ Erreur réseau", "error"));
        case 3:
          parsed = parseWeirdApiResponse(res);
          popup = document.getElementById("childPopup");
          select = document.getElementById("childSelect");
          evtCount = document.getElementById("evtCount");
          searchInput = document.getElementById("eventSearchInput");
          if (parsed.ok) {
            _context2.n = 4;
            break;
          }
          popup.style.display = "none";
          return _context2.a(2, showStatus("🔴 " + parsed.error, "error"));
        case 4:
          // Store events for filtering
          allEvents = parsed.events;

          // Hide options modal and show child events popup
          document.getElementById("savOptionsModal").style.display = "none";
          popup.style.display = "block";

          // Clear and rebuild select with proper encoding
          select.innerHTML = "<option value=\"\">-- Choisissez un \xE9v\xE8nement --</option>";
          parsed.events.forEach(function (evt) {
            var opt = document.createElement("option");
            opt.value = evt.evtNo;
            // Ensure proper encoding of event text
            var eventCode = String(evt.evtNo || "");
            var eventLib = String(evt.lib || "(sans lib)");
            opt.textContent = "".concat(eventCode, " - ").concat(eventLib);
            select.appendChild(opt);
          });
          evtCount.innerText = "".concat(parsed.count, " \xE9v\xE8nements trouv\xE9s");

          // Show search input if events exist
          if (parsed.count > 0) {
            searchInput.style.display = "block";
            searchInput.value = "";
          } else {
            searchInput.style.display = "none";
          }
          showStatus("\uD83D\uDFE2 ".concat(parsed.count, " \xE9v\xE8nements r\xE9cup\xE9r\xE9s"), "success");
        case 5:
          return _context2.a(2);
      }
    }, _callee2);
  }));
  return _loadChildEventsForSav.apply(this, arguments);
}
function filterEvents(searchTerm) {
  var select = document.getElementById("childSelect");
  var searchValue = searchTerm.toLowerCase().trim();

  // Clear current options (except placeholder)
  select.innerHTML = "<option value=\"\">-- Choisissez un \xE9v\xE8nement --</option>";

  // Filter events
  var filteredEvents = allEvents.filter(function (evt) {
    var eventCode = String(evt.evtNo || "").toLowerCase();
    var eventLib = String(evt.lib || "").toLowerCase();
    return eventCode.includes(searchValue) || eventLib.includes(searchValue);
  });

  // Add filtered events to select
  filteredEvents.forEach(function (evt) {
    var opt = document.createElement("option");
    opt.value = evt.evtNo;
    var eventCode = String(evt.evtNo || "");
    var eventLib = String(evt.lib || "(sans lib)");
    opt.textContent = "".concat(eventCode, " - ").concat(eventLib);
    select.appendChild(opt);
  });

  // Update count
  var evtCount = document.getElementById("evtCount");
  if (searchValue) {
    evtCount.innerText = "".concat(filteredEvents.length, " \xE9v\xE8nements trouv\xE9s");
  } else {
    evtCount.innerText = "".concat(allEvents.length, " \xE9v\xE8nements trouv\xE9s");
  }
}
function confirmLinkedEvent() {
  var select = document.getElementById("childSelect");
  var chosen = select.value;
  if (!chosen) {
    return showStatus("⚠️ Sélectionnez un évènement", "error");
  }

  // Store the selected event
  cachedPayload.evenement.evt_lie = chosen;

  // Get the event details for display
  var selectedOption = select.options[select.selectedIndex];
  var eventText = selectedOption.textContent;

  // Hide child events popup
  document.getElementById("childPopup").style.display = "none";

  // Show confirmation modal
  document.getElementById("confirmLinkedText").innerText = "Confirmer l'\xE9v\xE8nement:\n".concat(eventText);
  document.getElementById("confirmLinkedModal").style.display = "block";
}
function sendWithLinkedEvent() {
  return _sendWithLinkedEvent.apply(this, arguments);
}
function _sendWithLinkedEvent() {
  _sendWithLinkedEvent = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee3() {
    return _regenerator().w(function (_context3) {
      while (1) switch (_context3.n) {
        case 0:
          document.getElementById("confirmLinkedModal").style.display = "none";
          _context3.n = 1;
          return send("1");
        case 1:
          return _context3.a(2);
      }
    }, _callee3);
  }));
  return _sendWithLinkedEvent.apply(this, arguments);
}
function returnFromConfirmation() {
  document.getElementById("confirmLinkedModal").style.display = "none";
  cachedPayload.evenement.evt_lie = "";

  // Show child events popup again
  document.getElementById("childPopup").style.display = "block";
  showStatus("⏳ Retour à la sélection...", "info");
}
function returnFromLinkedEvents() {
  document.getElementById("childPopup").style.display = "none";
  cachedPayload.evenement.evt_lie = "";

  // Show options modal again
  document.getElementById("savOptionsModal").style.display = "block";
  showStatus("");
  document.getElementById("status").style.display = "none";
}

/* ======================
   SEND
====================== */
function send(_x) {
  return _send.apply(this, arguments);
}
function _send() {
  _send = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee4(type) {
    var _parsed$json, _resultStr$match, _resultStr$match2, item, body, emailBase64, res, text, parsed, resultStr, code, evt, _t;
    return _regenerator().w(function (_context4) {
      while (1) switch (_context4.p = _context4.n) {
        case 0:
          document.getElementById("status").style.display = "block";
          if (cachedPayload) {
            _context4.n = 1;
            break;
          }
          return _context4.a(2, showStatus("⚠️ Aucun email prêt", "error"));
        case 1:
          _context4.p = 1;
          item = Office.context.mailbox.item;
          showStatus("⌛ Lecture...", "info");
          _context4.n = 2;
          return new Promise(function (resolve, reject) {
            item.body.getAsync(Office.CoercionType.Text, function (r) {
              r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value) : reject();
            });
          });
        case 2:
          body = _context4.v;
          cachedPayload.evenement.type = type;
          cachedPayload.evenement.evt_lie = cachedPayload.evenement.evt_lie || "";
          emailBase64 = buildEmailBase64(item, body);
          cachedPayload.evenement.pj = emailBase64 || "";
          showStatus("🚀 Envoi...", "info");
          _context4.n = 3;
          return fetch("https://maisondelarose.org/proxy/proxy.php", {
            method: "POST",
            headers: {
              "Content-Type": "application/json"
            },
            body: JSON.stringify(cachedPayload)
          });
        case 3:
          res = _context4.v;
          _context4.n = 4;
          return res.text();
        case 4:
          text = _context4.v;
          parsed = JSON.parse(text);
          resultStr = (parsed === null || parsed === void 0 || (_parsed$json = parsed.json) === null || _parsed$json === void 0 ? void 0 : _parsed$json.result) || "";
          code = (_resultStr$match = resultStr.match(/"resultcode"\s*:\s*"(\d+)"/)) === null || _resultStr$match === void 0 ? void 0 : _resultStr$match[1];
          evt = (_resultStr$match2 = resultStr.match(/"EvtNo"\s*:\s*"([^"]+)"/)) === null || _resultStr$match2 === void 0 || (_resultStr$match2 = _resultStr$match2[1]) === null || _resultStr$match2 === void 0 ? void 0 : _resultStr$match2.trim();
          if (code === "0") {
            showStatus("\uD83C\uDF89 SUCCESS \u2014 Code ".concat(evt), "success");
          } else {
            showStatus("\u274C Erreur", "error");
          }
          _context4.n = 6;
          break;
        case 5:
          _context4.p = 5;
          _t = _context4.v;
          showStatus("❌ Erreur de communication", "error");
        case 6:
          return _context4.a(2);
      }
    }, _callee4, null, [[1, 5]]);
  }));
  return _send.apply(this, arguments);
}
}();
// This entry needs to be wrapped in an IIFE because it needs to be in strict mode.
!function() {
"use strict";
/*!************************************!*\
  !*** ./src/taskpane/taskpane.html ***!
  \************************************/
__webpack_require__.r(__webpack_exports__);
// Imports
var ___HTML_LOADER_IMPORT_0___ = new URL(/* asset import */ __webpack_require__(/*! ./taskpane.css */ "./src/taskpane/taskpane.css"), __webpack_require__.b);
// Module
var code = "<!-- Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT License. -->\r\n<!-- This file shows how to design a first-run page that provides a welcome screen to the user about the features of the add-in. -->\r\n\r\n<!DOCTYPE html>\r\n<html>\r\n\r\n<head>\r\n    <meta charset=\"UTF-8\" />\r\n    <meta http-equiv=\"X-UA-Compatible\" content=\"IE=Edge\" />\r\n    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\r\n    <title>Divalto Task Pane Add-in</title>\r\n\r\n    <!-- Office JavaScript API -->\r\n    <" + "script type=\"text/javascript\" src=\"https://appsforoffice.microsoft.com/lib/1/hosted/office.js\"><" + "/script>\r\n\r\n    <!-- For more information on Fluent UI, visit https://developer.microsoft.com/fluentui#/. -->\r\n    <link rel=\"stylesheet\" href=\"https://res-1.cdn.office.net/files/fabric-cdn-prod_20230815.002/office-ui-fabric-core/11.1.0/css/fabric.min.css\"/>\r\n\r\n    <!-- Template styles -->\r\n    <link href=\"" + ___HTML_LOADER_IMPORT_0___ + "\" rel=\"stylesheet\" type=\"text/css\" />\r\n</head>\r\n\r\n<body class=\"divalto-app\">\r\n\r\n  <!-- PROFILE -->\r\n  <section class=\"profile-card\">\r\n    <div class=\"avatar\" id=\"avatar\"></div>\r\n    <div>\r\n      <div class=\"username\" id=\"user\"></div>\r\n      <div class=\"email\" id=\"userEmail\"></div>\r\n    </div>\r\n  </section>\r\n\r\n  <!-- ACTIONS -->\r\n  <main class=\"actions-card\">\r\n    <button class=\"action-btn secondary\" id=\"btnSav\">\r\n      Évènement SAV\r\n    </button>\r\n\r\n    <button class=\"action-btn secondary\" id=\"btnComm\">\r\n      Évènement Négoce\r\n    </button>\r\n\r\n    <button class=\"action-btn secondary\" id=\"btnDDP\">\r\n      Entrée de demande de prix\r\n    </button>\r\n\r\n    <button class=\"action-btn secondary\" id=\"btnCDE\">\r\n      Entrée de commande\r\n    </button>\r\n\r\n    <button class=\"action-btn secondary\" id=\"btnDDI\">\r\n      Demande d'information\r\n    </button>\r\n\r\n    <!-- SAV Options Modal -->\r\n    <div id=\"savOptionsModal\" class=\"popup\" style=\"display:none;\">\r\n      <div class=\"popup-header\">\r\n        Type d'évènement SAV\r\n      </div>\r\n      <div class=\"btnPopup\" style=\"flex-direction: column; gap: 10px;\">\r\n        <button id=\"btnSavNew\" class=\"action-btn secondary\" style=\"margin-top:10px;\">\r\n          Nouveau\r\n        </button>\r\n        <button id=\"btnSavLinked\" class=\"action-btn secondary\" style=\"margin-top:10px;\">\r\n          Évènement lié\r\n        </button>\r\n        <button id=\"btnSavCancel\" class=\"action-btn secondary\" style=\"margin-top:10px; background: #999;\">\r\n          Retour\r\n        </button>\r\n      </div>\r\n    </div>\r\n\r\n    <!-- Child Events Popup -->\r\n    <div id=\"childPopup\" class=\"popup\" style=\"display:none;\">\r\n      <div class=\"popup-header\">\r\n        <span id=\"evtCount\">0 évènements trouvés</span>\r\n      </div>\r\n      \r\n      <input \r\n        type=\"text\" \r\n        id=\"eventSearchInput\" \r\n        placeholder=\"Rechercher par code ou nom...\" \r\n        class=\"event-search-input\"\r\n        style=\"display:none;\"\r\n      />\r\n\r\n      <select id=\"childSelect\" class=\"event-select\"></select>\r\n      <div class=\"btnPopup\">\r\n        <button id=\"confirmEvt\" class=\"action-btn primary\" style=\"margin-top:10px;\">\r\n          Confirmer\r\n        </button>\r\n        <button id=\"AnnuleEvt\" class=\"action-btn secondary\" style=\"margin-top:10px;\">\r\n          Retour\r\n        </button>\r\n      </div>\r\n    </div>\r\n\r\n    <!-- Confirmation Modal for Linked Event -->\r\n    <div id=\"confirmLinkedModal\" class=\"popup\" style=\"display:none;\">\r\n      <div class=\"popup-header\">\r\n        Confirmer l'évènement lié\r\n      </div>\r\n      <div id=\"confirmLinkedText\" style=\"padding: 12px; text-align: center; color: #333; margin-bottom: 10px;\">\r\n      </div>\r\n      <div class=\"btnPopup\">\r\n        <button id=\"btnConfirmLinked\" class=\"action-btn primary\" style=\"margin-top:10px;\">\r\n          Confirmer\r\n        </button>\r\n        <button id=\"btnCancelLinked\" class=\"action-btn secondary\" style=\"margin-top:10px;\">\r\n          Retour\r\n        </button>\r\n      </div>\r\n    </div>\r\n\r\n\r\n\r\n    <div id=\"status\" class=\"status\"></div>\r\n    <!-- <div id=\"pj\" ></div> -->\r\n    <div id=\"debug\" style=\"\r\n    background:#111;\r\n    color:#0f0;\r\n    padding:10px;\r\n    margin-top:20px;\r\n    height:200px;\r\n    overflow:auto;\r\n    font-size:12px;\r\n    border-radius:6px;\r\n    display:none;\r\n\"></div>\r\n\r\n  </main>\r\n\r\n</body>\r\n\r\n\r\n</html>\r\n";
// Exports
/* harmony default export */ __webpack_exports__["default"] = (code);
}();
/******/ })()
;
//# sourceMappingURL=taskpane.js.map