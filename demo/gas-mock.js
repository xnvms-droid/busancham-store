/**
 * 체험판(demo) 전용 — Google Apps Script 환경을 브라우저에서 흉내 내는 모의 객체
 * index.html?demo=1 로 접속하면 apps-script/Code.gs 를 이 모의 환경에서 그대로 실행합니다.
 * 데이터는 해당 브라우저의 localStorage 에만 저장되며 실제 구글 시트와는 무관합니다.
 */
(function (global) {
  'use strict';
  var KEY = 'hujeom-demo-db-v1';
  var db = null;

  function reviver(k, v) {
    if (v && typeof v === 'object' && typeof v.__d === 'number') return new Date(v.__d);
    return v;
  }
  function replacer(k, v) {
    var raw = this[k];
    if (raw instanceof Date) return { __d: raw.getTime() };
    return v;
  }
  function load() {
    db = null;
    try { db = JSON.parse(localStorage.getItem(KEY), reviver); } catch (e) {}
    if (!db || !db.sheets) db = { sheets: {}, order: [], props: {} };
  }
  function save() {
    try { localStorage.setItem(KEY, JSON.stringify(db, replacer)); } catch (e) {}
  }

  // 정의되지 않은 서식 관련 메서드(setFontWeight 등)는 모두 무시하고 자기 자신을 반환
  function chain(target) {
    return new Proxy(target, {
      get: function (t, prop) {
        if (prop in t) return t[prop];
        return function () { return chain(t); };
      }
    });
  }

  function Range(sheet, row, col, nr, nc) {
    var rows = function () { return db.sheets[sheet.name]; };
    return chain({
      getValues: function () {
        var out = [];
        for (var i = 0; i < nr; i++) {
          var r = rows()[row - 1 + i] || [];
          var line = [];
          for (var j = 0; j < nc; j++) {
            var v = r[col - 1 + j];
            line.push(v === undefined || v === null ? '' : v);
          }
          out.push(line);
        }
        return out;
      },
      getValue: function () { return this.getValues()[0][0]; },
      setValues: function (vals) {
        var data = rows();
        for (var i = 0; i < vals.length; i++) {
          var ri = row - 1 + i;
          while (data.length <= ri) data.push([]);
          for (var j = 0; j < vals[i].length; j++) data[ri][col - 1 + j] = vals[i][j];
        }
        return this;
      },
      setValue: function (v) { return this.setValues([[v]]); }
    });
  }

  function Sheet(name) {
    var self = chain({
      name: name,
      getName: function () { return name; },
      getLastRow: function () {
        var data = db.sheets[name];
        var n = data.length;
        while (n > 0 && (!data[n - 1] || data[n - 1].every(function (v) { return v === '' || v == null; }))) n--;
        return n;
      },
      getRange: function (a, b, c, d) {
        if (typeof a === 'string') return chain({});
        return Range(self, a, b, c || 1, d || 1);
      },
      appendRow: function (arr) { db.sheets[name].push(arr.slice()); return self; },
      deleteRows: function (start, n) { db.sheets[name].splice(start - 1, n); return self; }
    });
    return self;
  }

  var spreadsheet = chain({
    getSheetByName: function (name) { return db.sheets[name] ? Sheet(name) : null; },
    insertSheet: function (name) { db.sheets[name] = []; db.order.push(name); return Sheet(name); },
    getSheets: function () { return db.order.map(function (n) { return Sheet(n); }); },
    deleteSheet: function (s) { delete db.sheets[s.getName()]; db.order = db.order.filter(function (n) { return n !== s.getName(); }); }
  });

  global.SpreadsheetApp = {
    getActiveSpreadsheet: function () { return spreadsheet; },
    flush: function () {},
    newDataValidation: function () { return chain({}); },
    getUi: function () { throw new Error('체험판에는 시트 UI가 없습니다.'); }
  };
  global.LockService = { getScriptLock: function () { return { waitLock: function () {}, tryLock: function () { return true; }, releaseLock: function () {} }; } };
  global.CacheService = { getScriptCache: function () { return { get: function () { return null; }, put: function () {}, remove: function () {} }; } };
  global.PropertiesService = {
    getScriptProperties: function () {
      return {
        getProperty: function (k) { return db.props[k] == null ? null : db.props[k]; },
        setProperty: function (k, v) { db.props[k] = String(v); }
      };
    }
  };
  global.ContentService = {
    MimeType: { JSON: 'application/json' },
    createTextOutput: function (s) {
      return { setMimeType: function () { return this; }, getContent: function () { return s; } };
    }
  };
  global.Utilities = {
    getUuid: function () {
      if (global.crypto && crypto.randomUUID) return crypto.randomUUID();
      return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
        var r = Math.random() * 16 | 0; return (c === 'x' ? r : (r & 0x3 | 0x8)).toString(16);
      });
    }
  };

  global.DemoBackend = {
    call: function (params, post) {
      load();
      if (!db.sheets['설정']) {
        global.setup();
        // 체험판: 비밀번호 없이 접속, 행사명 표시
        db.sheets['설정'][1][1] = '후원주점 (체험판)';
        db.sheets['설정'][3][1] = '';
      }
      var out = post
        ? global.doPost({ postData: { contents: JSON.stringify(params) } })
        : global.doGet({ parameter: params });
      save();
      var res = JSON.parse(out.getContent());
      return new Promise(function (resolve) { setTimeout(function () { resolve(res); }, 120); });
    },
    reset: function () { try { localStorage.removeItem(KEY); } catch (e) {} }
  };
})(window);
