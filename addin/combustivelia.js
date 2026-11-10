/* Combustível IA · Add-In de página para o MyGeotab (versão 0.2)
 * Motor ligado parado por veículo e, dentro de cada veículo, os 3 motoristas
 * que mais contribuíram. Somente leitura. Respeita o filtro de grupos do usuário.
 */
(function () {
  'use strict';

  var DAYS = 7;
  var TOP_VEHICLES = 10;
  var TOP_DRIVERS = 3;
  var UNKNOWN_DRIVER = 'UnknownDriverId';

  // ---- lógica pura (testada em tests/addin.test.js) ------------------------
  var DURATION = /^(?:(\d+)\.)?(\d{1,2}):(\d{2}):(\d{2})(?:\.\d+)?$/;

  function parseDuration(value) {
    if (value === null || value === undefined) return 0;
    if (typeof value === 'number') return value;
    var m = DURATION.exec(String(value).trim());
    if (!m) return 0;
    return (Number(m[1] || 0) * 86400) + Number(m[2]) * 3600 + Number(m[3]) * 60 + Number(m[4]);
  }

  function refId(ref) {
    if (!ref) return null;
    return typeof ref === 'object' ? ref.id : String(ref);
  }

  function driverNames(users) {
    var names = {};
    (users || []).forEach(function (u) {
      var full = [u.firstName, u.lastName].filter(Boolean).join(' ');
      names[u.id] = full || u.name || u.id;
    });
    return names;
  }

  /** Por veículo: tempo parado, % do tempo de motor ligado e os motoristas que
   *  mais contribuíram, cada um como parte do tempo parado DAQUELE veículo. */
  function summarizeIdle(devices, trips, users) {
    var names = driverNames(users);
    var byId = {};
    devices.forEach(function (d) {
      byId[d.id] = { id: d.id, name: d.name || d.id, idleS: 0, drivingS: 0, unknownS: 0, byDriver: {} };
    });
    trips.forEach(function (t) {
      var row = byId[refId(t.device)];
      if (!row) return; // veículo fora do filtro de grupos
      var idle = parseDuration(t.idlingDuration);
      row.idleS += idle;
      row.drivingS += parseDuration(t.drivingDuration);
      if (!idle) return;
      var id = refId(t.driver);
      if (!id || id === UNKNOWN_DRIVER) { row.unknownS += idle; return; }
      var d = row.byDriver[id] || (row.byDriver[id] = { id: id, name: names[id] || 'Motorista ' + id, idleS: 0, trips: 0 });
      d.idleS += idle;
      d.trips += 1;
    });

    var totalIdle = 0, totalEngine = 0;
    var rows = Object.keys(byId).map(function (k) {
      var r = byId[k];
      var engine = r.idleS + r.drivingS;
      totalIdle += r.idleS;
      totalEngine += engine;
      var drivers = Object.keys(r.byDriver).map(function (id) { return r.byDriver[id]; });
      drivers.forEach(function (d) { d.share = r.idleS > 0 ? d.idleS / r.idleS : 0; });
      drivers.sort(function (a, b) { return b.idleS - a.idleS; });
      return {
        id: r.id,
        name: r.name,
        idleS: r.idleS,
        share: engine > 0 ? r.idleS / engine : 0,
        drivers: drivers.slice(0, TOP_DRIVERS),
        unknownShare: r.idleS > 0 ? r.unknownS / r.idleS : 0
      };
    });
    rows.sort(function (a, b) { return b.idleS - a.idleS; });
    return {
      vehicles: rows.length,
      totalIdleS: totalIdle,
      share: totalEngine > 0 ? totalIdle / totalEngine : 0,
      ranking: rows.filter(function (r) { return r.idleS > 0; })
    };
  }

  // ---- formatação ------------------------------------------------------------
  var hoursFmt = new Intl.NumberFormat('pt-BR', { maximumFractionDigits: 1, minimumFractionDigits: 1 });
  var pctFmt = new Intl.NumberFormat('pt-BR', { style: 'percent', maximumFractionDigits: 0 });
  function hours(s) { return hoursFmt.format(s / 3600) + ' h'; }

  // ---- interface -------------------------------------------------------------
  function $(id) { return document.getElementById(id); }
  function setStatus(text) { $('cia-status').textContent = text || ''; }

  // Todo texto vindo da base do cliente entra como textContent, nunca como HTML
  function el(tag, className, text) {
    var node = document.createElement(tag);
    if (className) node.className = className;
    if (text !== undefined) node.textContent = text;
    return node;
  }

  function vehicleItem(v) {
    var li = el('li', 'cia-vehicle');
    var head = el('div', 'cia-vehicle-head');
    head.appendChild(el('span', 'cia-vehicle-name', v.name));
    head.appendChild(el('span', 'cia-vehicle-hours', hours(v.idleS)));
    li.appendChild(head);
    li.appendChild(el('p', 'cia-vehicle-share', pctFmt.format(v.share) + ' do tempo de motor ligado'));

    if (v.drivers.length) {
      var list = el('ol', 'cia-drivers');
      v.drivers.forEach(function (d, i) {
        var item = el('li', 'cia-driver');
        item.appendChild(el('span', 'cia-driver-pos', String(i + 1)));
        item.appendChild(el('span', 'cia-driver-name', d.name));
        item.appendChild(el('span', 'cia-driver-detail', hours(d.idleS) + ' · ' + d.trips + (d.trips === 1 ? ' viagem' : ' viagens')));
        item.appendChild(el('span', 'cia-driver-share', pctFmt.format(d.share)));
        list.appendChild(item);
      });
      li.appendChild(list);
    }
    if (v.unknownShare > 0) {
      li.appendChild(el('p', 'cia-vehicle-unknown', pctFmt.format(v.unknownShare) + ' sem motorista identificado'));
    }
    return li;
  }

  function render(summary, start, end) {
    var dateFmt = new Intl.DateTimeFormat('pt-BR', { day: '2-digit', month: '2-digit' });
    $('cia-period').textContent = 'De ' + dateFmt.format(start) + ' a ' + dateFmt.format(end);
    $('cia-kpi-hours').textContent = hours(summary.totalIdleS);
    $('cia-kpi-share').textContent = pctFmt.format(summary.share);
    $('cia-kpi-vehicles').textContent = String(summary.vehicles);
    $('cia-kpis').hidden = false;

    var list = $('cia-ranking');
    while (list.firstChild) list.removeChild(list.firstChild);
    summary.ranking.slice(0, TOP_VEHICLES).forEach(function (v) { list.appendChild(vehicleItem(v)); });
    $('cia-ranking-card').hidden = summary.ranking.length === 0;
    setStatus(summary.ranking.length === 0 ? 'Nenhum tempo de motor ligado parado no período.' : '');
  }

  function call(api, method, params) {
    return new Promise(function (resolve, reject) { api.call(method, params, resolve, reject); });
  }

  function load(api, state) {
    var end = new Date();
    var start = new Date(end.getTime() - DAYS * 86400000);
    var groups = state && state.getGroupFilter ? state.getGroupFilter() : [];
    setStatus('Carregando dados da frota…');
    var deviceSearch = { fromDate: start.toISOString() };
    if (groups && groups.length) deviceSearch.groups = groups;

    return Promise.all([
      call(api, 'Get', { typeName: 'Device', search: deviceSearch }),
      call(api, 'Get', { typeName: 'Trip', search: { fromDate: start.toISOString(), toDate: end.toISOString() } }),
      call(api, 'Get', { typeName: 'User', search: { isDriver: true } })
    ]).then(function (res) {
      render(summarizeIdle(res[0], res[1], res[2]), start, end);
    }).catch(function (err) {
      setStatus('Não foi possível carregar os dados. ' + ((err && err.message) || ''));
    });
  }

  // ---- registro no MyGeotab --------------------------------------------------
  if (typeof window !== 'undefined') {
    window.geotab = window.geotab || { addin: {} };
    window.geotab.addin = window.geotab.addin || {};
    window.geotab.addin.combustivelia = function () {
      return {
        initialize: function (api, state, callback) { callback(); },
        focus: function (api, state) { load(api, state); },  // roda também ao mudar o filtro de grupos
        blur: function () {}
      };
    };
  }

  if (typeof module !== 'undefined' && module.exports) {
    module.exports = { parseDuration: parseDuration, summarizeIdle: summarizeIdle };
  }
})();
