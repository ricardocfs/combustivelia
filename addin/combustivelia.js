/* Combustível IA · Add-In de página para o MyGeotab (versão 0.3)
 * Motor ligado parado por veículo, em horas, litros e R$, com os 3 motoristas
 * que mais contribuíram em cada veículo. Check-up de dados: quais frentes de
 * economia esta frota permite analisar. Somente leitura. Respeita o filtro de grupos.
 */
(function () {
  'use strict';

  var VERSION = '0.3.0';
  var DAYS = 7;
  var TOP_VEHICLES = 10;
  var TOP_DRIVERS = 3;
  var UNKNOWN_DRIVER = 'UnknownDriverId';

  // Preço padrão: média nacional de revenda do diesel S10, ANP, semana de 18 a 24/10/2026
  var DEFAULT_PRICE = 7.32;
  var PRICE_SOURCE = 'média ANP do diesel S10, 18 a 24/10/2026';
  // Estimativa quando o veículo não informa o combustível gasto parado:
  // ~0,8 galão (≈3 L) por hora para caminhão pesado (Departamento de Energia dos EUA, AFDC)
  var IDLE_RATE_LPH = 3;

  var DIAG = {
    idleFuel: 'DiagnosticDeviceTotalIdleFuelId',
    totalFuel: 'DiagnosticDeviceTotalFuelId',
    rpm: 'DiagnosticEngineSpeedId',
    coolant: 'DiagnosticEngineCoolantTemperatureId'
  };

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

  /** Litros gastos parado no período, por veículo, a partir do contador acumulado
   *  da Geotab (maior leitura menos menor leitura). Só entra quem tem consumo > 0. */
  function idleFuelByDevice(records) {
    var range = {};
    (records || []).forEach(function (r) {
      var id = refId(r.device);
      var v = Number(r.data);
      if (!id || !isFinite(v)) return;
      var x = range[id] || (range[id] = { min: v, max: v });
      if (v < x.min) x.min = v;
      if (v > x.max) x.max = v;
    });
    var out = {};
    Object.keys(range).forEach(function (id) {
      var liters = range[id].max - range[id].min;
      if (liters > 0) out[id] = liters;
    });
    return out;
  }

  /** Converte o tempo parado em litros e R$. Usa o litro medido pela Geotab quando
   *  existe; senão estima pelas horas paradas. O custo de cada motorista é a parte
   *  dele no tempo parado daquele veículo. Reordena pelo custo. */
  function valueIdle(summary, fuelByDevice, price, rateLph) {
    var fuel = fuelByDevice || {};
    var rate = rateLph === undefined ? IDLE_RATE_LPH : rateLph;
    var totalLiters = 0, totalCost = 0, measured = 0;
    var ranking = summary.ranking.map(function (v) {
      var hasMeasure = fuel[v.id] > 0;
      var liters = hasMeasure ? fuel[v.id] : (v.idleS / 3600) * rate;
      var cost = liters * price;
      if (hasMeasure) measured += 1;
      totalLiters += liters;
      totalCost += cost;
      var drivers = v.drivers.map(function (d) {
        var copy = {};
        Object.keys(d).forEach(function (k) { copy[k] = d[k]; });
        copy.cost = cost * d.share;
        return copy;
      });
      var row = {};
      Object.keys(v).forEach(function (k) { row[k] = v[k]; });
      row.liters = liters;
      row.estimated = !hasMeasure;
      row.cost = cost;
      row.drivers = drivers;
      return row;
    });
    ranking.sort(function (a, b) { return b.cost - a.cost; });
    var out = {};
    Object.keys(summary).forEach(function (k) { out[k] = summary[k]; });
    out.ranking = ranking;
    out.totalLiters = totalLiters;
    out.totalCost = totalCost;
    out.measuredVehicles = measured;
    out.price = price;
    return out;
  }

  /** Check-up de dados: o que esta frota permite analisar em cada frente de economia. */
  function assessData(c) {
    var trips = c.trips || [];
    var devices = c.devices || [];
    var hasTrips = trips.length > 0;
    var hasSpeed = trips.some(function (t) { return t.maximumSpeed !== undefined && t.maximumSpeed !== null; });
    var hasDistance = trips.some(function (t) { return Number(t.distance) > 0; });
    var withVin = devices.filter(function (d) {
      var vin = String(d.vehicleIdentificationNumber || '').trim();
      return vin.length >= 11 && vin !== '?';
    }).length;
    var n = devices.length;
    var measuredIdle = c.idleFuelVehicles || 0;

    function item(label, ok, detail) { return { label: label, ok: !!ok, detail: detail || '' }; }
    function status(core, extras) {
      if (!core) return 'none';
      return extras.every(Boolean) ? 'ok' : 'partial';
    }

    var conducao = [
      item('Tempo de motor ligado parado', hasTrips),
      item('Combustível gasto parado, medido pelo veículo', measuredIdle > 0, measuredIdle + ' de ' + n + ' veículos'),
      item('Rotação do motor', c.rpm),
      item('Velocidade nas viagens', hasSpeed)
    ];
    var mecanica = [
      item('Consumo total de combustível', c.totalFuel),
      item('Distância percorrida', hasDistance),
      item('Temperatura do motor', c.coolant),
      item('Códigos de falha do motor', c.faults),
      item('Chassi (VIN) para comparar marca, modelo e ano', withVin > 0, withVin + ' de ' + n + ' veículos')
    ];
    var operacao = [
      item('Locais das paradas', hasTrips),
      item('Áreas cadastradas (clientes, docas, pátios)', c.zones)
    ];

    var pillars = [
      { key: 'conducao', name: 'Condução', question: 'O jeito de dirigir está gastando mais?', status: status(hasTrips, [c.rpm, hasSpeed]), items: conducao },
      { key: 'mecanica', name: 'Mecânica', question: 'Algum caminhão consome acima do normal?', status: status(c.totalFuel && hasDistance, [c.coolant, c.faults, withVin > 0]), items: mecanica },
      { key: 'operacao', name: 'Operação', question: 'O uso da frota gera desperdício?', status: status(hasTrips, [c.zones]), items: operacao }
    ];
    return { pillars: pillars, ready: pillars.filter(function (p) { return p.status !== 'none'; }).length };
  }

  // ---- formatação ------------------------------------------------------------
  var hoursFmt = new Intl.NumberFormat('pt-BR', { maximumFractionDigits: 1, minimumFractionDigits: 1 });
  var pctFmt = new Intl.NumberFormat('pt-BR', { style: 'percent', maximumFractionDigits: 0 });
  var brl = new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL', maximumFractionDigits: 0 });
  var litersFmt = new Intl.NumberFormat('pt-BR', { maximumFractionDigits: 0 });
  var priceFmt = new Intl.NumberFormat('pt-BR', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  function hours(s) { return hoursFmt.format(s / 3600) + ' h'; }
  function liters(l) { return litersFmt.format(l) + ' L'; }

  // ---- preço do diesel (por usuário, neste navegador) ---------------------------
  var PRICE_KEY = 'combustivelia.dieselPrice';
  function readPrice() {
    try {
      var v = Number(window.localStorage.getItem(PRICE_KEY));
      return v > 0 ? v : DEFAULT_PRICE;
    } catch (e) { return DEFAULT_PRICE; }
  }
  function savePrice(v) {
    try { window.localStorage.setItem(PRICE_KEY, String(v)); } catch (e) { /* sem armazenamento: vale só nesta visita */ }
  }
  function parsePrice(text) {
    var v = Number(String(text).replace(/\s|R\$/g, '').replace(',', '.'));
    return v > 0 && v < 100 ? v : null;
  }

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
  function clear(node) { while (node.firstChild) node.removeChild(node.firstChild); }

  function vehicleItem(v) {
    var li = el('li', 'cia-vehicle');
    var head = el('div', 'cia-vehicle-head');
    head.appendChild(el('span', 'cia-vehicle-name', v.name));
    head.appendChild(el('span', 'cia-vehicle-cost', brl.format(v.cost)));
    li.appendChild(head);
    var fuelText = (v.estimated ? '≈ ' + liters(v.liters) + ' estimados' : liters(v.liters));
    li.appendChild(el('p', 'cia-vehicle-share',
      hours(v.idleS) + ' parado · ' + fuelText + ' · ' + pctFmt.format(v.share) + ' do tempo de motor ligado'));

    if (v.drivers.length) {
      var list = el('ol', 'cia-drivers');
      v.drivers.forEach(function (d, i) {
        var item = el('li', 'cia-driver');
        item.appendChild(el('span', 'cia-driver-pos', String(i + 1)));
        item.appendChild(el('span', 'cia-driver-name', d.name));
        item.appendChild(el('span', 'cia-driver-detail',
          pctFmt.format(d.share) + ' do tempo parado · ' + hours(d.idleS) + ' · ' + d.trips + (d.trips === 1 ? ' viagem' : ' viagens')));
        item.appendChild(el('span', 'cia-driver-share', brl.format(d.cost)));
        list.appendChild(item);
      });
      li.appendChild(list);
    }
    if (v.unknownShare > 0) {
      li.appendChild(el('p', 'cia-vehicle-unknown', pctFmt.format(v.unknownShare) + ' sem motorista identificado'));
    }
    return li;
  }

  var STATUS_TEXT = { ok: 'Disponível', partial: 'Parcial', none: 'Sem dados' };

  function renderCheckup(assessment) {
    var box = $('cia-checkup');
    clear(box);
    assessment.pillars.forEach(function (p) {
      var card = el('div', 'cia-pillar');
      var head = el('div', 'cia-pillar-head');
      head.appendChild(el('span', 'cia-pillar-name', p.name));
      head.appendChild(el('span', 'cia-badge cia-badge-' + p.status, STATUS_TEXT[p.status]));
      card.appendChild(head);
      card.appendChild(el('p', 'cia-pillar-question', p.question));
      var list = el('ul', 'cia-checks');
      p.items.forEach(function (it) {
        var li = el('li', 'cia-check' + (it.ok ? ' cia-check-ok' : ''));
        li.appendChild(el('span', 'cia-check-mark', it.ok ? '✓' : '–'));
        var text = el('span', 'cia-check-text', it.label);
        if (it.detail) text.appendChild(el('span', 'cia-check-detail', ' · ' + it.detail));
        li.appendChild(text);
        list.appendChild(li);
      });
      card.appendChild(list);
      box.appendChild(card);
    });
    var partial = assessment.pillars.filter(function (p) { return p.status === 'partial'; }).length;
    $('cia-checkup-sub').textContent = 'Com os dados que já envia à Geotab, esta frota permite analisar ' + assessment.ready + ' de 3 frentes de economia' +
      (partial ? ' (' + partial + ' delas' + ' em parte).' : '.');
    $('cia-checkup-card').hidden = false;
  }

  var last = null; // último resultado, para recalcular quando o preço muda

  function render() {
    if (!last) return;
    var price = readPrice();
    var s = valueIdle(last.summary, last.fuel, price);
    var dateFmt = new Intl.DateTimeFormat('pt-BR', { day: '2-digit', month: '2-digit' });
    $('cia-period').textContent = 'De ' + dateFmt.format(last.start) + ' a ' + dateFmt.format(last.end);
    $('cia-kpi-cost').textContent = brl.format(s.totalCost);
    $('cia-kpi-liters').textContent = liters(s.totalLiters);
    $('cia-kpi-hours').textContent = hours(s.totalIdleS);
    $('cia-kpi-share').textContent = pctFmt.format(s.share) + ' do tempo de motor ligado';
    $('cia-kpi-vehicles').textContent = String(s.vehicles);
    $('cia-kpis').hidden = false;

    var n = s.ranking.length;
    var basis;
    if (n === 0) basis = '';
    else if (s.measuredVehicles === n) basis = 'Litros medidos pelos próprios veículos.';
    else if (s.measuredVehicles === 0) basis = 'Litros estimados em ' + IDLE_RATE_LPH + ' L por hora parado, média de caminhão pesado: estes veículos não informam o combustível gasto parado.';
    else basis = 'Litros medidos em ' + s.measuredVehicles + ' de ' + n + ' veículos; nos demais, estimados em ' + IDLE_RATE_LPH + ' L por hora parado.';
    $('cia-basis').textContent = basis;

    var input = $('cia-price');
    if (document.activeElement !== input) input.value = priceFmt.format(price);
    $('cia-price-source').textContent = price === DEFAULT_PRICE ? PRICE_SOURCE : 'valor informado por você';

    var list = $('cia-ranking');
    clear(list);
    s.ranking.slice(0, TOP_VEHICLES).forEach(function (v) { list.appendChild(vehicleItem(v)); });
    $('cia-ranking-card').hidden = n === 0;
    setStatus(n === 0 ? 'Nenhum tempo de motor ligado parado no período.' : '');
  }

  function bindPrice() {
    var input = $('cia-price');
    if (!input || input.getAttribute('data-bound')) return;
    input.setAttribute('data-bound', '1');
    input.addEventListener('change', function () {
      var v = parsePrice(input.value);
      if (v) savePrice(v);
      input.value = priceFmt.format(readPrice());
      render();
    });
  }

  function call(api, method, params) {
    return new Promise(function (resolve, reject) { api.call(method, params, resolve, reject); });
  }
  // Uma checagem que falha (sinal inexistente nesta base, por exemplo) vira "sem dados"
  function safe(p) { return p.then(function (r) { return r; }, function () { return null; }); }
  function exists(api, typeName, search) {
    var params = { typeName: typeName, resultsLimit: 1 };
    if (search) params.search = search;
    return safe(call(api, 'Get', params))
      .then(function (r) { return !!(r && r.length); });
  }

  function load(api, state) {
    var end = new Date();
    var start = new Date(end.getTime() - DAYS * 86400000);
    var from = start.toISOString(), to = end.toISOString();
    var groups = state && state.getGroupFilter ? state.getGroupFilter() : [];
    setStatus('Carregando dados da frota…');
    bindPrice();
    var deviceSearch = { fromDate: from };
    if (groups && groups.length) deviceSearch.groups = groups;
    function signal(id) { return { diagnosticSearch: { id: id }, fromDate: from, toDate: to }; }

    return Promise.all([
      call(api, 'Get', { typeName: 'Device', search: deviceSearch }),
      call(api, 'Get', { typeName: 'Trip', search: { fromDate: from, toDate: to } }),
      call(api, 'Get', { typeName: 'User', search: { isDriver: true } }),
      safe(call(api, 'Get', { typeName: 'StatusData', search: signal(DIAG.idleFuel) }))
    ]).then(function (res) {
      var devices = res[0], trips = res[1];
      var inFilter = {};
      devices.forEach(function (d) { inFilter[d.id] = true; });
      var fuelAll = idleFuelByDevice(res[3] || []);
      var fuel = {};
      Object.keys(fuelAll).forEach(function (id) { if (inFilter[id]) fuel[id] = fuelAll[id]; });
      var tripsIn = trips.filter(function (t) { return inFilter[refId(t.device)]; });

      last = { summary: summarizeIdle(devices, trips, res[2]), fuel: fuel, start: start, end: end };
      render();

      return Promise.all([
        exists(api, 'StatusData', signal(DIAG.rpm)),
        exists(api, 'StatusData', signal(DIAG.totalFuel)),
        exists(api, 'StatusData', signal(DIAG.coolant)),
        exists(api, 'FaultData', { fromDate: from, toDate: to }),
        exists(api, 'Zone')
      ]).then(function (c) {
        renderCheckup(assessData({
          devices: devices, trips: tripsIn, idleFuelVehicles: Object.keys(fuel).length,
          rpm: c[0], totalFuel: c[1], coolant: c[2], faults: c[3], zones: c[4]
        }));
      });
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
    module.exports = {
      VERSION: VERSION, DEFAULT_PRICE: DEFAULT_PRICE, IDLE_RATE_LPH: IDLE_RATE_LPH,
      parseDuration: parseDuration, summarizeIdle: summarizeIdle, idleFuelByDevice: idleFuelByDevice,
      valueIdle: valueIdle, assessData: assessData, parsePrice: parsePrice
    };
  }
})();
