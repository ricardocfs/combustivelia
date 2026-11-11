/* Combustível IA · Add-In de página para o MyGeotab (versão 0.6)
 * Abre com "O que fazer esta semana": poucas ações, cada uma com por quê, o que
 * verificar, o que fazer e quanto vale. Depois, motor ligado parado por veículo, em horas, litros e R$, com os 3 motoristas
 * que mais contribuíram em cada veículo. Check-up de dados: quais frentes de
 * economia esta frota permite analisar. Somente leitura. Respeita o filtro de grupos.
 */
(function () {
  'use strict';

  var VERSION = '0.6.1';
  var DAYS = 7;
  var TOP_VEHICLES = 10;
  var TOP_DRIVERS = 3;
  var UNKNOWN_DRIVER = 'UnknownDriverId';
  var LONG_IDLE_S = 3600;       // viagem com 1 h ou mais de motor ligado parado
  var MAX_ACTIONS = 5;          // um assessor dá poucas ações, as que mais valem
  var MIN_WEEK_COST = 25;       // abaixo disso (R$ na semana) não vira ação
  var OUTLIER = 1.5;            // "bem acima" = 50% acima da referência

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
      byId[d.id] = { id: d.id, name: d.name || d.id, idleS: 0, drivingS: 0, unknownS: 0, longIdleS: 0, byDriver: {} };
    });
    trips.forEach(function (t) {
      var row = byId[refId(t.device)];
      if (!row) return; // veículo fora do filtro de grupos
      var idle = parseDuration(t.idlingDuration);
      var driving = parseDuration(t.drivingDuration);
      var long = idle >= LONG_IDLE_S ? idle : 0; // viagem com 1 h ou mais parada: espera, não hábito
      row.idleS += idle;
      row.drivingS += driving;
      row.longIdleS += long;
      var id = refId(t.driver);
      if (!id || id === UNKNOWN_DRIVER) { row.unknownS += idle; return; }
      var d = row.byDriver[id] || (row.byDriver[id] = { id: id, name: names[id] || 'Motorista ' + id, idleS: 0, engineS: 0, longIdleS: 0, trips: 0 });
      d.idleS += idle;
      d.engineS += idle + driving;
      d.longIdleS += long;
      d.trips += 1;
    });

    var totalIdle = 0, totalEngine = 0;
    var rows = Object.keys(byId).map(function (k) {
      var r = byId[k];
      var engine = r.idleS + r.drivingS;
      totalIdle += r.idleS;
      totalEngine += engine;
      var drivers = Object.keys(r.byDriver).map(function (id) { return r.byDriver[id]; })
        .filter(function (d) { return d.idleS > 0; });
      drivers.forEach(function (d) { d.share = r.idleS > 0 ? d.idleS / r.idleS : 0; });
      drivers.sort(function (a, b) { return b.idleS - a.idleS; });
      return {
        id: r.id,
        name: r.name,
        idleS: r.idleS,
        engineS: engine,
        longIdleS: r.longIdleS,
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

  /** O assessor: transforma os números em poucas ações, cada uma dizendo o quê,
   *  quem, por quê, o que verificar, o que fazer e quanto vale. Antes de apontar
   *  um motorista, separa o que é da pessoa, do caminhão e da operação. */
  function advise(s, days) {
    var monthFactor = 30 / (days || DAYS);
    var fleetRate = s.share;
    var price = s.price;
    var actions = [];

    s.ranking.forEach(function (v) {
      if (v.cost < MIN_WEEK_COST || v.idleS <= 0) return;
      var litersPerS = v.liters / v.idleS;
      var longShare = v.longIdleS / v.idleS;
      function money(idleS) { return Math.max(0, idleS) * litersPerS * price * monthFactor; }

      if (longShare >= 0.5) {
        var opValue = money(v.longIdleS);
        actions.push({
          type: 'operacao', tag: 'Operação', short: 'Ver onde ele espera', vehicleId: v.id, value: opValue, sortKey: opValue,
          valueLabel: 'até ' + brl.format(opValue) + '/mês',
          title: 'Descobrir onde o ' + v.name + ' espera com o motor ligado',
          why: pctFmt.format(longShare) + ' do tempo parado dele foi em viagens com mais de 1 hora parada (' + hours(v.longIdleS) +
            '). Isso costuma ser espera em cliente, doca ou pernoite, não hábito do motorista.',
          check: 'Em Histórico de viagens, veja onde o ' + v.name + ' ficou parado nessas viagens. Se é sempre o mesmo lugar, o desperdício está ali.',
          todo: 'Negocie uma janela de carga com o cliente ou oriente desligar o motor durante a espera.'
        });
      } else {
        // Pessoa: o motorista fica parado bem mais que os colegas no MESMO caminhão?
        var best = null;
        v.drivers.forEach(function (d) {
          if (!(d.engineS > 0) || d.cost < MIN_WEEK_COST || d.longIdleS / d.idleS >= 0.5) return;
          var rate = d.idleS / d.engineS;
          var peersEngine = v.engineS - d.engineS;
          var peers = peersEngine >= 2 * 3600;
          var ref = peers ? (v.idleS - d.idleS) / peersEngine : fleetRate;
          if (!(ref > 0) || rate < OUTLIER * ref) return;
          var value = money(d.idleS - ref * d.engineS);
          if (!best || value > best.value) best = { d: d, rate: rate, ref: ref, peers: peers, value: value };
        });
        if (best) {
          actions.push({
            type: 'conducao', tag: 'Condução', short: 'Falar com ' + best.d.name, vehicleId: v.id, value: best.value, sortKey: best.value,
            valueLabel: 'até ' + brl.format(best.value) + '/mês',
            title: 'Conversar com ' + best.d.name + ' sobre o motor ligado parado',
            why: 'No ' + v.name + ', ' + best.d.name + ' ficou com o motor ligado parado ' + pctFmt.format(best.rate) + ' do tempo; ' +
              (best.peers ? 'os outros motoristas do mesmo caminhão ficam em ' : 'a média da frota é ') + pctFmt.format(best.ref) +
              '. Foram ' + hours(best.d.idleS) + ' e ' + brl.format(best.d.cost) + ' nesta semana.',
            check: 'Pergunte se houve espera para carregar, descarregar ou pernoite. Se houve, o problema é da operação, não do motorista.',
            todo: 'Se não houve espera, combine desligar o motor nas paradas e acompanhe aqui na próxima semana.'
          });
        } else if (fleetRate > 0 && v.share >= OUTLIER * fleetRate) {
          // Caminhão: todos que dirigem ficam parados demais, sem esperas longas
          var vValue = money(v.idleS - fleetRate * v.engineS);
          actions.push({
            type: 'equipe', tag: 'Condução', short: 'Regra para a equipe', vehicleId: v.id, value: vValue, sortKey: vValue,
            valueLabel: 'até ' + brl.format(vValue) + '/mês',
            title: 'Combinar uma regra de motor parado para o ' + v.name,
            why: 'O ' + v.name + ' fica com o motor ligado parado ' + pctFmt.format(v.share) + ' do tempo, contra ' + pctFmt.format(fleetRate) +
              ' da média da frota, sem se concentrar em um motorista nem em esperas longas. Foram ' + brl.format(v.cost) + ' nesta semana.',
            check: 'Veja se algo obriga a deixar o motor ligado parado nesse caminhão, como equipamento acionado pelo motor.',
            todo: 'Combine com todos que dirigem esse caminhão: motor desligado nas paradas. Se o modelo permitir, peça à concessionária para programar o desligamento automático em marcha lenta.'
          });
        }
      }

      // Dados: sem saber quem dirigia, ninguém pode ser orientado
      var unknownCost = v.cost * v.unknownShare;
      if (v.unknownShare >= 0.3 && unknownCost >= MIN_WEEK_COST) {
        var uValue = unknownCost * monthFactor;
        actions.push({
          type: 'dados', tag: 'Dados', short: 'Identificar motorista', vehicleId: v.id, value: 0, sortKey: uValue * 0.5,
          valueLabel: brl.format(uValue) + '/mês sem dono',
          title: 'Identificar quem dirige o ' + v.name,
          why: pctFmt.format(v.unknownShare) + ' do tempo parado dele (' + brl.format(unknownCost) +
            ' nesta semana) não tem motorista identificado. Sem saber quem dirigia, não dá para orientar ninguém.',
          check: 'Veja se o caminhão tem leitor de identificação de motorista e se os motoristas estão usando.',
          todo: 'Ative a identificação (chaveiro, cartão ou login no Geotab Drive) e cobre o uso na saída do pátio.'
        });
      }
    });

    actions.sort(function (a, b) { return b.sortKey - a.sortKey; });
    var top = actions.slice(0, MAX_ACTIONS);
    return {
      actions: top,                 // as que mais valem (relatório por e-mail)
      all: actions,                 // todas, para mostrar dentro de cada veículo
      monthlyValue: top.reduce(function (sum, a) { return sum + a.value; }, 0)
    };
  }

  // ---- Evolução da frota ------------------------------------------------------
  var WEEKS = 6;            // semanas no gráfico
  var BASE_WEEKS = 4;       // histórico do próprio veículo para comparar a última semana
  var MIN_KM_WEEK = 300;    // abaixo disso a semana do veículo não entra na comparação
  var NOISE = 0.03;         // variação de eficiência menor que 3% é ruído
  var TOP_MOVERS = 5;
  var WEEK_MS = 7 * 86400000;

  /** Eficiência (km/L) da frota por semana e os veículos que mais pioraram e
   *  melhoraram, cada um contra a própria média das 4 semanas anteriores.
   *  fuel: registros FuelUsed/FuelAndEnergyUsed (litros por evento). */
  function evolution(o) {
    var end = o.end.getTime();
    var n = o.weeks || WEEKS;
    var names = o.names || {};
    var price = o.price;
    function weekOf(t) {
      var ms = new Date(t).getTime();
      if (!isFinite(ms) || ms > end) return -1;
      var w = n - 1 - Math.floor((end - ms) / WEEK_MS);
      return w >= 0 ? w : -1;
    }
    function zeros() { var a = []; for (var i = 0; i < n; i++) a.push(0); return a; }
    var byId = {};
    (o.devices || []).forEach(function (d) {
      byId[d.id] = { id: d.id, name: d.name || d.id, km: zeros(), liters: zeros(), idleL: zeros(), drivingS: zeros(), faults: zeros(), driverKm: [] };
      for (var i = 0; i < n; i++) byId[d.id].driverKm.push({});
    });
    (o.trips || []).forEach(function (t) {
      var v = byId[refId(t.device)]; var w = weekOf(t.start || t.stop);
      if (!v || w < 0) return;
      var km = Number(t.distance) || 0;
      v.km[w] += km;
      v.drivingS[w] += parseDuration(t.drivingDuration);
      var id = refId(t.driver);
      if (id && id !== UNKNOWN_DRIVER) v.driverKm[w][id] = (v.driverKm[w][id] || 0) + km;
    });
    (o.fuel || []).forEach(function (f) {
      var v = byId[refId(f.device)]; var w = weekOf(f.dateTime);
      if (!v || w < 0) return;
      v.liters[w] += Number(f.totalFuelUsed) || 0;
      v.idleL[w] += Number(f.totalIdlingFuelUsedL) || 0;
    });
    (o.faults || []).forEach(function (f) {
      var v = byId[refId(f.device)]; var w = weekOf(f.dateTime);
      if (v && w >= 0) v.faults[w] += 1;
    });

    var vehicles = Object.keys(byId).map(function (k) { return byId[k]; });
    var weeks = [];
    for (var w = 0; w < n; w++) {
      var km = 0, l = 0, idle = 0;
      vehicles.forEach(function (v) {
        if (v.km[w] > 0 && v.liters[w] > 0) { km += v.km[w]; l += v.liters[w]; }
        idle += v.idleL[w];
      });
      weeks.push({ start: new Date(end - (n - w) * WEEK_MS), km: km, liters: l, kml: l > 0 ? km / l : 0, idleLiters: idle, idleCost: idle * price });
    }
    var lastW = n - 1;
    var baseFrom = Math.max(0, lastW - BASE_WEEKS);
    function sum(a, from, to) { var x = 0; for (var i = from; i < to; i++) x += a[i]; return x; }
    function mainDriver(map) {
      var best = null;
      Object.keys(map).forEach(function (id) { if (!best || map[id] > map[best]) best = id; });
      return best;
    }

    var movers = [];
    vehicles.forEach(function (v) {
      var kmL = v.km[lastW], lL = v.liters[lastW];
      var kmB = sum(v.km, baseFrom, lastW), lB = sum(v.liters, baseFrom, lastW);
      if (kmL < MIN_KM_WEEK || !(lL > 0) || kmB < 2 * MIN_KM_WEEK || !(lB > 0)) return;
      var kmlL = kmL / lL, kmlB = kmB / lB;
      if (kmlL < 0.5 || kmlL > 40 || kmlB < 0.5 || kmlB > 40) return; // leitura implausível
      var change = (kmlL - kmlB) / kmlB;
      if (Math.abs(change) < NOISE) return;
      var extraL = lL - kmL / kmlB;      // + gastou a mais; − economizou
      var causes = [], kind = 'none';
      var idleShareL = v.idleL[lastW] / lL, idleShareB = sum(v.idleL, baseFrom, lastW) / lB;
      var dS = v.drivingS[lastW], dB = sum(v.drivingS, baseFrom, lastW);
      var speedL = dS > 0 ? kmL / (dS / 3600) : 0, speedB = dB > 0 ? kmB / (dB / 3600) : 0;
      var drvL = mainDriver(v.driverKm[lastW]);
      var drvSeen = false;
      for (var i = baseFrom; i < lastW; i++) if (drvL && v.driverKm[i][drvL]) drvSeen = true;
      var drvName = drvL ? (names[drvL] || 'Motorista ' + drvL) : null;
      if (change < 0) {
        if (v.faults[lastW] > 0) { causes.push(v.faults[lastW] + (v.faults[lastW] === 1 ? ' registro' : ' registros') + ' de falha nesta semana.'); kind = 'faults'; }
        if (drvL && !drvSeen) { causes.push('Motorista diferente: ' + drvName + ' fez a maior parte dos quilômetros e não dirigiu este veículo nas 4 semanas anteriores.'); if (kind === 'none') kind = 'driver'; }
        if (idleShareL - idleShareB >= 0.05) { causes.push('O motor ligado parado subiu de ' + pctFmt.format(idleShareB) + ' para ' + pctFmt.format(idleShareL) + ' do diesel gasto.'); if (kind === 'none') kind = 'idle'; }
        if (speedB > 0 && speedL / speedB - 1 >= 0.08) { causes.push('Velocidade média subiu de ' + Math.round(speedB) + ' para ' + Math.round(speedL) + ' km/h.'); if (kind === 'none') kind = 'speed'; }
        if (!causes.length) causes.push('Nada nos dados de condução explica a piora: é o tipo de caso para a oficina.');
      } else {
        if (idleShareB - idleShareL >= 0.05) causes.push('O motor ligado parado caiu de ' + pctFmt.format(idleShareB) + ' para ' + pctFmt.format(idleShareL) + ' do diesel gasto.');
        if (speedB > 0 && 1 - speedL / speedB >= 0.08) causes.push('Velocidade média caiu de ' + Math.round(speedB) + ' para ' + Math.round(speedL) + ' km/h.');
        if (drvL && !drvSeen) causes.push('Motorista diferente: ' + drvName + ' fez a maior parte dos quilômetros.');
        if (!causes.length) causes.push('Os dados de condução não mostram uma mudança clara; pode ser rota ou carga mais leve.');
      }
      var todo = {
        faults: 'Leve à oficina e peça para verificar as falhas registradas nesta semana.',
        driver: 'Converse com ' + drvName + ' e compare a rota com a das semanas anteriores.',
        idle: 'Abra o veículo na aba Esta semana: o tempo parado explica a piora.',
        speed: 'Combine com quem dirige o limite de velocidade na estrada.',
        none: 'Peça uma inspeção: filtros de ar e de combustível, calibragem dos pneus e injeção.'
      }[kind];
      if (change > 0) todo = drvName ? 'Reconheça o resultado com ' + drvName + '.' : 'Veja o que mudou para repetir nos outros veículos.';
      movers.push({ id: v.id, name: v.name, kmlLast: kmlL, kmlBase: kmlB, change: change, km: kmL, liters: lL,
        extraLiters: extraL, value: extraL * price, causes: causes, todo: todo, kind: change < 0 ? kind : 'better' });
    });
    var worse = movers.filter(function (m) { return m.change < 0; }).sort(function (a, b) { return b.value - a.value; }).slice(0, TOP_MOVERS);
    var better = movers.filter(function (m) { return m.change > 0; }).sort(function (a, b) { return a.value - b.value; }).slice(0, TOP_MOVERS);

    var last = weeks[lastW], baseKm = 0, baseL = 0;
    for (var j = baseFrom; j < lastW; j++) { baseKm += weeks[j].km; baseL += weeks[j].liters; }
    var baseKml = baseL > 0 ? baseKm / baseL : 0;
    return {
      weeks: weeks, hasFuel: weeks.some(function (x) { return x.liters > 0; }),
      kml: last.kml, baseKml: baseKml, change: baseKml > 0 && last.kml > 0 ? (last.kml - baseKml) / baseKml : 0,
      liters: last.liters, cost: last.liters * price, compared: movers.length,
      worse: worse, better: better
    };
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

  function driversBlock(v) {
    var box = el('div', 'cia-row-drivers');
    box.appendChild(el('p', 'cia-row-label', 'Quem dirigia quando ficou parado'));
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
      box.appendChild(list);
    }
    if (v.unknownShare > 0) box.appendChild(el('p', 'cia-vehicle-unknown', pctFmt.format(v.unknownShare) + ' sem motorista identificado'));
    return box;
  }

  function actionBlock(a) {
    var box = el('div', 'cia-row-action');
    box.appendChild(el('span', 'cia-tag cia-tag-' + a.type, a.tag));
    box.appendChild(el('p', 'cia-action-title', a.title));
    var body = el('dl', 'cia-action-body');
    [['Por quê', a.why], ['Verifique', a.check], ['Sugestão de ação', a.todo],
     [a.type === 'dados' ? 'Sem dono' : 'Economia possível', a.valueLabel]].forEach(function (pair) {
      body.appendChild(el('dt', 'cia-action-label', pair[0]));
      body.appendChild(el('dd', 'cia-action-text', pair[1]));
    });
    box.appendChild(body);
    return box;
  }

  /** Uma linha por veículo: o essencial à vista; o porquê, o que verificar e a
   *  sugestão de ação abrem com um clique. */
  /** Por que um veículo não virou ação: o assessor também explica o normal. */
  function normalReason(v, fleetRate) {
    var own = pctFmt.format(v.share), fleet = pctFmt.format(fleetRate);
    if (!(fleetRate > 0) || v.share < OUTLIER * fleetRate) {
      return 'O motor fica ligado parado ' + own + ' do tempo, ' + (v.share <= fleetRate ? 'abaixo' : 'perto') + ' da média da frota (' + fleet + '). ' +
        'Nenhum motorista se destaca dos colegas e não há esperas longas. Nada fora do padrão nesta semana.';
    }
    if (v.cost < MIN_WEEK_COST) {
      return 'O motor fica ligado parado ' + own + ' do tempo, acima da média da frota (' + fleet + '), mas o custo na semana (' +
        brl.format(v.cost) + ') é pequeno demais para virar ação. Vale acompanhar se continuar subindo.';
    }
    return 'O tempo parado (' + own + ' contra ' + fleet + ' da frota) não aponta para um motorista, para o caminhão nem para esperas longas.';
  }

  function vehicleRow(v, rank, actions, open, fleetRate) {
    var li = el('li', 'cia-row');
    var details = el('details', 'cia-row-details');
    if (open) details.open = true;
    var summary = el('summary', 'cia-row-summary');
    summary.appendChild(el('span', 'cia-row-pos', String(rank)));
    var main = el('span', 'cia-row-main');
    main.appendChild(el('span', 'cia-row-name', v.name));
    main.appendChild(el('span', 'cia-row-meta', hours(v.idleS) + ' parado · ' + pctFmt.format(v.share) + ' do tempo de motor ligado'));
    summary.appendChild(main);
    var side = el('span', 'cia-row-side');
    side.appendChild(el('span', 'cia-row-cost', brl.format(v.cost)));
    var lead = actions.filter(function (a) { return a.type !== 'dados'; })[0] || actions[0];
    side.appendChild(el('span', 'cia-row-hint' + (lead ? ' cia-row-hint-action' : ''),
      lead ? lead.short + (actions.length > 1 ? ' + ' + (actions.length - 1) : '') : 'Dentro do normal'));
    summary.appendChild(side);
    summary.appendChild(el('span', 'cia-row-chevron', '›'));
    details.appendChild(summary);

    var body = el('div', 'cia-row-body');
    if (actions.length) actions.forEach(function (a) { body.appendChild(actionBlock(a)); });
    else {
      var ok = el('div', 'cia-row-action');
      ok.appendChild(el('span', 'cia-tag', 'Dentro do normal'));
      ok.appendChild(el('p', 'cia-row-ok', normalReason(v, fleetRate)));
      body.appendChild(ok);
    }
    var fuelText = v.estimated ? '≈ ' + liters(v.liters) + ' estimados' : liters(v.liters) + ' medidos pelo veículo';
    body.appendChild(el('p', 'cia-row-fuel', 'Diesel queimado parado: ' + fuelText + '.'));
    body.appendChild(driversBlock(v));
    details.appendChild(body);
    li.appendChild(details);
    return li;
  }

  function renderTop(s, advice) {
    var byVehicle = {};
    advice.all.forEach(function (a) { (byVehicle[a.vehicleId] = byVehicle[a.vehicleId] || []).push(a); });
    function priority(v) {
      return (byVehicle[v.id] || []).reduce(function (x, a) { return x + a.sortKey; }, 0);
    }
    // Quem tem ação sugerida vem primeiro, em ordem de valor; depois, os de maior gasto.
    var ordered = s.ranking.slice().sort(function (a, b) {
      var pa = priority(a), pb = priority(b);
      if ((pa > 0) !== (pb > 0)) return pb > 0 ? 1 : -1;
      return pa !== pb ? pb - pa : b.cost - a.cost;
    });
    var rows = ordered.slice(0, TOP_VEHICLES);
    var list = $('cia-ranking');
    clear(list);
    var opened = false;
    rows.forEach(function (v, i) {
      var acts = byVehicle[v.id] || [];
      var open = !opened && acts.length > 0;
      if (open) opened = true;
      list.appendChild(vehicleRow(v, i + 1, acts, open, s.share));
    });
    var withAction = Object.keys(byVehicle).length;
    var value = advice.all.reduce(function (x, a) { return x + a.value; }, 0);
    var sub;
    if (!rows.length) sub = '';
    else if (!withAction) sub = 'Nenhuma ação urgente: todos os veículos ficam perto da média da frota (' + pctFmt.format(s.share) +
      ' do tempo de motor ligado). Clique em um veículo para ver por que ele está dentro do normal.';
    else {
      sub = (withAction === 1 ? '1 veículo tem' : withAction + ' veículos têm') + ' ação sugerida';
      sub += value > 0 ? ', que pode recuperar até ' + brl.format(value) + ' por mês.' : '.';
      sub += ' Eles vêm primeiro, em ordem de valor; depois, os de maior gasto. Clique para ver o porquê.';
    }
    $('cia-ranking-title').textContent = rows.length < TOP_VEHICLES ? 'Veículos com motor ligado parado' : 'Top ' + TOP_VEHICLES + ' veículos';
    $('cia-ranking-sub').textContent = sub;
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

  var kmlFmt = new Intl.NumberFormat('pt-BR', { minimumFractionDigits: 1, maximumFractionDigits: 1 });
  var signedPct = new Intl.NumberFormat('pt-BR', { style: 'percent', maximumFractionDigits: 0, signDisplay: 'exceptZero' });
  var dayMonth = new Intl.DateTimeFormat('pt-BR', { day: '2-digit', month: '2-digit' });

  function barChart(box, values, labels, format) {
    clear(box);
    var max = Math.max.apply(null, values.concat([0]));
    values.forEach(function (v, i) {
      var col = el('div', 'cia-bar-col' + (i === values.length - 1 ? ' cia-bar-last' : ''));
      col.appendChild(el('span', 'cia-bar-value', v > 0 ? format(v) : '–'));
      var bar = el('div', 'cia-bar');
      bar.style.height = (max > 0 ? Math.max(2, Math.round(v / max * 100)) : 2) + '%';
      var track = el('div', 'cia-bar-track');
      track.appendChild(bar);
      col.appendChild(track);
      col.appendChild(el('span', 'cia-bar-label', labels[i]));
      box.appendChild(col);
    });
  }

  function moverRow(m, rank, worse) {
    var li = el('li', 'cia-row');
    var details = el('details', 'cia-row-details');
    var summary = el('summary', 'cia-row-summary');
    summary.appendChild(el('span', 'cia-row-pos', String(rank)));
    var main = el('span', 'cia-row-main');
    main.appendChild(el('span', 'cia-row-name', m.name));
    main.appendChild(el('span', 'cia-row-meta', kmlFmt.format(m.kmlBase) + ' → ' + kmlFmt.format(m.kmlLast) + ' km/L · ' + litersFmt.format(m.km) + ' km na semana'));
    summary.appendChild(main);
    var side = el('span', 'cia-row-side');
    side.appendChild(el('span', 'cia-row-cost' + (worse ? '' : ' cia-good'), (worse ? '+' : '−') + brl.format(Math.abs(m.value))));
    side.appendChild(el('span', 'cia-row-hint', signedPct.format(m.change) + ' de eficiência'));
    summary.appendChild(side);
    summary.appendChild(el('span', 'cia-row-chevron', '›'));
    details.appendChild(summary);
    var body = el('div', 'cia-row-body');
    var box = el('div', 'cia-row-action');
    var dl = el('dl', 'cia-action-body');
    dl.appendChild(el('dt', 'cia-action-label', worse ? 'Possíveis causas' : 'O que mudou'));
    var dd = el('dd', 'cia-action-text');
    var ul = el('ul', 'cia-causes');
    m.causes.forEach(function (c) { ul.appendChild(el('li', '', c)); });
    dd.appendChild(ul);
    dl.appendChild(dd);
    dl.appendChild(el('dt', 'cia-action-label', 'Sugestão de ação'));
    dl.appendChild(el('dd', 'cia-action-text', m.todo));
    dl.appendChild(el('dt', 'cia-action-label', worse ? 'Gasto a mais' : 'Economia'));
    dl.appendChild(el('dd', 'cia-action-text', liters(Math.abs(m.extraLiters)) + ' (' + brl.format(Math.abs(m.value)) + ') nesta semana, contra a média das 4 anteriores'));
    box.appendChild(dl);
    body.appendChild(box);
    details.appendChild(body);
    li.appendChild(details);
    return li;
  }

  function renderEvolution(e) {
    var status = $('cia-evo-status');
    if (!e.hasFuel) {
      status.textContent = 'Esta frota não envia o consumo de combustível à Geotab, então não dá para medir a eficiência (km/L).';
      $('cia-evo-content').hidden = true;
      return;
    }
    status.textContent = '';
    $('cia-evo-kml').textContent = e.kml > 0 ? kmlFmt.format(e.kml) + ' km/L' : '—';
    $('cia-evo-kml-note').textContent = e.baseKml > 0 && e.kml > 0
      ? (Math.abs(e.change) < NOISE ? 'Estável em relação à média das 4 semanas anteriores'
        : signedPct.format(e.change) + (e.change > 0 ? ', melhor' : ', pior') + ' que a média das 4 semanas anteriores')
      : '';
    $('cia-evo-liters').textContent = liters(e.liters);
    $('cia-evo-cost').textContent = brl.format(e.cost);
    $('cia-evo-compared').textContent = String(e.compared);
    var labels = e.weeks.map(function (w) { return dayMonth.format(w.start); });
    barChart($('cia-chart-kml'), e.weeks.map(function (w) { return w.kml; }), labels, function (v) { return kmlFmt.format(v); });
    barChart($('cia-chart-idle'), e.weeks.map(function (w) { return w.idleCost; }), labels, function (v) { return brl.format(v); });

    var worse = $('cia-worse'), better = $('cia-better');
    clear(worse); clear(better);
    e.worse.forEach(function (m, i) { worse.appendChild(moverRow(m, i + 1, true)); });
    e.better.forEach(function (m, i) { better.appendChild(moverRow(m, i + 1, false)); });
    var why = e.compared === 0
      ? 'Nenhum veículo rodou 300 km ou mais na semana com consumo registrado, então ainda não há comparação justa.'
      : null;
    $('cia-worse-empty').textContent = why || 'Nenhum veículo ficou mais de 3% menos eficiente que a própria média das 4 semanas anteriores. Numa semana normal, é o esperado.';
    $('cia-better-empty').textContent = why || 'Nenhum veículo ficou mais de 3% mais eficiente que a própria média das 4 semanas anteriores.';
    $('cia-worse-empty').hidden = e.worse.length > 0;
    $('cia-better-empty').hidden = e.better.length > 0;
    var total = e.worse.reduce(function (x, m) { return x + m.value; }, 0);
    $('cia-worse-sub').textContent = e.worse.length
      ? 'Juntos, gastaram ' + brl.format(total) + ' a mais nesta semana do que a própria média. Clique para ver as possíveis causas.'
      : '';
    $('cia-evo-content').hidden = false;
  }

  var evo = { api: null, state: null, loaded: false, loading: false, data: null };

  function loadEvolution() {
    if (evo.loaded || evo.loading || !evo.api) { if (evo.data) renderEvolution(evolution(Object.assign({}, evo.data, { price: readPrice() }))); return; }
    evo.loading = true;
    var api = evo.api, state = evo.state;
    var end = new Date();
    var start = new Date(end.getTime() - WEEKS * WEEK_MS);
    var from = start.toISOString(), to = end.toISOString();
    var groups = state && state.getGroupFilter ? state.getGroupFilter() : [];
    var deviceSearch = { fromDate: from };
    if (groups && groups.length) deviceSearch.groups = groups;
    var range = { fromDate: from, toDate: to };
    $('cia-evo-status').textContent = 'Carregando ' + WEEKS + ' semanas de dados…';
    // FuelAndEnergyUsed é o tipo novo; FuelUsed, o anterior. Usa o que a base tiver.
    var fuelP = safe(call(api, 'Get', { typeName: 'FuelAndEnergyUsed', search: range })).then(function (r) {
      return r && r.length ? r : safe(call(api, 'Get', { typeName: 'FuelUsed', search: range }));
    });
    Promise.all([
      call(api, 'Get', { typeName: 'Device', search: deviceSearch }),
      call(api, 'Get', { typeName: 'Trip', search: range }),
      call(api, 'Get', { typeName: 'User', search: { isDriver: true } }),
      fuelP,
      safe(call(api, 'Get', { typeName: 'FaultData', search: { fromDate: new Date(end.getTime() - WEEK_MS).toISOString(), toDate: to } }))
    ]).then(function (r) {
      evo.data = { devices: r[0], trips: r[1], names: driverNames(r[2]), fuel: r[3] || [], faults: r[4] || [], end: end };
      evo.loaded = true; evo.loading = false;
      renderEvolution(evolution(Object.assign({}, evo.data, { price: readPrice() })));
    }).catch(function (err) {
      evo.loading = false;
      $('cia-evo-status').textContent = 'Não foi possível carregar a evolução. ' + ((err && err.message) || '');
    });
  }

  function showTab(name) {
    ['semana', 'evolucao'].forEach(function (t) {
      var on = t === name;
      $('cia-tab-' + t).setAttribute('aria-selected', on ? 'true' : 'false');
      $('cia-tab-' + t).tabIndex = on ? 0 : -1;
      $('cia-panel-' + t).hidden = !on;
    });
    $('cia-title').textContent = name === 'semana' ? 'Sua frota esta semana' : 'Evolução da frota';
    $('cia-period').hidden = name !== 'semana';
    if (name === 'evolucao') loadEvolution();
  }

  function bindTabs() {
    var bar = $('cia-tabs');
    if (!bar || bar.getAttribute('data-bound')) return;
    bar.setAttribute('data-bound', '1');
    $('cia-tab-semana').addEventListener('click', function () { showTab('semana'); });
    $('cia-tab-evolucao').addEventListener('click', function () { showTab('evolucao'); });
    bar.addEventListener('keydown', function (ev) {
      if (ev.key !== 'ArrowRight' && ev.key !== 'ArrowLeft') return;
      var next = $('cia-tab-semana').getAttribute('aria-selected') === 'true' ? 'evolucao' : 'semana';
      showTab(next); $('cia-tab-' + next).focus();
    });
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

    renderTop(s, advise(s, DAYS));
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
      if (evo.data) renderEvolution(evolution(Object.assign({}, evo.data, { price: readPrice() })));
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
    bindTabs();
    evo = { api: api, state: state, loaded: false, loading: false, data: null };
    if ($('cia-panel-evolucao') && !$('cia-panel-evolucao').hidden) loadEvolution();
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
      valueIdle: valueIdle, advise: advise, evolution: evolution, assessData: assessData, parsePrice: parsePrice
    };
  }
})();
