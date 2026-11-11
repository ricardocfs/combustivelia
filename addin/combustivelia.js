/* Combustível IA · Add-In de página para o MyGeotab (versão 0.8)
 * Abre com "O que fazer esta semana": poucas ações, cada uma com por quê, o que
 * verificar, o que fazer e quanto vale. Depois, motor ligado parado por veículo, em horas, litros e R$, com os 3 motoristas
 * que mais contribuíram em cada veículo. Check-up de dados: quais frentes de
 * economia esta frota permite analisar. Somente leitura. Respeita o filtro de grupos.
 */
(function () {
  'use strict';

  var VERSION = '0.8.1';
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
          // Comparação justa: contra os colegas do mesmo caminhão, mas nunca abaixo da média
          // da frota. Quem fica parado menos que a média da frota não é apontado.
          var peerRate = peersEngine >= 2 * 3600 ? (v.idleS - d.idleS) / peersEngine : -1;
          var peers = peerRate >= fleetRate;
          var ref = peers ? peerRate : fleetRate;
          if (!(ref > 0) || rate < OUTLIER * ref) return;
          var value = money(d.idleS - ref * d.engineS);
          if (!best || value > best.value) best = { d: d, rate: rate, ref: ref, peers: peers, peerRate: peerRate, value: value };
        });
        if (best) {
          actions.push({
            type: 'conducao', tag: 'Condução', short: 'Falar com ' + best.d.name, vehicleId: v.id, driverId: best.d.id, value: best.value, sortKey: best.value,
            valueLabel: 'até ' + brl.format(best.value) + '/mês',
            title: 'Conversar com ' + best.d.name + ' sobre o motor ligado parado',
            why: 'No ' + v.name + ', ' + best.d.name + ' ficou com o motor ligado parado ' + pctFmt.format(best.rate) + ' do tempo; ' +
              (best.peers ? 'os outros motoristas do mesmo caminhão ficam em ' + pctFmt.format(best.ref)
                : 'a média da frota é ' + pctFmt.format(best.ref) +
                  (best.peerRate >= 0 ? ' e os outros motoristas do mesmo caminhão ficam em ' + pctFmt.format(best.peerRate) : '')) +
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
  var WEEKS = 6;            // semanas nos gráficos de gasto e km/L
  var LOAD_WEEKS = 8;       // semanas carregadas: 4 de tendência + 4 de histórico para cada uma
  var TREND_WEEKS = 4;      // semanas no "por onde o diesel está vazando, semana a semana"
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
      var km = 0, l = 0, idle = 0, all = 0;
      vehicles.forEach(function (v) {
        if (v.km[w] > 0 && v.liters[w] > 0) { km += v.km[w]; l += v.liters[w]; }
        idle += v.idleL[w];
        all += v.liters[w];
      });
      weeks.push({ start: new Date(end - (n - w) * WEEK_MS), km: km, liters: l, kml: l > 0 ? km / l : 0,
        totalLiters: all, cost: all * price, idleLiters: idle, idleCost: idle * price });
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
    var byVehicle = {};
    vehicles.forEach(function (v) {
      var kmL = v.km[lastW], lL = v.liters[lastW];
      var kmB = sum(v.km, baseFrom, lastW), lB = sum(v.liters, baseFrom, lastW);
      if (kmL < MIN_KM_WEEK || !(lL > 0) || kmB < 2 * MIN_KM_WEEK || !(lB > 0)) return;
      var kmlL = kmL / lL, kmlB = kmB / lB;
      if (kmlL < 0.5 || kmlL > 40 || kmlB < 0.5 || kmlB > 40) return; // leitura implausível
      var change = (kmlL - kmlB) / kmlB;
      byVehicle[v.id] = { kmlLast: kmlL, kmlBase: kmlB, change: change };
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
        idle: 'Abra o veículo na aba Última semana: o tempo parado explica a piora.',
        speed: 'Combine com quem dirige o limite de velocidade na estrada.',
        none: 'Peça uma inspeção: filtros de ar e de combustível, calibragem dos pneus e injeção.'
      }[kind];
      if (change > 0) todo = 'Veja ' + (drvName ? 'com ' + drvName + ' ' : '') + 'o que mudou: rota, carga ou jeito de dirigir. Se foi o jeito de dirigir, reconheça e use como exemplo.';
      movers.push({ id: v.id, name: v.name, kmlLast: kmlL, kmlBase: kmlB, change: change, km: kmL, liters: lL,
        extraLiters: extraL, value: extraL * price, causes: causes, todo: todo, kind: change < 0 ? kind : 'better',
        consumption: kmlB / kmlL - 1,   // variação do consumo por km (+ = mais diesel por km)
        driverName: drvName, speedLast: speedL, speedBase: speedB, faults: v.faults[lastW] });
    });
    var worseAll = movers.filter(function (m) { return m.change < 0; }).sort(function (a, b) { return b.value - a.value; });
    var betterAll = movers.filter(function (m) { return m.change > 0; }).sort(function (a, b) { return a.value - b.value; });
    var worse = worseAll.slice(0, TOP_MOVERS), better = betterAll.slice(0, TOP_MOVERS);
    var extraValue = worseAll.reduce(function (x, m) { return x + m.value; }, 0);
    var savedValue = -betterAll.reduce(function (x, m) { return x + m.value; }, 0);

    var last = weeks[lastW], baseKm = 0, baseL = 0;
    for (var j = baseFrom; j < lastW; j++) { baseKm += weeks[j].km; baseL += weeks[j].liters; }
    var baseKml = baseL > 0 ? baseKm / baseL : 0;
    function baseAvg(f) { var x = 0, k = 0; for (var i = baseFrom; i < lastW; i++) { x += weeks[i][f]; k += 1; } return k ? x / k : 0; }
    var baseCost = baseAvg('cost'), baseIdle = baseAvg('idleCost');
    return {
      costChange: baseCost > 0 ? (last.cost - baseCost) / baseCost : 0, idleCost: last.idleCost,
      idleChange: baseIdle > 0 ? (last.idleCost - baseIdle) / baseIdle : 0,
      weeks: weeks, hasFuel: weeks.some(function (x) { return x.liters > 0; }),
      kml: last.kml, baseKml: baseKml, change: baseKml > 0 && last.kml > 0 ? (last.kml - baseKml) / baseKml : 0,
      liters: last.totalLiters, cost: last.cost, compared: movers.length,
      extraValue: extraValue, worseCount: worseAll.length, savedValue: savedValue, betterCount: betterAll.length,
      worse: worse, better: better, worseAll: worseAll, byVehicle: byVehicle
    };
  }

  // ---- Mais frentes na tela "Esta semana" ---------------------------------------
  var MONTH_PER_WEEK = 30 / 7;
  var MIN_KM_DRIVER = 500;       // quilometragem mínima de cada lado para comparar motoristas
  var DRIVER_GAP = 0.08;         // diferença de km/L que vale uma conversa
  var DRIVER_DAYS = 35;          // janela da comparação entre motoristas (5 semanas)

  /** Veículos que pioraram contra o próprio histórico viram ações de mecânica ou de
   *  condução, conforme a causa provável. O tempo parado já tem ação própria. */
  function fuelActions(e) {
    var out = [];
    (e.worseAll || []).forEach(function (m) {
      if (m.value < MIN_WEEK_COST || m.kind === 'idle') return;
      var value = m.value * MONTH_PER_WEEK;
      var base = 'Fez ' + kmlFmt.format(m.kmlLast) + ' km/L nesta semana, contra ' + kmlFmt.format(m.kmlBase) +
        ' km/L de média nas 4 anteriores: ' + liters(m.extraLiters) + ' e ' + brl.format(m.value) + ' a mais.';
      var a = { vehicleId: m.id, value: value, sortKey: value, valueLabel: 'até ' + brl.format(value) + '/mês' };
      if (m.kind === 'faults') {
        a.type = 'mecanica'; a.tag = 'Mecânica'; a.short = 'Levar à oficina';
        a.title = 'Levar o ' + m.name + ' à oficina: consumo por km ' + pctFmt.format(m.kmlBase / m.kmlLast - 1) + ' acima da própria média';
        a.why = base + ' Na mesma semana, ' + m.faults + (m.faults === 1 ? ' falha foi registrada.' : ' falhas foram registradas.');
        a.check = 'No MyGeotab, quais falhas apareceram neste veículo. Falhas de injeção, turbo, filtro de partículas ou sensores explicam consumo alto.';
        a.todo = 'Agendar a oficina para verificar as falhas registradas, antes que vire quebra na estrada.';
      } else if (m.kind === 'none') {
        a.type = 'mecanica'; a.tag = 'Mecânica'; a.short = 'Revisar na oficina';
        a.title = 'Revisar o ' + m.name + ': consumo por km ' + pctFmt.format(m.kmlBase / m.kmlLast - 1) + ' acima da própria média';
        a.why = base + ' Tempo parado, velocidade e motorista não mudaram: a causa provável está no caminhão.';
        a.check = 'Se houve mudança de rota, de carga ou de pneus nesta semana. Se não houve, a suspeita fica no caminhão.';
        a.todo = 'Inspeção de filtros de ar e de combustível, calibragem dos pneus e injeção.';
      } else if (m.kind === 'driver') {
        a.type = 'motorista'; a.tag = 'Condução'; a.short = 'Falar com ' + m.driverName;
        a.title = 'Conversar com ' + m.driverName + ' sobre o consumo do ' + m.name;
        a.why = base + ' ' + m.driverName + ' fez a maior parte dos quilômetros e não dirigia este veículo nas 4 semanas anteriores.';
        a.check = 'Se a rota ou a carga mudaram junto com o motorista. Se forem as mesmas, a diferença está no jeito de dirigir.';
        a.todo = 'Mostrar a comparação a ' + m.driverName + ' e combinar marchas mais altas, menos aceleração forte e motor desligado nas paradas.';
      } else if (m.kind === 'speed') {
        a.type = 'velocidade'; a.tag = 'Condução'; a.short = 'Limite de velocidade';
        a.title = 'Combinar o limite de velocidade no ' + m.name;
        a.why = base + ' A velocidade média subiu de ' + Math.round(m.speedBase) + ' para ' + Math.round(m.speedLast) + ' km/h; em caminhão, o consumo cresce rápido com a velocidade.';
        a.check = 'Se o trajeto mudou para mais estrada. Se for a mesma rota, a diferença é do pé no acelerador.';
        a.todo = 'Combinar com quem dirige um limite de velocidade na estrada e acompanhar aqui na próxima semana.';
      } else return;
      out.push(a);
    });
    return out;
  }

  /** Mesmo caminhão, motoristas diferentes: quem gasta mais diesel por km do que os
   *  colegas nas últimas 5 semanas. É a comparação mais justa do jeito de dirigir. */
  function driverGaps(o) {
    var end = o.end.getTime(), from = end - DRIVER_DAYS * 86400000;
    var names = o.names || {}, devName = {};
    (o.devices || []).forEach(function (d) { devName[d.id] = d.name || d.id; });
    var tripsBy = {};
    (o.trips || []).forEach(function (t) {
      var dev = refId(t.device); if (!devName[dev]) return;
      var a = new Date(t.start).getTime();
      if (!(a >= from && a <= end)) return;
      var b = t.stop ? new Date(t.stop).getTime() : a + (parseDuration(t.drivingDuration) + parseDuration(t.idlingDuration)) * 1000;
      (tripsBy[dev] = tripsBy[dev] || []).push({ a: a, b: b, km: Number(t.distance) || 0, driver: refId(t.driver), liters: 0 });
    });
    Object.keys(tripsBy).forEach(function (dev) { tripsBy[dev].sort(function (x, y) { return x.a - y.a; }); });
    var TOL = 10 * 60000;
    (o.fuel || []).forEach(function (f) {
      var list = tripsBy[refId(f.device)]; if (!list) return;
      var t = new Date(f.dateTime).getTime();
      for (var i = list.length - 1; i >= 0; i--) {
        if (list[i].a <= t + TOL) { if (t <= list[i].b + TOL) list[i].liters += Number(f.totalFuelUsed) || 0; break; }
      }
    });
    var out = [];
    Object.keys(tripsBy).forEach(function (dev) {
      var by = {}, totKm = 0, totL = 0;
      tripsBy[dev].forEach(function (t) {
        if (!(t.km > 0) || !(t.liters > 0) || !t.driver || t.driver === UNKNOWN_DRIVER) return;
        var d = by[t.driver] || (by[t.driver] = { km: 0, liters: 0 });
        d.km += t.km; d.liters += t.liters; totKm += t.km; totL += t.liters;
      });
      var best = null;
      Object.keys(by).forEach(function (id) {
        var d = by[id], pKm = totKm - d.km, pL = totL - d.liters;
        if (d.km < MIN_KM_DRIVER || pKm < MIN_KM_DRIVER || !(pL > 0)) return;
        var kmlD = d.km / d.liters, kmlP = pKm / pL;
        if (kmlD > (1 - DRIVER_GAP) * kmlP) return;
        var extra = d.liters - d.km / kmlP;
        var value = extra * o.price * (30 / DRIVER_DAYS);
        if (!best || value > best.value) best = { id: id, kmlD: kmlD, kmlP: kmlP, extra: extra, value: value };
      });
      if (!best || best.value < MIN_WEEK_COST * MONTH_PER_WEEK) return;
      var nm = names[best.id] || 'Motorista ' + best.id;
      out.push({
        type: 'kml_motorista', tag: 'Condução', short: 'Falar com ' + nm, vehicleId: dev, driverId: best.id,
        value: best.value, sortKey: best.value, valueLabel: 'até ' + brl.format(best.value) + '/mês',
        title: 'Conversar com ' + nm + ' sobre o consumo no ' + devName[dev],
        why: 'Nas últimas 5 semanas, ' + nm + ' fez ' + kmlFmt.format(best.kmlD) + ' km/L no ' + devName[dev] +
          '; os outros motoristas do mesmo caminhão fizeram ' + kmlFmt.format(best.kmlP) + ' km/L. A diferença foi de ' +
          liters(best.extra) + ' no período.',
        check: 'Se as rotas e as cargas de ' + nm + ' são parecidas com as dos colegas. Se forem mais pesadas, a diferença é da operação.',
        todo: 'Mostrar a comparação a ' + nm + ' e combinar marchas mais altas, menos aceleração forte e motor desligado nas paradas.'
      });
    });
    return out;
  }

  /** Junta as ações das três frentes sem contar o mesmo diesel duas vezes. */
  function combineActions(idleActions, more) {
    var all = idleActions.slice();
    more.forEach(function (a) {
      var copy = Object.assign({}, a);
      if (a.type === 'kml_motorista') {
        // o tempo parado do mesmo motorista já entra no km/L: desconta a ação de tempo parado
        idleActions.forEach(function (b) {
          if (b.vehicleId === a.vehicleId && b.type === 'conducao' && b.driverId === a.driverId) copy.value = Math.max(0, copy.value - b.value);
        });
        if (copy.value < MIN_WEEK_COST * MONTH_PER_WEEK) return;
        copy.sortKey = copy.value; copy.valueLabel = 'até ' + brl.format(copy.value) + '/mês';
      }
      all.push(copy);
    });
    all.sort(function (x, y) { return y.sortKey - x.sortKey; });
    return all;
  }

  function idleFuelFromRecords(fuel, from, to) {
    var out = {};
    (fuel || []).forEach(function (f) {
      var t = new Date(f.dateTime).getTime();
      if (!(t >= from && t < to)) return;
      var l = Number(f.totalIdlingFuelUsedL) || 0;
      if (l > 0) { var id = refId(f.device); out[id] = (out[id] || 0) + l; }
    });
    return out;
  }

  /** Todas as ações de uma semana (as 3 frentes), com os mesmos critérios em
   *  qualquer semana: é o que mantém as duas abas coerentes. */
  function weekActions(o, endMs, price, idleFallback) {
    var startMs = endMs - WEEK_MS;
    var tripsW = (o.trips || []).filter(function (t) { var a = new Date(t.start).getTime(); return a >= startMs && a < endMs; });
    var summary = summarizeIdle(o.devices || [], tripsW, o.users || []);
    var idle = idleFuelFromRecords(o.fuel, startMs, endMs);
    if (!Object.keys(idle).length && idleFallback) idle = idleFallback;
    var s = valueIdle(summary, idle, price);
    var e = evolution(Object.assign({}, o, { end: new Date(endMs), price: price, weeks: Math.max(5, o.chartWeeks || 5) }));
    var more = e.hasFuel ? fuelActions(e).concat(driverGaps(Object.assign({}, o, { end: new Date(endMs), price: price }))) : [];
    return { s: s, e: e, actions: combineActions(advise(s, DAYS).all, more) };
  }

  /** "Por onde o diesel está vazando", semana a semana, nas últimas n semanas. */
  function leakTrend(o, price, n) {
    var end = o.end.getTime(), out = [];
    for (var k = (n || TREND_WEEKS) - 1; k >= 0; k--) {
      var endK = end - k * WEEK_MS;
      var w = weekActions(o, endK, price);
      var byName = {};
      leakByPillar(w.actions).forEach(function (l) { byName[l.name] = l.value; });
      out.push({ start: new Date(endK - WEEK_MS), total: w.actions.reduce(function (x, a) { return x + a.value; }, 0), byPillar: byName });
    }
    return out;
  }

  var PILLAR_OF = { conducao: 'Condução', equipe: 'Condução', motorista: 'Condução', velocidade: 'Condução', kml_motorista: 'Condução', mecanica: 'Mecânica', operacao: 'Operação' };

  /** R$ por mês recuperável em cada frente, do maior para o menor. */
  function leakByPillar(actions) {
    var sum = { 'Mecânica': 0, 'Condução': 0, 'Operação': 0 }, count = { 'Mecânica': 0, 'Condução': 0, 'Operação': 0 };
    actions.forEach(function (a) { var p = PILLAR_OF[a.type]; if (p) { sum[p] += a.value; count[p] += 1; } });
    return Object.keys(sum).map(function (k) { return { name: k, value: sum[k], count: count[k] }; })
      .sort(function (a, b) { return b.value - a.value; });
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
  function normalReason(v, fleetRate, kml) {
    var fuelNote = kml ? ' Consumo: ' + kmlFmt.format(kml.kmlLast) + ' km/L nesta semana, dentro da própria média (' + kmlFmt.format(kml.kmlBase) + ' km/L).' : '';
    return idleReason(v, fleetRate) + fuelNote;
  }

  function idleReason(v, fleetRate) {
    if (!(v.idleS > 0)) return 'Sem tempo de motor ligado parado na semana.';
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

  /** Impacto financeiro de um veículo, para ordenar e para o número à direita.
   *  1º economia possível (R$/mês), 2º dinheiro sem motorista identificado, 3º gasto da semana. */
  function impactOf(v, actions) {
    var saving = 0, unowned = 0;
    actions.forEach(function (a) {
      if (a.type === 'dados') unowned += a.sortKey * 2; else saving += a.value;
    });
    if (saving > 0) return { tier: 3, amount: saving, label: 'até ' + brl.format(saving) + '/mês' };
    if (unowned > 0) return { tier: 2, amount: unowned, label: brl.format(unowned) + '/mês sem dono' };
    return { tier: 1, amount: v.cost, label: brl.format(v.cost) + ' na semana' };
  }

  function vehicleRow(v, rank, actions, open, fleetRate, kml) {
    var li = el('li', 'cia-row');
    var details = el('details', 'cia-row-details');
    if (open) details.open = true;
    var summary = el('summary', 'cia-row-summary');
    summary.appendChild(el('span', 'cia-row-pos', String(rank)));
    var main = el('span', 'cia-row-main');
    main.appendChild(el('span', 'cia-row-name', v.name));
    var impact = impactOf(v, actions);
    var meta = [];
    if (kml) meta.push(kmlFmt.format(kml.kmlLast) + ' km/L');
    if (v.idleS > 0) meta.push((impact.tier > 1 ? brl.format(v.cost) + ' parado · ' : '') + hours(v.idleS) + ' parado · ' + pctFmt.format(v.share) + ' do tempo de motor ligado');
    main.appendChild(el('span', 'cia-row-meta', meta.join(' · ') || 'Sem tempo parado na semana'));
    summary.appendChild(main);
    var side = el('span', 'cia-row-side');
    side.appendChild(el('span', 'cia-row-cost' + (impact.tier === 1 ? ' cia-row-cost-normal' : ''), impact.label));
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
      ok.appendChild(el('p', 'cia-row-ok', normalReason(v, fleetRate, kml)));
      body.appendChild(ok);
    }
    if (v.idleS > 0) {
      var fuelText = v.estimated ? '≈ ' + liters(v.liters) + ' estimados' : liters(v.liters) + ' medidos pelo veículo';
      body.appendChild(el('p', 'cia-row-fuel', 'Diesel queimado parado: ' + fuelText + ' (' + brl.format(v.cost) + ').'));
      body.appendChild(driversBlock(v));
    }
    details.appendChild(body);
    li.appendChild(details);
    return li;
  }

  function renderTop(s, actions, devices, kmlBy) {
    var byVehicle = {};
    actions.forEach(function (a) { (byVehicle[a.vehicleId] = byVehicle[a.vehicleId] || []).push(a); });
    // Universo: todo veículo com tempo parado ou com alguma ação
    var rowsById = {};
    s.ranking.forEach(function (v) { rowsById[v.id] = v; });
    (devices || []).forEach(function (d) {
      if (!rowsById[d.id] && byVehicle[d.id]) rowsById[d.id] = { id: d.id, name: d.name || d.id, idleS: 0, share: 0, cost: 0, liters: 0, estimated: true, drivers: [], unknownShare: 0 };
    });
    // Sempre o maior impacto financeiro primeiro
    var ordered = Object.keys(rowsById).map(function (k) { return rowsById[k]; }).sort(function (a, b) {
      var ia = impactOf(a, byVehicle[a.id] || []), ib = impactOf(b, byVehicle[b.id] || []);
      return ia.tier !== ib.tier ? ib.tier - ia.tier : ib.amount - ia.amount;
    });
    var rows = ordered.slice(0, TOP_VEHICLES);
    var list = $('cia-ranking');
    clear(list);
    var opened = false;
    rows.forEach(function (v, i) {
      var acts = byVehicle[v.id] || [];
      var open = !opened && acts.length > 0;
      if (open) opened = true;
      list.appendChild(vehicleRow(v, i + 1, acts, open, s.share, kmlBy ? kmlBy[v.id] : null));
    });
    var withAction = Object.keys(byVehicle).length;
    var value = actions.reduce(function (x, a) { return x + a.value; }, 0);
    var sub;
    if (!rows.length) sub = '';
    else if (!withAction) sub = 'Nenhuma ação urgente nas três frentes. Clique em um veículo para ver por que ele está dentro do normal.';
    else {
      sub = (withAction === 1 ? '1 veículo tem' : withAction + ' veículos têm') + ' ação sugerida';
      sub += value > 0 ? ', que pode recuperar até ' + brl.format(value) + ' por mês.' : '.';
      sub += ' Em ordem de impacto financeiro. Clique para ver o porquê.';
    }
    $('cia-ranking-title').textContent = 'Top ' + Math.min(TOP_VEHICLES, rows.length) + ' veículos';
    $('cia-ranking-sub').textContent = sub;
    return value;
  }

  function renderLeak(leaks, hasFuel) {
    var box = $('cia-leak');
    clear(box);
    var total = leaks.reduce(function (x, l) { return x + l.value; }, 0);
    var bar = $('cia-leak-bar');
    clear(bar);
    var shades = { 'Mecânica': 'var(--cia-accent)', 'Condução': 'var(--cia-navy)', 'Operação': 'var(--cia-rail)' };
    leaks.forEach(function (l) {
      if (total > 0 && l.value > 0) {
        var seg = el('div', 'cia-leak-seg');
        seg.style.width = (l.value / total * 100) + '%';
        seg.style.background = shades[l.name];
        bar.appendChild(seg);
      }
      var col = el('div', 'cia-leak-col');
      var head = el('div', 'cia-leak-head');
      var dot = el('span', 'cia-leak-dot'); dot.style.background = shades[l.name];
      head.appendChild(dot);
      head.appendChild(el('span', 'cia-leak-name', l.name));
      col.appendChild(head);
      col.appendChild(el('p', 'cia-leak-value', l.value > 0 ? 'até ' + brl.format(l.value) : brl.format(0)));
      var note;
      if (l.value > 0) note = (l.count === 1 ? '1 ação' : l.count + ' ações') + ' · ' + pctFmt.format(l.value / total) + ' do total';
      else if (l.name === 'Mecânica' && !hasFuel) note = 'Sem dados de consumo para avaliar';
      else note = 'Nada fora do normal na semana';
      col.appendChild(el('p', 'cia-leak-note', note));
      box.appendChild(col);
    });
    bar.hidden = !(total > 0);
    $('cia-leak-sub').textContent = total > 0
      ? 'Por mês, se as ações sugeridas forem feitas. Em ordem de valor.'
      : 'Nenhuma das três frentes tem ação sugerida nesta semana.';
    $('cia-leak-card').hidden = false;
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
  var compact = new Intl.NumberFormat('pt-BR', { notation: 'compact', maximumFractionDigits: 1 });
  function brlShort(v) { return v >= 10000 ? 'R$ ' + compact.format(v) : brl.format(v); }
  function sentence(t) { return t.replace(/\.\.$/, '.'); }

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
    side.appendChild(el('span', 'cia-row-hint', 'consumo por km ' + signedPct.format(m.consumption)));
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
    dl.appendChild(el('dd', 'cia-action-text', sentence(m.todo)));
    dl.appendChild(el('dt', 'cia-action-label', worse ? 'Gasto a mais' : 'Diesel a menos'));
    dl.appendChild(el('dd', 'cia-action-text', liters(Math.abs(m.extraLiters)) + ' (' + brl.format(Math.abs(m.value)) + ') nesta semana, contra a média das 4 anteriores'));
    box.appendChild(dl);
    body.appendChild(box);
    details.appendChild(body);
    li.appendChild(details);
    return li;
  }

  /** Variação contra a média das 4 semanas anteriores, sem julgar se é bom ou ruim. */
  function versusBase(change, noun) {
    if (Math.abs(change) < NOISE) return noun + ' estável em relação à média das 4 semanas anteriores';
    return noun + ' ' + pctFmt.format(Math.abs(change)) + (change > 0 ? ' maior' : ' menor') + ' que a média das 4 semanas anteriores';
  }

  /** "12% maior", "estável" ou "3,5 vezes" a média: legível mesmo quando o salto é grande. */
  function vsAvg(now, avg, what) {
    var r = now / avg;
    if (Math.abs(r - 1) < NOISE) return 'Estável em relação à ' + what;
    if (r >= 2) return new Intl.NumberFormat('pt-BR', { maximumFractionDigits: 1 }).format(r) + ' vezes a ' + what;
    return pctFmt.format(Math.abs(r - 1)) + (r > 1 ? ' maior' : ' menor') + ' que a ' + what;
  }

  var PILLAR_SHADE = { 'Mecânica': 'var(--cia-accent)', 'Condução': 'var(--cia-navy)', 'Operação': 'var(--cia-rail)' };

  function stackedChart(box, trend) {
    clear(box);
    var max = Math.max.apply(null, trend.map(function (w) { return w.total; }).concat([0]));
    trend.forEach(function (w, i) {
      var col = el('div', 'cia-bar-col' + (i === trend.length - 1 ? ' cia-bar-last' : ''));
      col.appendChild(el('span', 'cia-bar-value', w.total > 0 ? brlShort(w.total) : brl.format(0)));
      var track = el('div', 'cia-bar-track');
      var stack = el('div', 'cia-stack');
      stack.style.height = (max > 0 ? Math.max(2, Math.round(w.total / max * 100)) : 2) + '%';
      ['Operação', 'Condução', 'Mecânica'].forEach(function (k) {
        var v = w.byPillar[k] || 0;
        if (!(v > 0) || !(w.total > 0)) return;
        var seg = el('div', 'cia-stack-seg');
        seg.style.height = (v / w.total * 100) + '%';
        seg.style.background = PILLAR_SHADE[k];
        seg.title = k + ': ' + brl.format(v);
        stack.appendChild(seg);
      });
      track.appendChild(stack);
      col.appendChild(track);
      col.appendChild(el('span', 'cia-bar-label', dayMonth.format(w.start)));
      box.appendChild(col);
    });
  }

  /** Visão ampla e de tendência, com os mesmos critérios da aba Esta semana. */
  function renderEvolution(e, trend) {
    var status = $('cia-evo-status');
    if (!e.hasFuel) {
      status.textContent = 'Esta frota não envia o consumo de combustível à Geotab, então não dá para mostrar gasto nem consumo por km.';
      $('cia-evo-content').hidden = true;
      return;
    }
    status.textContent = '';
    var lastT = trend[trend.length - 1], prior = trend.slice(0, -1);
    var priorAvg = prior.length ? prior.reduce(function (x, w) { return x + w.total; }, 0) / prior.length : 0;

    $('cia-evo-cost').textContent = brl.format(e.cost);
    $('cia-evo-cost-note').textContent = liters(e.liters) + ' · ' + versusBase(e.costChange, 'gasto').replace(/^gasto /, '');
    $('cia-evo-rec').textContent = brl.format(lastT.total);
    $('cia-evo-rec-note').textContent = priorAvg > 0
      ? vsAvg(lastT.total, priorAvg, 'média das ' + prior.length + ' semanas anteriores')
      : 'por mês, nas ações da última semana';
    $('cia-evo-kml').textContent = e.kml > 0 ? kmlFmt.format(e.kml) + ' km/L' : '—';
    // km/L maior = consumo por km menor
    $('cia-evo-kml-note').textContent = e.baseKml > 0 && e.kml > 0 ? versusBase(e.baseKml / e.kml - 1, 'Consumo por km') : '';
    $('cia-evo-moves').textContent = e.worseCount + (e.worseCount === 1 ? ' aumentou' : ' aumentaram');
    $('cia-evo-moves-note').textContent = e.betterCount + (e.betterCount === 1 ? ' reduziu' : ' reduziram') +
      '; os demais variaram menos de 3% contra a própria média';

    stackedChart($('cia-chart-leak'), trend);
    var legend = $('cia-leak-trend');
    clear(legend);
    ['Mecânica', 'Condução', 'Operação'].map(function (k) {
      var now = lastT.byPillar[k] || 0;
      var avg = prior.length ? prior.reduce(function (x, w) { return x + (w.byPillar[k] || 0); }, 0) / prior.length : 0;
      return { name: k, now: now, avg: avg };
    }).sort(function (a, b) { return b.now - a.now; }).forEach(function (l) {
      var col = el('div', 'cia-leak-col');
      var head = el('div', 'cia-leak-head');
      var dot = el('span', 'cia-leak-dot'); dot.style.background = PILLAR_SHADE[l.name];
      head.appendChild(dot); head.appendChild(el('span', 'cia-leak-name', l.name));
      col.appendChild(head);
      col.appendChild(el('p', 'cia-leak-value', l.now > 0 ? 'até ' + brl.format(l.now) : brl.format(0)));
      var note;
      if (!(l.avg > 0) && !(l.now > 0)) note = 'Nada fora do normal nas últimas semanas';
      else if (!(l.avg > 0)) note = 'Novo na última semana';
      else note = vsAvg(l.now, l.avg, 'média das semanas anteriores');
      col.appendChild(el('p', 'cia-leak-note', note));
      legend.appendChild(col);
    });

    var labels = e.weeks.map(function (w) { return dayMonth.format(w.start); });
    barChart($('cia-chart-cost'), e.weeks.map(function (w) { return w.cost; }), labels, brlShort);
    barChart($('cia-chart-kml'), e.weeks.map(function (w) { return w.kml; }), labels, function (v) { return kmlFmt.format(v); });

    $('cia-evo-content').hidden = false;
  }

  function showEvolution() {
    if (!evo.data) return;
    var price = readPrice();
    renderEvolution(evolution(Object.assign({}, evo.data, { price: price, weeks: WEEKS })), leakTrend(evo.data, price, TREND_WEEKS));
  }

  var evo = { api: null, state: null, loaded: false, loading: false, data: null };

  // A carga de dados é uma só (load); a aba Evolução só mostra o que já chegou.
  function loadEvolution() {
    if (evo.data) { showEvolution(); return; }
    $('cia-evo-status').textContent = 'Carregando ' + LOAD_WEEKS + ' semanas de dados…';
  }

  function showTab(name) {
    ['semana', 'evolucao'].forEach(function (t) {
      var on = t === name;
      $('cia-tab-' + t).setAttribute('aria-selected', on ? 'true' : 'false');
      $('cia-tab-' + t).tabIndex = on ? 0 : -1;
      $('cia-panel-' + t).hidden = !on;
    });
    $('cia-title').textContent = name === 'semana' ? 'Sua frota na última semana' : 'Evolução da frota';
    $('cia-period').hidden = name !== 'semana';
    if (name === 'evolucao') loadEvolution();
  }

  function bindTabs() {
    var bar = $('cia-tabs');
    if (!bar || bar.getAttribute('data-bound')) return;
    bar.setAttribute('data-bound', '1');
    $('cia-tab-semana').addEventListener('click', function () { showTab('semana'); });
    $('cia-tab-evolucao').addEventListener('click', function () { showTab('evolucao'); });
    var go = $('cia-go-semana');
    if (go) go.addEventListener('click', function () { showTab('semana'); $('cia-tab-semana').focus(); });
    bar.addEventListener('keydown', function (ev) {
      if (ev.key !== 'ArrowRight' && ev.key !== 'ArrowLeft') return;
      var next = $('cia-tab-semana').getAttribute('aria-selected') === 'true' ? 'evolucao' : 'semana';
      showTab(next); $('cia-tab-' + next).focus();
    });
  }

  var last = null; // último resultado, para recalcular quando o preço muda

  function render() {
    if ($('cia-version')) $('cia-version').textContent = 'Combustível IA · versão ' + VERSION;
    if (!last) return;
    var price = readPrice();
    var week = weekActions(evo.data, last.end.getTime(), price, last.fuel);
    var s = week.s, actions = week.actions;
    var e = evolution(Object.assign({}, evo.data, { price: price, weeks: WEEKS }));

    var dateFmt = new Intl.DateTimeFormat('pt-BR', { day: '2-digit', month: '2-digit' });
    $('cia-period').textContent = 'De ' + dateFmt.format(last.start) + ' a ' + dateFmt.format(new Date(last.end.getTime() - 1)) + ' · últimos 7 dias fechados';
    var value = actions.reduce(function (x, a) { return x + a.value; }, 0);
    var nActions = actions.filter(function (a) { return a.value > 0; }).length;
    $('cia-kpi-value').textContent = brl.format(value);
    $('cia-kpi-value-note').textContent = nActions ? 'em ' + (nActions === 1 ? '1 ação sugerida' : nActions + ' ações sugeridas') : 'nenhuma ação sugerida';
    $('cia-kpi-cost').textContent = brl.format(s.totalCost);
    $('cia-kpi-share').textContent = 'Na semana · ' + hours(s.totalIdleS) + ' · ' + pctFmt.format(s.share) + ' do tempo de motor ligado';
    var acting = {};
    actions.forEach(function (a) { acting[a.vehicleId] = true; });
    $('cia-kpi-acting').textContent = String(Object.keys(acting).length);
    $('cia-kpi-vehicles').textContent = 'de ' + (last.devices || []).length + ' veículos analisados';
    $('cia-kpis').hidden = false;

    renderLeak(leakByPillar(actions), !!(e && e.hasFuel));

    var n = s.ranking.length;
    var basis;
    if (n === 0) basis = '';
    else if (s.measuredVehicles === n) basis = 'Diesel parado medido pelos próprios veículos.';
    else if (s.measuredVehicles === 0) basis = 'Diesel parado estimado em ' + IDLE_RATE_LPH + ' L por hora, média de caminhão pesado: estes veículos não informam o combustível gasto parado.';
    else basis = 'Diesel parado medido em ' + s.measuredVehicles + ' de ' + n + ' veículos; nos demais, estimado em ' + IDLE_RATE_LPH + ' L por hora.';
    if (e && e.hasFuel) basis += ' Consumo de cada veículo comparado com a própria média das 4 semanas anteriores.';
    $('cia-basis').textContent = basis;

    var input = $('cia-price');
    if (document.activeElement !== input) input.value = priceFmt.format(price);
    $('cia-price-source').textContent = price === DEFAULT_PRICE ? PRICE_SOURCE : 'valor informado por você';

    renderTop(s, actions, last.devices, e ? e.byVehicle : null);
    $('cia-ranking-card').hidden = false;
    setStatus('');
    if (!$('cia-panel-evolucao').hidden) showEvolution();
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
    // Sempre os últimos 7 dias fechados: termina à meia-noite de hoje (hoje não entra)
    var end = new Date(); end.setHours(0, 0, 0, 0);
    var start = new Date(end.getTime() - DAYS * 86400000);
    var evoStart = new Date(end.getTime() - LOAD_WEEKS * WEEK_MS);
    var from = start.toISOString(), to = end.toISOString();
    var range6 = { fromDate: evoStart.toISOString(), toDate: to };
    var groups = state && state.getGroupFilter ? state.getGroupFilter() : [];
    setStatus('Carregando ' + LOAD_WEEKS + ' semanas de dados da frota…');
    bindPrice();
    bindTabs();
    evo = { api: api, state: state, loaded: false, loading: true, data: null };
    var deviceSearch = { fromDate: evoStart.toISOString() };
    if (groups && groups.length) deviceSearch.groups = groups;
    function signal(id) { return { diagnosticSearch: { id: id }, fromDate: from, toDate: to }; }
    // FuelAndEnergyUsed é o tipo novo; FuelUsed, o anterior. Usa o que a base tiver.
    var fuelP = safe(call(api, 'Get', { typeName: 'FuelAndEnergyUsed', search: range6 })).then(function (r) {
      return r && r.length ? r : safe(call(api, 'Get', { typeName: 'FuelUsed', search: range6 }));
    });

    return Promise.all([
      call(api, 'Get', { typeName: 'Device', search: deviceSearch }),
      call(api, 'Get', { typeName: 'Trip', search: range6 }),
      call(api, 'Get', { typeName: 'User', search: { isDriver: true } }),
      safe(call(api, 'Get', { typeName: 'StatusData', search: signal(DIAG.idleFuel) })),
      fuelP,
      safe(call(api, 'Get', { typeName: 'FaultData', search: { fromDate: new Date(end.getTime() - TREND_WEEKS * WEEK_MS).toISOString(), toDate: to } }))
    ]).then(function (res) {
      var devices = res[0], trips6 = res[1], users = res[2];
      var inFilter = {};
      devices.forEach(function (d) { inFilter[d.id] = true; });
      var t0 = start.getTime();
      var trips = trips6.filter(function (t) { return new Date(t.start).getTime() >= t0; });
      var fuelAll = idleFuelByDevice(res[3] || []);
      var fuel = {};
      Object.keys(fuelAll).forEach(function (id) { if (inFilter[id]) fuel[id] = fuelAll[id]; });
      var tripsIn = trips.filter(function (t) { return inFilter[refId(t.device)]; });

      evo.data = { devices: devices, trips: trips6, users: users, names: driverNames(users), fuel: res[4] || [], faults: res[5] || [], end: end };
      evo.loaded = true; evo.loading = false;
      last = { summary: summarizeIdle(devices, trips, users), fuel: fuel, start: start, end: end, devices: devices };
      render();

      return Promise.all([
        exists(api, 'StatusData', signal(DIAG.rpm)),
        exists(api, 'StatusData', signal(DIAG.totalFuel)),
        exists(api, 'StatusData', signal(DIAG.coolant)),
        exists(api, 'Zone')
      ]).then(function (c) {
        renderCheckup(assessData({
          devices: devices, trips: tripsIn, idleFuelVehicles: Object.keys(fuel).length,
          rpm: c[0], totalFuel: c[1] || (res[4] || []).length > 0, coolant: c[2],
          faults: (res[5] || []).some(function (f) { return new Date(f.dateTime).getTime() >= t0; }), zones: c[3]
        }));
      });
    }).catch(function (err) {
      evo.loading = false;
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
      valueIdle: valueIdle, advise: advise, evolution: evolution, fuelActions: fuelActions, driverGaps: driverGaps,
      combineActions: combineActions, leakByPillar: leakByPillar, weekActions: weekActions, leakTrend: leakTrend, assessData: assessData, parsePrice: parsePrice
    };
  }
})();
