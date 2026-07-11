/* ==========================================================================
   Primeros Productos Pehuenia S.A. — Panel de Fleteros
   app.js  ·  Lógica de la herramienta (vanilla JS, sin dependencias)
   ========================================================================== */
(function () {
  "use strict";

  var CONFIG = window.__PPP_CONFIG__ || {};
  var UMBRAL = CONFIG.umbrales || { bueno: 90, medio: 75 };
  var DIAS = CONFIG.diasHistorial || 14;

  var MESES = ["ene", "feb", "mar", "abr", "may", "jun",
               "jul", "ago", "sep", "oct", "nov", "dic"];

  // ---- Utilidades seguras: un init que falla no rompe el resto ----------
  function safe(fn, name) {
    try { fn(); } catch (e) { console.error("[PPP] Error en " + name + ":", e); }
  }
  function $(sel, ctx) { return (ctx || document).querySelector(sel); }
  function el(tag, cls, html) {
    var n = document.createElement(tag);
    if (cls) n.className = cls;
    if (html != null) n.innerHTML = html;
    return n;
  }
  function pct(x) { return x == null ? null : Math.round(x * 1000) / 10; } // 1 decimal
  function claseColor(p) {
    if (p == null) return "n";
    if (p >= UMBRAL.bueno) return "ok";
    if (p >= UMBRAL.medio) return "mid";
    return "low";
  }
  function fmtFecha(iso) {
    var p = iso.split("-");
    if (p.length !== 3) return iso;
    return parseInt(p[2], 10) + " " + MESES[parseInt(p[1], 10) - 1];
  }

  // ---- Parser CSV robusto (comillas, comas internas, saltos) ------------
  function parseCSV(text) {
    var rows = [], row = [], cur = "", inQ = false;
    for (var i = 0; i < text.length; i++) {
      var c = text[i], n = text[i + 1];
      if (inQ) {
        if (c === '"' && n === '"') { cur += '"'; i++; }
        else if (c === '"') { inQ = false; }
        else { cur += c; }
      } else {
        if (c === '"') { inQ = true; }
        else if (c === ",") { row.push(cur); cur = ""; }
        else if (c === "\n") { row.push(cur); rows.push(row); row = []; cur = ""; }
        else if (c === "\r") { /* ignorar */ }
        else { cur += c; }
      }
    }
    if (cur.length || row.length) { row.push(cur); rows.push(row); }
    return rows.filter(function (r) { return r.some(function (v) { return String(v).trim() !== ""; }); });
  }

  function norm(s) {
    return String(s || "").trim().toLowerCase()
      .normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/\s+/g, "_");
  }

  // Normaliza fechas a ISO (AAAA-MM-DD). Acepta AAAA-MM-DD y DD/MM/AAAA.
  function normFecha(s) {
    s = String(s || "").trim();
    var iso = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
    if (iso) return iso[1] + "-" + ("0" + iso[2]).slice(-2) + "-" + ("0" + iso[3]).slice(-2);
    var dmy = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/);
    if (dmy) {
      var y = dmy[3].length === 2 ? "20" + dmy[3] : dmy[3];
      return y + "-" + ("0" + dmy[2]).slice(-2) + "-" + ("0" + dmy[1]).slice(-2);
    }
    return s;
  }

  // Convierte filas CSV en registros, aceptando variantes de nombres de columna.
  function registrosDesdeCSV(text) {
    var rows = parseCSV(text);
    if (!rows.length) return [];
    var head = rows[0].map(norm);
    function idx() {
      for (var a = 0; a < arguments.length; a++) {
        var k = head.indexOf(arguments[a]);
        if (k !== -1) return k;
      }
      return -1;
    }
    var iF = idx("fecha", "dia", "fecha_entrega");
    var iN = idx("fletero", "nombre", "chofer", "transportista", "repartidor");
    var iZ = idx("zona", "localidad", "ciudad", "region");
    var iAsig = idx("entregas_asignadas", "asignadas", "entregas_totales", "total_entregas");
    var iReal = idx("entregas_realizadas", "realizadas", "entregas_ok", "entregadas");
    var iCar = idx("cartones_a_retornar", "cartones_totales", "cartones", "a_retornar");
    var iRet = idx("cartones_retornados", "retornados", "cartones_ok");
    var iEfE = idx("ef_entrega_pct", "efectividad_entrega", "ef_entrega");
    var iEfR = idx("ef_retorno_pct", "efectividad_retorno", "ef_retorno");

    var out = [];
    for (var r = 1; r < rows.length; r++) {
      var row = rows[r];
      var reg = {
        fecha: iF > -1 ? normFecha(row[iF]) : "",
        fletero: iN > -1 ? String(row[iN]).trim() : "",
        zona: iZ > -1 ? String(row[iZ]).trim() : ""
      };
      function num(k) {
        if (k < 0) return null;
        var v = String(row[k]).replace("%", "").replace(",", ".").trim();
        if (v === "") return null;
        var f = parseFloat(v);
        return isNaN(f) ? null : f;
      }
      reg.entregas_asignadas = num(iAsig);
      reg.entregas_realizadas = num(iReal);
      reg.cartones_a_retornar = num(iCar);
      reg.cartones_retornados = num(iRet);
      // Porcentajes directos si vinieran (los normalizamos a fracción 0..1)
      var pe = num(iEfE), pr = num(iEfR);
      reg._efE = pe == null ? null : (pe > 1 ? pe / 100 : pe);
      reg._efR = pr == null ? null : (pr > 1 ? pr / 100 : pr);
      if (reg.fletero && reg.fecha) out.push(reg);
    }
    return out;
  }

  // Calcula efectividades por registro (fracción 0..1 o null).
  function conEfectividad(reg) {
    var efE = reg._efE;
    if (efE == null && reg.entregas_asignadas > 0)
      efE = reg.entregas_realizadas / reg.entregas_asignadas;
    var efR = reg._efR;
    if (efR == null && reg.cartones_a_retornar > 0)
      efR = reg.cartones_retornados / reg.cartones_a_retornar;
    return { efE: (efE == null ? null : efE), efR: (efR == null ? null : efR) };
  }

  // ---- Agregaciones -----------------------------------------------------
  function agrupaPorFletero(registros) {
    var map = {};
    registros.forEach(function (r) {
      var k = r.fletero;
      if (!map[k]) map[k] = { nombre: k, zona: r.zona || "", regs: [] };
      if (r.zona && !map[k].zona) map[k].zona = r.zona;
      map[k].regs.push(r);
    });
    Object.keys(map).forEach(function (k) {
      map[k].regs.sort(function (a, b) { return a.fecha < b.fecha ? -1 : 1; });
    });
    return map;
  }

  // Promedio ponderado por volumen sobre el período (más justo que promediar %).
  function promedioPeriodo(regs) {
    var sa = 0, sr = 0, sc = 0, sct = 0, nE = 0, nR = 0, accE = 0, accR = 0;
    regs.forEach(function (r) {
      if (r.entregas_asignadas > 0) { sa += r.entregas_asignadas; sr += r.entregas_realizadas || 0; }
      if (r.cartones_a_retornar > 0) { sc += r.cartones_a_retornar; sct += r.cartones_retornados || 0; }
      var e = conEfectividad(r);
      if (e.efE != null && !(r.entregas_asignadas > 0)) { accE += e.efE; nE++; }
      if (e.efR != null && !(r.cartones_a_retornar > 0)) { accR += e.efR; nR++; }
    });
    var efE = sa > 0 ? sr / sa : (nE ? accE / nE : null);
    var efR = sc > 0 ? sct / sc : (nR ? accR / nR : null);
    return { efE: efE, efR: efR, entregas: sr, asignadas: sa, cartones: sct, cartonesTot: sc };
  }

  function ultimoRegistro(regs) { return regs.length ? regs[regs.length - 1] : null; }

  function fechasUnicas(regs) {
    var s = {};
    regs.forEach(function (r) { s[r.fecha] = 1; });
    return Object.keys(s).sort();
  }
  // Filtra los registros a las últimas N fechas con datos (a nivel empresa).
  function soloUltimasFechas(regs, todasLasFechas, n) {
    var keep = {};
    todasLasFechas.slice(-n).forEach(function (f) { keep[f] = 1; });
    return regs.filter(function (r) { return keep[r.fecha]; });
  }
  var NOMBRES_MES = ["enero", "febrero", "marzo", "abril", "mayo", "junio",
    "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"];

  // ---- Componentes visuales --------------------------------------------
  function anillo(pValue, etiqueta, sub) {
    var p = pValue == null ? 0 : Math.max(0, Math.min(100, pValue));
    var cls = claseColor(pValue);
    var R = 52, C = 2 * Math.PI * R;
    var wrap = el("div", "ring ring--" + cls);
    wrap.innerHTML =
      '<svg viewBox="0 0 130 130" class="ring__svg" aria-hidden="true">' +
        '<circle class="ring__track" cx="65" cy="65" r="' + R + '"></circle>' +
        '<circle class="ring__val" cx="65" cy="65" r="' + R + '" ' +
          'stroke-dasharray="' + C.toFixed(1) + '" stroke-dashoffset="' + C.toFixed(1) + '"></circle>' +
      '</svg>' +
      '<div class="ring__center">' +
        '<span class="ring__num" data-count="' + (pValue == null ? -1 : pValue) + '">' +
          (pValue == null ? "—" : "0") + '</span>' +
        '<span class="ring__pct">' + (pValue == null ? "" : "%") + '</span>' +
      '</div>';
    var block = el("div", "metric");
    block.appendChild(wrap);
    block.appendChild(el("div", "metric__label", etiqueta + (sub ? '<span class="metric__sub">' + sub + '</span>' : "")));
    // guardamos datos para animar al entrar en viewport
    wrap._ring = { C: C, p: p, valEl: $(".ring__val", wrap), numEl: $(".ring__num", wrap), raw: pValue };
    return { block: block, wrap: wrap };
  }

  function animaAnillo(wrap) {
    var d = wrap._ring;
    if (!d || wrap._done) return;
    wrap._done = true;
    var offset = d.C * (1 - d.p / 100);
    requestAnimationFrame(function () {
      d.valEl.style.strokeDashoffset = offset.toFixed(1);
    });
    if (d.raw == null) return;
    var start = null, dur = 950;
    function step(t) {
      if (start == null) start = t;
      var k = Math.min(1, (t - start) / dur);
      var eased = 1 - Math.pow(1 - k, 3);
      d.numEl.textContent = (d.raw * eased).toFixed(1).replace(".0", "");
      if (k < 1) requestAnimationFrame(step);
      else d.numEl.textContent = (Math.round(d.raw * 10) / 10).toString().replace(/\.0$/, "");
    }
    requestAnimationFrame(step);
  }

  // Mini gráfico de barras de los últimos DIAS.
  function miniBarras(regs, cual, titulo) {
    var box = el("div", "spark");
    box.appendChild(el("div", "spark__title", titulo));
    var chart = el("div", "spark__chart");
    var recientes = regs.slice(-DIAS);
    if (!recientes.length) {
      chart.appendChild(el("div", "spark__empty", "Sin datos"));
    } else {
      recientes.forEach(function (r) {
        var e = conEfectividad(r);
        var v = cual === "E" ? e.efE : e.efR;
        var p = v == null ? 0 : Math.round(v * 100);
        var col = el("div", "spark__col");
        var bar = el("div", "spark__bar spark__bar--" + claseColor(v == null ? null : p));
        bar.style.height = "2px";
        bar.setAttribute("data-h", Math.max(4, p));
        bar.title = fmtFecha(r.fecha) + " · " + (v == null ? "—" : p + "%");
        col.appendChild(bar);
        chart.appendChild(col);
      });
    }
    box.appendChild(chart);
    return box;
  }

  function animaBarras(scope) {
    Array.prototype.forEach.call(scope.querySelectorAll(".spark__bar"), function (b, i) {
      var h = parseFloat(b.getAttribute("data-h")) || 4;
      setTimeout(function () { b.style.height = h + "%"; }, 40 + i * 22);
    });
  }

  function chip(p) {
    var c = claseColor(p);
    var t = p == null ? "—" : (Math.round(p * 10) / 10).toString().replace(/\.0$/, "") + "%";
    return '<span class="chip chip--' + c + '">' + t + "</span>";
  }

  // ---- Vistas -----------------------------------------------------------
  function vistaFletero(datos, nombre) {
    var g = datos.porFletero[nombre];
    var cont = el("div", "view");
    if (!g) { cont.appendChild(el("p", "muted", "Sin datos para este fletero todavía.")); return cont; }

    var ult = ultimoRegistro(g.regs);
    var fechasTodas = fechasUnicas(datos.registros);
    // Mes en curso = mes de la última fecha con datos de la empresa.
    var mesPrefijo = fechasTodas.length ? fechasTodas[fechasTodas.length - 1].slice(0, 7) : "";
    var mesNombre = mesPrefijo ? NOMBRES_MES[parseInt(mesPrefijo.slice(5), 10) - 1] : "mes";
    var promMes = promedioPeriodo(g.regs.filter(function (r) { return r.fecha.indexOf(mesPrefijo) === 0; }));
    var prom = promedioPeriodo(soloUltimasFechas(g.regs, fechasTodas, DIAS));

    // Encabezado del fletero
    var head = el("div", "person");
    head.innerHTML =
      '<div class="person__id"><span class="person__avatar">' +
        (nombre.trim().charAt(0).toUpperCase() || "?") + '</span>' +
        '<div><h2 class="person__name">' + nombre + '</h2>' +
        '<p class="person__meta">' + (g.zona ? g.zona + " · " : "") +
        'Último parte: ' + fmtFecha(ult.fecha) + '</p></div></div>';
    cont.appendChild(head);

    // Anillos con el total del mes en curso
    var grid = el("div", "metrics reveal");
    var aE = anillo(pct(promMes.efE), "Efectividad de entrega", "total " + mesNombre);
    var aR = anillo(pct(promMes.efR), "Retorno de cartón", "total " + mesNombre);
    grid.appendChild(aE.block);
    grid.appendChild(aR.block);
    cont.appendChild(grid);
    cont._rings = [aE.wrap, aR.wrap];

    // Promedios y contadores
    var proms = el("div", "avgs reveal");
    proms.innerHTML =
      '<div class="avg"><span class="avg__k">Promedio ' + DIAS + ' días · Entrega</span>' + chip(pct(prom.efE)) + '</div>' +
      '<div class="avg"><span class="avg__k">Promedio ' + DIAS + ' días · Cartón</span>' + chip(pct(prom.efR)) + '</div>' +
      '<div class="avg"><span class="avg__k">Entregas de ' + mesNombre + '</span><b>' + promMes.entregas + ' / ' + promMes.asignadas + '</b></div>' +
      '<div class="avg"><span class="avg__k">Cartones de ' + mesNombre + '</span><b>' + promMes.cartones + ' / ' + promMes.cartonesTot + '</b></div>';
    cont.appendChild(proms);

    // Estadísticas del mes: rechazos totales/parciales, clientes y boletas
    var st = ((window.__PPP_DATA__ && window.__PPP_DATA__.estadisticasFletero) || {})[nombre];
    if (st) {
      var proms2 = el("div", "avgs reveal");
      proms2.innerHTML =
        '<div class="avg"><span class="avg__k">Rechazos totales <small>(cliente completo)</small></span><b class="rojo">' + st.recTot + '</b></div>' +
        '<div class="avg"><span class="avg__k">Rechazos parciales <small>(alguna boleta)</small></span><b class="ambar">' + st.recPar + '</b></div>' +
        '<div class="avg"><span class="avg__k">Clientes entregados · ' + mesNombre + '</span><b>' + st.cliEnt + ' / ' + st.cliSac + '</b></div>' +
        '<div class="avg"><span class="avg__k">Boletas entregadas · ' + mesNombre + '</span><b>' + st.compEnt + ' / ' + st.compSac + '</b></div>';
      cont.appendChild(proms2);
    }

    // Mini gráficos
    var sparks = el("div", "sparks reveal");
    sparks.appendChild(miniBarras(g.regs, "E", "Entrega · últimos días"));
    sparks.appendChild(miniBarras(g.regs, "R", "Cartón · últimos días"));
    cont.appendChild(sparks);

    // Motivos de rechazo de este fletero (% sobre sus propios rechazos)
    var mpf = (window.__PPP_DATA__ && window.__PPP_DATA__.motivosPorFletero) || {};
    var mios = mpf[nombre] || [];
    if (mios.length) {
      var total = 0;
      mios.forEach(function (m) { total += m.cantidad; });
      var max = mios[0].cantidad || 1;
      var card = el("div", "chart reveal");
      var rows = mios.map(function (m) {
        var p = Math.round(100 * m.cantidad / total);
        var w = Math.max(4, Math.round(100 * m.cantidad / max));
        return '<div class="chart__row" title="' + m.motivo.replace(/"/g, "&quot;") + ' · ' + m.cantidad + ' rechazos (' + p + '%)">' +
          '<div class="chart__top"><span class="chart__label">' + m.motivo + '</span>' +
          '<b class="chart__val">' + p + '% <span class="chart__cnt">(' + m.cantidad + ')</span></b></div>' +
          '<i class="chart__track"><i class="rank__fill rank__fill--low" style="width:2%" data-w="' + w + '"></i></i>' +
        '</div>';
      }).join("");
      card.innerHTML = '<h2 class="chart__title">📋 Sus motivos de rechazo · ' + total + ' en total</h2>' + rows;
      cont.appendChild(card);
    }

    return cont;
  }

  function vistaGeneral(datos) {
    var cont = el("div", "view");
    var nombres = Object.keys(datos.porFletero);

    // Totales de la empresa del mes en curso
    var fechasTodas = fechasUnicas(datos.registros);
    var mesPrefijo = fechasTodas.length ? fechasTodas[fechasTodas.length - 1].slice(0, 7) : "";
    var mesNombre = mesPrefijo ? NOMBRES_MES[parseInt(mesPrefijo.slice(5), 10) - 1] : "mes";
    function delMes(regs) {
      return regs.filter(function (r) { return r.fecha.indexOf(mesPrefijo) === 0; });
    }
    var todos = [];
    nombres.forEach(function (n) { todos = todos.concat(datos.porFletero[n].regs); });
    var promEmp = promedioPeriodo(delMes(todos));

    var res = el("div", "metrics reveal");
    var aE = anillo(pct(promEmp.efE), "Entrega · empresa", "total " + mesNombre);
    var aR = anillo(pct(promEmp.efR), "Cartón · empresa", "total " + mesNombre);
    res.appendChild(aE.block);
    res.appendChild(aR.block);
    cont.appendChild(res);
    cont._rings = [aE.wrap, aR.wrap];

    // Datos base para los rankings (total del mes en curso)
    var filas = nombres.map(function (n) {
      var p = promedioPeriodo(delMes(datos.porFletero[n].regs));
      return { nombre: n, zona: datos.porFletero[n].zona, efE: pct(p.efE), efR: pct(p.efR) };
    });

    // Premios según la métrica y el porcentaje (reglas de la empresa).
    function premioPara(campo, v) {
      if (v == null) return null;
      if (campo === "efE") {           // Efectividad de entrega
        if (v >= 95) return 100000;
        if (v >= 90) return 50000;
      } else {                          // Retorno de cartón
        if (v >= 80) return 150000;
        if (v >= 70) return 100000;
        if (v >= 60) return 50000;
      }
      return null;
    }
    function fmtPremio(n) {
      var s = String(n).replace(/\B(?=(\d{3})+(?!\d))/g, ".");
      return "$" + s;
    }

    // Un ranking por métrica: cada uno ordena y muestra solo su valor.
    function tablaRanking(titulo, campo, etiqueta) {
      var lista = filas.filter(function (f) { return f[campo] != null; })
        .sort(function (a, b) { return (b[campo] || 0) - (a[campo] || 0); });
      if (!lista.length) return null;

      var tabla = el("div", "rank reveal");
      var head =
        '<div class="rank__head"><span>#</span><span>Fletero</span>' +
        '<span class="rank__num">Premio</span>' +
        '<span class="rank__num">' + etiqueta + '</span></div>';
      var body = lista.map(function (f, i) {
        var medal = i === 0 ? "🥇" : i === 1 ? "🥈" : i === 2 ? "🥉" : (i + 1);
        var bar = Math.max(4, Math.min(100, f[campo] || 0));
        var premio = premioPara(campo, f[campo]);
        var premioHTML = premio == null
          ? '<span class="rank__num rank__prize rank__prize--none">—</span>'
          : '<span class="rank__num rank__prize">💰 ' + fmtPremio(premio) + '</span>';
        return '<button class="rank__row" data-fletero="' + f.nombre.replace(/"/g, "&quot;") + '">' +
          '<span class="rank__pos">' + medal + '</span>' +
          '<span class="rank__name"><b>' + f.nombre + '</b>' +
            (f.zona ? '<em>' + f.zona + '</em>' : '') +
            '<i class="rank__track"><i class="rank__fill rank__fill--' + claseColor(f[campo]) + '" style="width:2%" data-w="' + bar + '"></i></i>' +
          '</span>' +
          premioHTML +
          '<span class="rank__num">' + chip(f[campo]) + '</span>' +
        '</button>';
      }).join("");
      tabla.innerHTML =
        '<h2 class="rank__title">' + titulo + '</h2>' +
        '<div class="rank__grid rank__grid--simple">' + head + body + '</div>' +
        '<p class="rank__hint">Tocá un fletero para ver su detalle.</p>';
      return tabla;
    }

    // Gráfico de barras horizontales para los top 5 (una sola serie).
    function graficoBarras(titulo, items, unidad) {
      if (!items.length) return null;
      var max = items[0].cantidad || 1;
      var card = el("div", "chart reveal");
      var rows = items.map(function (it) {
        var w = Math.max(4, Math.round(100 * it.cantidad / max));
        return '<div class="chart__row" title="' + it.etiqueta.replace(/"/g, "&quot;") + ' · ' + it.cantidad + ' ' + unidad + '">' +
          '<div class="chart__top"><span class="chart__label">' + it.etiqueta + '</span>' +
          '<b class="chart__val">' + it.cantidad + '</b></div>' +
          '<i class="chart__track"><i class="rank__fill rank__fill--low" style="width:2%" data-w="' + w + '"></i></i>' +
        '</div>';
      }).join("");
      card.innerHTML = '<h2 class="chart__title">' + titulo + '</h2>' + rows;
      return card;
    }

    // Datos: top 5 motivos (los genera el robot) y top 5 fleteros con más rechazos.
    var motivosData = (window.__PPP_DATA__ && window.__PPP_DATA__.motivos) || [];
    var topMotivos = motivosData.slice(0, 5).map(function (m) {
      return { etiqueta: m.motivo, cantidad: m.cantidad };
    });
    // Top 5 por rechazos TOTALES de cliente (no recibió ninguna de sus boletas)
    var statsData = (window.__PPP_DATA__ && window.__PPP_DATA__.estadisticasFletero) || {};
    var rechazosPorFletero = Object.keys(statsData).map(function (n) {
      return { etiqueta: n, cantidad: statsData[n].recTot || 0 };
    }).filter(function (f) { return f.cantidad > 0; })
      .sort(function (a, b) { return b.cantidad - a.cantidad; })
      .slice(0, 5);

    // Tarjetas gráficas de rechazos, ARRIBA de los rankings.
    var gMot = graficoBarras("📋 Motivos de rechazo más comunes", topMotivos, "rechazos");
    var gRech = graficoBarras("⚠️ Rechazos totales de cliente · " + mesNombre, rechazosPorFletero, "clientes");
    if (gMot || gRech) {
      var fila = el("div", "charts");
      if (gMot) fila.appendChild(gMot);
      if (gRech) fila.appendChild(gRech);
      cont.appendChild(fila);
    }

    var rankE = tablaRanking("🚚 Ranking · Efectividad de entrega · total " + mesNombre, "efE", "Entrega");
    var rankR = tablaRanking("📦 Ranking · Retorno de cartón · total " + mesNombre, "efR", "Cartón");
    if (rankE) cont.appendChild(rankE);
    if (rankR) cont.appendChild(rankR);

    return cont;
  }

  // ---- Reveal + animaciones al entrar en viewport -----------------------
  function activarReveal(scope) {
    var els = scope.querySelectorAll(".reveal");
    if (!("IntersectionObserver" in window)) {
      Array.prototype.forEach.call(els, function (e) { e.classList.add("in"); });
      return;
    }
    var io = new IntersectionObserver(function (entries) {
      entries.forEach(function (en) {
        if (en.isIntersecting) { en.target.classList.add("in"); io.unobserve(en.target); }
      });
    }, { threshold: 0.05 });
    Array.prototype.forEach.call(els, function (e) { io.observe(e); });
    // Red de seguridad: si algo quedó oculto, revelarlo.
    setTimeout(function () {
      Array.prototype.forEach.call(els, function (e) { e.classList.add("in"); });
    }, 6000);
  }

  // ---- Render principal -------------------------------------------------
  var STATE = { datos: null, seleccion: "__general__" };

  function render() {
    var main = $("#panel");
    if (!main) return;
    main.innerHTML = "";
    var v = STATE.seleccion === "__general__"
      ? vistaGeneral(STATE.datos)
      : vistaFletero(STATE.datos, STATE.seleccion);
    main.appendChild(v);

    activarReveal(main);
    // animaciones
    setTimeout(function () {
      if (v._rings) v._rings.forEach(animaAnillo);
      animaBarras(main);
      Array.prototype.forEach.call(main.querySelectorAll(".rank__fill"), function (f, i) {
        setTimeout(function () { f.style.width = (f.getAttribute("data-w") || 2) + "%"; }, 120 + i * 60);
      });
    }, 120);

    // click en filas del ranking
    Array.prototype.forEach.call(main.querySelectorAll(".rank__row"), function (row) {
      row.addEventListener("click", function () {
        var n = row.getAttribute("data-fletero");
        seleccionar(n);
        var sel = $("#selector"); if (sel) sel.value = n;
        window.scrollTo({ top: 0, behavior: "smooth" });
      });
    });
  }

  function seleccionar(nombre) {
    STATE.seleccion = nombre;
    try { localStorage.setItem("ppp_fletero", nombre); } catch (e) {}
    render();
  }

  function poblarSelector(datos) {
    var sel = $("#selector");
    if (!sel) return;
    sel.innerHTML = "";
    var opt0 = el("option", null, "📊 Resumen general");
    opt0.value = "__general__";
    sel.appendChild(opt0);
    Object.keys(datos.porFletero).sort().forEach(function (n) {
      var o = el("option", null, n);
      o.value = n;
      sel.appendChild(o);
    });
    sel.addEventListener("change", function () { seleccionar(sel.value); });

    // Recordar selección previa (útil para el móvil de cada fletero)
    var prev = null;
    try { prev = localStorage.getItem("ppp_fletero"); } catch (e) {}
    if (prev && (prev === "__general__" || datos.porFletero[prev])) {
      STATE.seleccion = prev;
      sel.value = prev;
    }
  }

  function ultimaActualizacion(registros) {
    var max = "";
    registros.forEach(function (r) { if (r.fecha > max) max = r.fecha; });
    var lbl = $("#update-date");
    if (lbl) lbl.textContent = max ? fmtFecha(max) + " de " + (max.split("-")[0]) : "—";
  }

  function prepararDatos(registros) {
    var porFletero = agrupaPorFletero(registros);
    STATE.datos = { registros: registros, porFletero: porFletero };
    poblarSelector(STATE.datos);
    ultimaActualizacion(registros);
    render();
    var badge = $("#origen");
    if (badge) {
      var esEjemplo = window.__PPP_DATA__ && window.__PPP_DATA__.generadoDeEjemplo && !STATE._live;
      badge.hidden = !esEjemplo;
    }
  }

  function cargar() {
    var base = (window.__PPP_DATA__ && window.__PPP_DATA__.registros) || [];
    // Mostramos ya mismo los datos disponibles (ejemplo o embebidos).
    prepararDatos(base);

    var url = CONFIG.SHEET_CSV_URL;
    if (!url) return; // sin planilla conectada → nos quedamos con el ejemplo

    fetch(url, { cache: "no-store" })
      .then(function (r) { if (!r.ok) throw new Error("HTTP " + r.status); return r.text(); })
      .then(function (txt) {
        var regs = registrosDesdeCSV(txt);
        if (regs.length) {
          STATE._live = true;
          prepararDatos(regs);
        }
      })
      .catch(function (e) {
        console.warn("[PPP] No se pudo leer la planilla, se mantienen los datos de ejemplo.", e);
      });
  }

  // ---- Splash + arranque ------------------------------------------------
  function ocultarSplash() {
    var s = $("#splash");
    if (s) { s.classList.add("hide"); setTimeout(function () { if (s.parentNode) s.parentNode.removeChild(s); }, 700); }
  }

  function init() {
    safe(cargar, "cargar");
    setTimeout(ocultarSplash, 550);   // ocultar splash cuando ya hay datos
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", init);
  } else { init(); }

  // Red de seguridad para el splash (por si algo falla)
  setTimeout(ocultarSplash, 4500);
})();
