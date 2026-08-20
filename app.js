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
  // Porcentajes SIN decimales (regla de la empresa, 20/8): de ,50 para arriba
  // redondea para arriba y de ,49 para abajo. Los premios se calculan sobre
  // este número redondeado, así el fletero cobra por lo que ve en pantalla.
  function pct(x) { return x == null ? null : Math.round(x * 100); }
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

  function fmtPlata(n) {
    return "$" + String(Math.round(n || 0)).replace(/\B(?=(\d{3})+(?!\d))/g, ".");
  }
  function fmtMillones(n) {
    if (n >= 1000000) return "$" + (Math.round(n / 100000) / 10).toString().replace(".", ",") + "M";
    return "$" + Math.round(n / 1000) + " mil";
  }

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
      d.numEl.textContent = Math.round(d.raw * eased);
      if (k < 1) requestAnimationFrame(step);
      else d.numEl.textContent = Math.round(d.raw);
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
    var t = p == null ? "—" : Math.round(p) + "%";
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

    // Estadísticas del mes
    var st = ((window.__PPP_DATA__ && window.__PPP_DATA__.estadisticasFletero) || {})[nombre];
    var celdas = [];
    if (st) {
      if (st.repartos) {
        celdas.push('<div class="avg"><span class="avg__k">Repartos hechos · ' + mesNombre + '</span><b>' + st.repartos + '</b></div>');
      }
      celdas.push('<div class="avg"><span class="avg__k">Clientes rechazados completos</span><b class="rojo">' + st.recTot + '</b></div>');
      celdas.push('<div class="avg"><span class="avg__k">Boletas rechazadas completas</span><b class="ambar">' + st.recBol + '</b></div>');
      celdas.push('<div class="avg"><span class="avg__k">Clientes entregados · ' + mesNombre + '</span><b>' + st.cliEnt + ' / ' + st.cliSac + '</b></div>');
      celdas.push('<div class="avg"><span class="avg__k">Boletas entregadas · ' + mesNombre + '</span><b>' + st.compEnt + ' / ' + st.compSac + '</b></div>');
      celdas.push('<div class="avg"><span class="avg__k">Items rechazados · ' + mesNombre + ' <small>(productos)</small></span><b class="ambar">' + (st.itemsRech || 0) + '</b></div>');
      celdas.push('<div class="avg"><span class="avg__k">Plata rechazada · ' + mesNombre + '</span><b class="rojo">' + fmtPlata(st.impRech || 0) + '</b></div>');
    }
    celdas.push('<div class="avg"><span class="avg__k">Cartones de ' + mesNombre + '</span><b>' + promMes.cartones + ' / ' + promMes.cartonesTot + '</b></div>');
    var proms2 = el("div", "avgs reveal");
    proms2.innerHTML = celdas.join("");
    cont.appendChild(proms2);

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
        return '<div class="chart__row" title="' + m.motivo.replace(/"/g, "&quot;") + ' · ' + p + '% de sus boletas rechazadas">' +
          '<div class="chart__top"><span class="chart__label">' + m.motivo + '</span>' +
          '<b class="chart__val">' + p + '%</b></div>' +
          '<i class="chart__track"><i class="rank__fill rank__fill--low" style="width:2%" data-w="' + w + '"></i></i>' +
        '</div>';
      }).join("");
      card.innerHTML = '<h2 class="chart__title">📋 Motivos de sus boletas rechazadas</h2>' + rows;
      cont.appendChild(card);
    }

    // Su entrega por proveedor (en plata)
    var pfl = ((window.__PPP_DATA__ && window.__PPP_DATA__.proveedoresPorFletero) || {})[nombre] || [];
    if (pfl.length) {
      var maxP = 0;
      pfl.forEach(function (p) { if (p.pct > maxP) maxP = p.pct; });
      var cardP = el("div", "chart reveal");
      var rowsP = pfl.map(function (p) {
        var w = Math.max(4, Math.round(100 * p.pct / (maxP || 1)));
        return '<div class="chart__row">' +
          '<div class="chart__top"><span class="chart__label">' + p.prov + '</span>' +
          '<b class="chart__val">' + Math.round(p.pct) + '%</b></div>' +
          '<i class="chart__track"><i class="rank__fill rank__fill--' + claseColor(p.pct) + '" style="width:2%" data-w="' + w + '"></i></i>' +
        '</div>';
      }).join("");
      cardP.innerHTML = '<h2 class="chart__title">🏭 Su entrega por proveedor</h2>' + rowsP;
      cont.appendChild(cardP);
    }

    return cont;
  }

  // ---- Tarjeta de cierre de mes -----------------------------------------
  // Banner con los números FINALES del mes anterior ya cerrado. Aparece del
  // día 10 en adelante; ?cierre=1 la fuerza para previsualizar cualquier día;
  // se puede cerrar con la X (reaparece el mes siguiente).
  function fmtNum(n) { return String(Math.round(n || 0)).replace(/\B(?=(\d{3})+(?!\d))/g, "."); }
  function fmtPctN(n) {
    if (n == null) return "—";
    return Math.round(n) + "%";
  }
  function tarjetaCierreMes(datos) {
    var ma = datos.mesAnterior;
    if (!ma || !ma.ranking || !ma.ranking.length) return null;
    var forzar = /[?&]cierre=1/.test(location.search);
    if (!forzar) {
      if (new Date().getDate() < 10) return null;
      try { if (localStorage.getItem("ppp_cierre_" + ma.clave)) return null; } catch (e) {}
    }
    var nombreMes = (NOMBRES_MES[(ma.mes || 1) - 1] || "") + " " + (ma.anio || "");
    function dato(k, v) {
      return '<div class="cierre__dato"><span>' + k + '</span><b>' + v + '</b></div>';
    }
    var filas =
      dato("Repartos", fmtNum(ma.repartos)) +
      dato("Retorno de cartón", fmtPctN(ma.cartonGeneral)) +
      dato("Boletas entregadas", fmtNum(ma.boletasEnt) + " / " + fmtNum(ma.boletasSac)) +
      dato("Clientes entregados", fmtNum(ma.clientesEnt) + " / " + fmtNum(ma.clientesSac)) +
      dato("Plata rechazada", fmtPlata(ma.plataRech)) +
      dato("Premios pagados", fmtPlata(ma.premiosTotal));
    var card = el("div", "cierre reveal");
    card.innerHTML =
      '<button class="cierre__x" aria-label="Cerrar">&times;</button>' +
      '<div class="cierre__top">🎉 Mirá cómo cerró ' + nombreMes + '</div>' +
      '<div class="cierre__ef"><span class="cierre__efnum">' + fmtPctN(ma.efGeneral) + '</span>' +
        '<span class="cierre__eflbl">efectividad de entrega del mes<br>' + fmtNum(ma.fleteros) + ' fleteros</span></div>' +
      '<div class="cierre__grid">' + filas + '</div>' +
      '<button class="cierre__ver">Ver cómo quedó cada fletero →</button>';
    card.querySelector(".cierre__x").addEventListener("click", function () {
      try { localStorage.setItem("ppp_cierre_" + ma.clave, "1"); } catch (e) {}
      card.parentNode && card.parentNode.removeChild(card);
    });
    card.querySelector(".cierre__ver").addEventListener("click", function () {
      seleccionar("__cierre__");
      window.scrollTo({ top: 0, behavior: "smooth" });
    });
    return card;
  }

  // ---- Vista: ranking final del mes cerrado -----------------------------
  // Días hábiles (lunes a viernes) de un mes ya cerrado. Los feriados NO se
  // descuentan (regla de la empresa: el feriado que no se trabaja se compensa
  // repartiendo el sábado). Solo se usa como respaldo, si el dato de asistencia
  // no viene guardado en el historial (meses viejos).
  function habilesDeMes(anio, mes) {
    var dias = new Date(anio, mes, 0).getDate();
    var h = 0;
    for (var d = 1; d <= dias; d++) {
      var w = new Date(anio, mes - 1, d).getDay();
      if (w >= 1 && w <= 5) h++;
    }
    return h;
  }

  function vistaCierreDetalle(datos) {
    var cont = el("div", "view");
    var ma = datos.mesAnterior;
    var volver = el("button", "volver", "← Volver al resumen");
    volver.addEventListener("click", function () {
      seleccionar("__general__");
      var sel = $("#selector"); if (sel) sel.value = "__general__";
      window.scrollTo({ top: 0, behavior: "smooth" });
    });
    cont.appendChild(volver);
    if (!ma || !ma.ranking || !ma.ranking.length) {
      cont.appendChild(el("p", "muted", "No hay detalle del mes anterior todavía."));
      return cont;
    }
    var nombreMesCorto = NOMBRES_MES[(ma.mes || 1) - 1] || "";
    var nombreMes = nombreMesCorto + " " + (ma.anio || "");
    var ASIST_MIN_PREMIO = 85;
    var habiles = habilesDeMes(ma.anio || 2026, ma.mes || 1);

    // Tarjeta para llevarse la tabla en papel al momento de pagar
    var pc = el("div", "printcard reveal");
    pc.innerHTML =
      '<div class="printcard__txt"><b>🖨️ Imprimí los premios</b></div>' +
      '<button class="printcard__btn" type="button">Imprimir tabla</button>';
    pc.querySelector(".printcard__btn").addEventListener("click", function () { window.print(); });
    cont.appendChild(pc);

    // Encabezado que solo aparece en el papel (la web ya tiene su propio título)
    var enc = el("div", "printhead");
    var hoyTxt = new Date().toLocaleDateString("es-AR");
    enc.innerHTML = '<b>Primeros Productos Pehuenia S.A. · Premios de ' + nombreMes + '</b>' +
      '<span>Ranking final de fleteros · impreso el ' + hoyTxt + '</span>';
    cont.appendChild(enc);

    var lista = ma.ranking.slice().sort(function (a, b) { return (b.efE || 0) - (a.efE || 0); });
    var tabla = el("div", "rank reveal");
    var head =
      '<div class="rank__head"><span>#</span><span>Fletero</span>' +
      '<span class="rank__num">Rep.</span><span class="rank__num">Asist.</span>' +
      '<span class="rank__num">Entrega</span>' +
      '<span class="rank__num">Cartón</span><span class="rank__num">Premio</span>' +
      '<span class="rank__firma">Firma</span></div>';
    var totalPremios = 0;
    var body = lista.map(function (f, i) {
      var medal = i === 0 ? "🥇" : i === 1 ? "🥈" : i === 2 ? "🥉" : (i + 1);
      var bar = Math.max(4, Math.min(100, f.efE || 0));
      totalPremios += (f.premio || 0);
      var premioHTML = (f.premio > 0)
        ? '<span class="rank__num rank__prize"><i class="rank__prize-ico">💰</i>' + fmtPlata(f.premio) + '</span>'
        : '<span class="rank__num rank__prize rank__prize--none">—</span>';
      // Asistencia: viene del historial; si es un mes viejo se recalcula acá.
      var asist = (f.asist != null) ? f.asist
        : (habiles > 0 ? Math.min(100, Math.round(100 * (f.repartos || 0) / habiles)) : null);
      var asistHTML = asist == null
        ? '<span class="rank__num">—</span>'
        : '<span class="rank__num"><span class="chip ' +
          (asist >= ASIST_MIN_PREMIO ? "chip--ok" : "chip--low") + '"' +
          ' title="' + (f.repartos || 0) + ' repartos en ' + habiles + ' días hábiles">' +
          asist + '%</span></span>';
      return '<div class="rank__row rank__row--static">' +
        '<span class="rank__pos">' + medal + '</span>' +
        '<span class="rank__name"><b>' + f.nombre + '</b>' +
          '<i class="rank__track"><i class="rank__fill rank__fill--' + claseColor(f.efE) + '" style="width:2%" data-w="' + bar + '"></i></i>' +
        '</span>' +
        '<span class="rank__num">' + f.repartos + '</span>' +
        asistHTML +
        '<span class="rank__num">' + chip(f.efE) + '</span>' +
        '<span class="rank__num">' + chip(f.efC) + '</span>' +
        premioHTML +
        '<span class="rank__firma"></span>' +
      '</div>';
    }).join("");
    var cobran = lista.filter(function (f) { return f.premio > 0; }).length;
    var filaTotal =
      '<div class="rank__row rank__row--total">' +
      '<span class="rank__pos"></span>' +
      '<span class="rank__name"><b>TOTAL A PAGAR</b><em>' + cobran + ' de ' + lista.length + ' fleteros cobran</em></span>' +
      '<span class="rank__num"></span><span class="rank__num"></span>' +
      '<span class="rank__num"></span><span class="rank__num"></span>' +
      '<span class="rank__num rank__prize">' + fmtPlata(totalPremios) + '</span>' +
      '<span class="rank__firma"></span>' +
      '</div>';
    tabla.innerHTML =
      '<h2 class="rank__title">🏁 Así cerró ' + nombreMes + ' · ranking final</h2>' +
      '<div class="rank__grid rank__grid--cierre">' + head + body + filaTotal + '</div>' +
      '<p class="rank__hint">Asistencia, entrega, cartón y premio definitivos de cada fletero en ' +
        nombreMesCorto + '. Los premios requieren ' + ASIST_MIN_PREMIO + '% de asistencia o más.</p>';
    cont.appendChild(tabla);
    return cont;
  }

  function vistaGeneral(datos) {
    var cont = el("div", "view");
    var cierre = tarjetaCierreMes(datos);
    if (cierre) cont.appendChild(cierre);
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

    // Días hábiles (lunes a viernes) del mes, con UN DÍA DE ATRASO: el último día
    // con datos no cuenta porque sus viajes todavía no están cerrados.
    // Los feriados NO se descuentan (regla de la empresa: el feriado que no se
    // trabaja se compensa repartiendo el sábado; data.js trae la lista igual).
    var ultimaFecha = fechasTodas.length ? fechasTodas[fechasTodas.length - 1] : "";
    var habiles = 0;
    if (mesPrefijo && ultimaFecha) {
      var aa = parseInt(mesPrefijo.slice(0, 4), 10), mm = parseInt(mesPrefijo.slice(5), 10);
      var ultDia = parseInt(ultimaFecha.slice(8), 10);
      for (var dd = 1; dd < ultDia; dd++) {   // < : excluye el día en curso
        var dow = new Date(aa, mm - 1, dd).getDay();
        if (dow >= 1 && dow <= 5) habiles++;
      }
    }

    // Datos base para los rankings (total del mes en curso)
    var filas = nombres.map(function (n) {
      var regsMes = delMes(datos.porFletero[n].regs);
      var p = promedioPeriodo(regsMes);
      // Asistencia: REPARTOS HECHOS (sin contar el día en curso) vs días hábiles
      // cerrados, con tope 100% (un doble reparto compensa un día no trabajado,
      // pero no suma más de 100). Si el data.js es viejo y no trae "repartos",
      // se cuenta 1 por día con entregas, como antes.
      var trab = 0;
      regsMes.forEach(function (r) {
        if (r.fecha === ultimaFecha) return;
        if (r.repartos) trab += r.repartos;
        else if (r.entregas_asignadas > 0) trab += 1;
      });
      var asist = habiles > 0 ? Math.min(100, Math.round(100 * trab / habiles)) : null;
      return { nombre: n, zona: datos.porFletero[n].zona, efE: pct(p.efE), efR: pct(p.efR), asist: asist, diasTrab: trab };
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
    // Sin 85% de asistencia (días con reparto / días hábiles) no se cobra premio.
    var ASIST_MIN = 85;
    function tablaRanking(titulo, campo, etiqueta, conAsist, premiosTxt) {
      var lista = filas.filter(function (f) { return f[campo] != null; })
        .sort(function (a, b) { return (b[campo] || 0) - (a[campo] || 0); });
      if (!lista.length) return null;

      var tabla = el("div", "rank reveal");
      var head =
        '<div class="rank__head"><span>#</span><span>Fletero</span>' +
        (conAsist ? '<span class="rank__num">Asist.</span>' : '') +
        '<span class="rank__num">Premio</span>' +
        '<span class="rank__num">' + etiqueta + '</span></div>';
      var body = lista.map(function (f, i) {
        var medal = i === 0 ? "🥇" : i === 1 ? "🥈" : i === 2 ? "🥉" : (i + 1);
        var bar = Math.max(4, Math.min(100, f[campo] || 0));
        var cumpleAsist = (f.asist == null || f.asist >= ASIST_MIN);
        var premio = cumpleAsist ? premioPara(campo, f[campo]) : null;
        var premioHTML = premio == null
          ? '<span class="rank__num rank__prize rank__prize--none"' +
            (!cumpleAsist ? ' title="No cobra premio: asistencia menor al ' + ASIST_MIN + '%"' : '') + '>—</span>'
          : '<span class="rank__num rank__prize">💰 ' + fmtPremio(premio) + '</span>';
        var asistHTML = "";
        if (conAsist) {
          asistHTML = f.asist == null
            ? '<span class="rank__num rank__prize--none">—</span>'
            : '<span class="rank__num"><span class="chip ' + (f.asist >= ASIST_MIN ? "chip--ok" : "chip--low") + '"' +
              ' title="' + f.diasTrab + ' repartos en ' + habiles + ' días hábiles">' + f.asist + '%</span></span>';
        }
        return '<button class="rank__row" data-fletero="' + f.nombre.replace(/"/g, "&quot;") + '">' +
          '<span class="rank__pos">' + medal + '</span>' +
          '<span class="rank__name"><b>' + f.nombre + '</b>' +
            (f.zona ? '<em>' + f.zona + '</em>' : '') +
            '<i class="rank__track"><i class="rank__fill rank__fill--' + claseColor(f[campo]) + '" style="width:2%" data-w="' + bar + '"></i></i>' +
          '</span>' +
          asistHTML +
          premioHTML +
          '<span class="rank__num">' + chip(f[campo]) + '</span>' +
        '</button>';
      }).join("");
      tabla.innerHTML =
        '<h2 class="rank__title">' + titulo + '</h2>' +
        '<div class="rank__grid ' + (conAsist ? "rank__grid--asist" : "rank__grid--simple") + '">' + head + body + '</div>' +
        '<p class="rank__hint">Tocá un fletero para ver su detalle.' +
        (premiosTxt ? ' ' + premiosTxt : '') + '</p>';
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
    if (gRech) {
      gRech.classList.add("chart--link");
      gRech.innerHTML += '<p class="chart__more">Tocá acá para ver el análisis completo: zonas, vendedores y clientes →</p>';
      gRech.addEventListener("click", function () {
        seleccionar("__rechazos__");
        var sel = $("#selector"); if (sel) sel.value = "__general__";
        window.scrollTo({ top: 0, behavior: "smooth" });
      });
    }
    if (gMot || gRech) {
      var fila = el("div", "charts");
      if (gMot) fila.appendChild(gMot);
      if (gRech) fila.appendChild(gRech);
      cont.appendChild(fila);
    }

    var rankE = tablaRanking("🚚 Ranking · Efectividad de entrega · total " + mesNombre, "efE", "Entrega", true,
      "💰 Premios por efectividad de entrega: de 90% a 94,99% cobrás <b>$50.000</b> · de 95% a 100% cobrás <b>$100.000</b>. Requisito: " + ASIST_MIN + "% de asistencia o más.");
    var rankR = tablaRanking("📦 Ranking · Retorno de cartón · total " + mesNombre, "efR", "Cartón", false,
      "💰 Premios por retorno de cartón: de 60% a 69,99% cobrás <b>$50.000</b> · de 70% a 79,99% cobrás <b>$100.000</b> · de 80% a 100% cobrás <b>$150.000</b>. Requisito: " + ASIST_MIN + "% de asistencia o más.");
    if (rankE) cont.appendChild(rankE);
    if (rankR) cont.appendChild(rankR);

    return cont;
  }

  // Vista de análisis de rechazos (se abre tocando la tarjeta del resumen)
  function vistaRechazos(datos) {
    var cont = el("div", "view");
    var an = (window.__PPP_DATA__ && window.__PPP_DATA__.analisisRechazos) || null;
    var fechasT = fechasUnicas(datos.registros);
    var mesPrefijo = fechasT.length ? fechasT[fechasT.length - 1].slice(0, 7) : "";
    var mesNombre = mesPrefijo ? NOMBRES_MES[parseInt(mesPrefijo.slice(5), 10) - 1] : "mes";

    var volver = el("button", "volver", "← Volver al resumen");
    volver.addEventListener("click", function () {
      seleccionar("__general__");
      var sel = $("#selector"); if (sel) sel.value = "__general__";
    });
    cont.appendChild(volver);
    cont.appendChild(el("h2", "rank__title", "🔎 Análisis de rechazos · " + mesNombre));

    if (!an) {
      cont.appendChild(el("p", "muted", "Todavía no hay datos de análisis (falta cargar el reporte de ventas)."));
      return cont;
    }

    // Plata rechazada del mes (los montos facturados siguen privados)
    if (an.importe) {
      var stat = el("div", "chart reveal");
      stat.innerHTML =
        '<h2 class="chart__title">💸 Plata rechazada en ' + mesNombre + '</h2>' +
        '<div class="bigmoney">' + fmtPlata(an.importe) + '</div>' +
        '<p class="chart__note">Suma de todas las notas de crédito por rechazo del mes, productos sueltos incluidos (no los canjes).</p>';
      cont.appendChild(stat);
    }

    function tarjeta(titulo, items) {
      if (!items || !items.length) return null;
      var max = 0;
      items.forEach(function (it) { if (it._v > max) max = it._v; });
      var card = el("div", "chart reveal");
      var rows = items.map(function (it) {
        var w = Math.max(4, Math.round(100 * it._v / (max || 1)));
        return '<div class="chart__row"><div class="chart__top">' +
          '<span class="chart__label">' + it._l + '</span>' +
          '<b class="chart__val">' + it._t + '</b></div>' +
          '<i class="chart__track"><i class="rank__fill rank__fill--' + (it._c || "low") + '" style="width:2%" data-w="' + w + '"></i></i>' +
        '</div>';
      }).join("");
      card.innerHTML = '<h2 class="chart__title">' + titulo + '</h2>' + rows;
      return card;
    }

    var zonas = (an.zonas || []).map(function (z) {
      return { _l: z.nombre, _v: z.pct, _t: Math.round(z.pct) + "% <span class='chart__cnt'>(" + z.rech + " de " + z.sac + ")</span>" };
    });
    var vendedores = (an.vendedores || []).map(function (v) {
      return { _l: v.nombre, _v: v.pct, _t: Math.round(v.pct) + "% <span class='chart__cnt'>(" + v.rech + " de " + v.sac + ")</span>" };
    });
    var clientes = (an.clientes || []).map(function (c) {
      return { _l: c.nombre + (c.loc ? " · " + c.loc : ""), _v: c.cantidad, _t: c.cantidad + " rechazos" };
    });

    var fila = el("div", "charts");
    var tz = tarjeta("📍 % de rechazo por zona <span class='chart__cnt'>(ventas caídas completas)</span>", zonas);
    var tv = tarjeta("🧑‍💼 % de rechazo por vendedor <span class='chart__cnt'>(ventas caídas completas)</span>", vendedores);
    if (tz) fila.appendChild(tz);
    if (tv) fila.appendChild(tv);
    if (tz || tv) cont.appendChild(fila);

    // % entregado por proveedor (en plata) + clientes que más rechazan
    var provs = (an.proveedores || []).map(function (p) {
      return { _l: p.nombre, _v: p.pct, _c: claseColor(p.pct), _t: Math.round(p.pct) + "%" };
    });
    var fila2 = el("div", "charts");
    var tp = tarjeta("🏭 % entregado por proveedor <span class='chart__cnt'>(en plata)</span>", provs);
    var tc = tarjeta("🏪 Clientes que más rechazan", clientes);
    if (tp) fila2.appendChild(tp);
    if (tc) fila2.appendChild(tc);
    if (tp || tc) cont.appendChild(fila2);

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
    var v = STATE.seleccion === "__general__" ? vistaGeneral(STATE.datos)
      : STATE.seleccion === "__rechazos__" ? vistaRechazos(STATE.datos)
      : STATE.seleccion === "__cierre__" ? vistaCierreDetalle(STATE.datos)
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
    // No persistimos la vista transitoria del cierre de mes
    if (nombre !== "__cierre__") { try { localStorage.setItem("ppp_fletero", nombre); } catch (e) {} }
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
    var mesAnterior = (window.__PPP_DATA__ && window.__PPP_DATA__.mesAnterior) || null;
    STATE.datos = { registros: registros, porFletero: porFletero, mesAnterior: mesAnterior };
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
