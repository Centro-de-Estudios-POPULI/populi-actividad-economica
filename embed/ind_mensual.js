/* ════════════════════════════════════════════════════════════════════════════
   Monitor de Actividad Económica · indicadores adelantados (cemento, permisos de construcción,
   manufactura, energía eléctrica, transporte, carga ferroviaria): los seis como una familia.

   Pieza común montada sobre el molde (window.PM). Cada embed declara su configuración y llama a
   IND.armar({...}); los datos salen de data/ind_*.json (scripts/datos_indicadores.py): ningún número
   se escribe a mano. Sirve para series mensuales (frec 12) y trimestrales (frec 4).

   Tres vistas, las mismas en los seis gráficos:
   · «Año por año» (la de entrada): eje = ene…dic (o T1…T4), una línea por año. El año en curso es el
     protagonista (rojo, 2,5 px, único relleno, punto final con su cifra), el anterior va en turquesa y
     la franja gris es el rango (mín–máx) de los cinco años anteriores: «por encima / por debajo de lo
     normal» se ve sin contar líneas. El año en curso es PARCIAL: se compara con el mismo tramo.
   · «Crecimiento»: variación de los últimos 12 meses (4 trimestres) frente a los 12 anteriores, en serie
     continua: quita la estacionalidad y el ruido del mes. Línea en tinta, relleno por signo contra cero
     (turquesa = sube, rojo = baja) y la cifra final rotulada junto al punto (lenguaje del deck y del blog).
   · «Nivel»: la suma (o el promedio) de los últimos 12 meses con un año base = 100 elegible; relleno por
     signo contra la línea de 100. Para una tasa (uso de la capacidad) la referencia es su promedio.
   ════════════════════════════════════════════════════════════════════════════ */
(function () {
  'use strict';
  var C = PM.C, NB = '\u00a0', WJ = '\u2060';   // espacio duro y «word joiner»: «ene–jul 2026» no se parte en el guion
  // Colores comunes a los seis (validados en los dos temas contra la tarjeta: rojo 5,8:1 y 3,2:1, turquesa
  // 3,7:1 y 4,9:1; rojo-turquesa ΔE 15 con deuteranopía). Año en curso = rojo (Foco); año anterior =
  // turquesa; contexto = grises. En las vistas continuas el rojo y el turquesa pasan a ser el signo y la
  // línea va en tinta: el rojo de foco y el de «baja» nunca conviven en una misma vista.
  var ROJO = C.rojo, PREVIO = C.turquesa, TINTA = '#001219';
  var BANDA = { claro: '#8A9699', oscuro: '#6E7A7D' };       // franja del rango normal (con transparencia)
  var CONTEXTO = { claro: '#8A9699', oscuro: '#6E7A7D' };    // un año de contexto que el lector prende
  var CTX_TIP = { claro: '#C9CDCE', oscuro: '#8A9699' };     // su clave sobre el tooltip oscuro
  var NORMAL = 5;                                           // años del rango normal
  var BASE_REF = 2019;                                      // año de comparación de la cuarta cifra (prepandemia)

  var estilo = document.createElement('style');
  estilo.textContent =
    '.anios{margin:-2px 0 9px;padding:0 0 9px;border-bottom:1px solid var(--border)}' +
    '.anio-f{display:flex;align-items:baseline;gap:8px;padding:2px 0;font-size:.68rem;line-height:1.35}' +
    '.anio-a{font-family:"JetBrains Mono",monospace;font-weight:700;color:var(--pizarra);min-width:34px}' +
    '.anio-v{font-family:"JetBrains Mono",monospace;font-weight:600;color:var(--tinta);flex:1;white-space:nowrap}' +
    '.pill[aria-disabled="true"]{cursor:default}' +
    '.ref .lp.area{border-top:0}' +
    '.hz-item[hidden]{display:none}';
  document.head.appendChild(estilo);

  function pad(m) { return (m < 10 ? '0' : '') + m; }
  function suma(a, b) { return a + b; }
  function r2(x) { return x == null || isNaN(x) ? null : Math.round(x * 100) / 100; }
  function varPct(a, b) { return a == null || b == null || !b ? null : (a / b - 1) * 100; }
  function tono(x) { return x == null || Math.abs(x) < 0.05 ? 'neutro' : x > 0 ? 'bueno' : 'malo'; }
  function signoCol(x) { return x == null || x >= 0 ? PREVIO : ROJO; }

  // ── Unidades ───────────────────────────────────────────────────────────────
  var UNIDAD = {
    t: { eje: 'miles de t', div: 1e3, txt: 't' },
    m2: { eje: 'miles de m²', div: 1e3, txt: 'm²' },
    permisos: { eje: 'permisos', div: 1, txt: 'permisos' },
    indice: { eje: 'índice, 2017 = 100', div: 1, txt: '' },
    pct: { eje: '% de la capacidad', div: 1, txt: '%' }
  };
  function cifra(v, u) {
    if (v == null || isNaN(v)) return '—';
    if (u === 'indice') return PM.num(v, 1);
    if (u === 'pct') return PM.pct(v, 1);
    if (u === 'permisos') return PM.num(v, 0) + NB + 'permisos';
    var t = UNIDAD[u].txt, a = Math.abs(v);
    if (a >= 1e6) return PM.num(v / 1e6, 2) + NB + 'MM' + NB + t;
    if (a >= 1e4) return PM.num(v / 1e3, 1) + NB + 'mil' + NB + t;
    return PM.num(v, 0) + NB + t;
  }

  function armar(cfg) {
    var D, K, idx = {}, P, grafico, frec = cfg.frec || 12;
    var estado = { k: cfg.series[0].k, vista: 'anio', medida: 'mes', base: BASE_REF, i: null };
    var activos = null;   // años y franja a la vista en «Año por año»

    // ── Acceso a los datos ───────────────────────────────────────────────────
    var serie = function () { return cfg.series.filter(function (s) { return s.k === estado.k; })[0]; };
    var tipo = function () { return serie().tipo || cfg.tipo; };          // 'flujo' · 'indice' · 'tasa'
    var U = function () { return serie().unidad || cfg.unidad; };
    var S = function () { return D.series[estado.k]; };
    var tasa = function () { return tipo() === 'tasa'; };
    var flujo = function () { return tipo() === 'flujo'; };
    function pos(i) { var p = P[i]; return frec === 12 ? p.m : p.t - 1; }
    function clave(y, q) { return frec === 12 ? y + '-' + pad(q + 1) : y + '-T' + (q + 1); }
    function v(y, q) { var i = idx[clave(y, q)]; return i == null ? null : S()[i]; }
    function ultimo() { var s = S(); for (var i = s.length - 1; i >= 0; i--) if (s[i] != null) return i; return -1; }
    function primero() { var s = S(); for (var i = 0; i < s.length; i++) if (s[i] != null) return i; return 0; }
    function fin() { var i = ultimo(); return { y: P[i].a, q: pos(i) }; }
    function hace(i) { var p = P[i]; return idx[clave(p.a - 1, pos(i))]; }     // el mismo período un año antes, por FECHA
    // comparación: % en flujos e índices, puntos porcentuales en una tasa
    function dif(a, b) { return a == null || b == null ? null : tasa() ? a - b : varPct(a, b); }
    function fdif(x, dec) { return x == null ? '—' : tasa() ? PM.pp(x, dec == null ? 1 : dec) : PM.pct(x, dec == null ? 1 : dec, true); }
    function tramo(y, hasta) {             // valores de los períodos 0…hasta del año y (null si falta alguno)
      var out = [];
      for (var q = 0; q <= hasta; q++) { var x = v(y, q); if (x == null) return null; out.push(x); }
      return out;
    }
    function agregado(y, hasta) { var t = tramo(y, hasta); if (!t) return null; var s = t.reduce(suma, 0); return flujo() ? s : s / t.length; }
    function doce(i) {                     // suma (flujo) o promedio de los últimos 12 meses / 4 trimestres, sin huecos
      if (i == null || i < frec - 1) return null;
      var s = 0, a = S();
      for (var j = i - frec + 1; j <= i; j++) { if (a[j] == null) return null; s += a[j]; }
      return flujo() ? s : s / frec;
    }
    function crec(i) { var h = hace(i); return h == null ? null : dif(doce(i), doce(h)); }
    function baseVal(b) { return doce(idx[clave(b, frec - 1)]); }        // el año base entero = sus 12 meses
    function referenciaNivel() {           // 100 (año base) o, en una tasa, su promedio histórico
      if (!tasa()) return 100;
      var t = 0, n = 0;
      S().forEach(function (x) { if (x != null) { t += x; n++; } });
      return t / n;
    }
    function nivel(i) {
      var d = doce(i);
      if (d == null) return null;
      if (tasa()) return d;
      var b = baseVal(estado.base);
      return b ? d / b * 100 : null;
    }

    // ── Textos de período ───────────────────────────────────────────────────
    var TRIM_L = ['primer trimestre', 'segundo trimestre', 'tercer trimestre', 'cuarto trimestre'];
    function nomP(q, largo) { return frec === 12 ? (largo ? PM.MES_L[q] : PM.MES[q]) : (largo ? TRIM_L[q] : 'T' + (q + 1)); }
    function fechaP(y, q) { return PM.fecha(clave(y, q)); }
    function tramoCorto(y, q) {
      if (q === frec - 1) return String(y);
      if (frec === 4) return (q === 0 ? 'T1' : 'T1–' + WJ + 'T' + (q + 1)) + NB + y;
      return (q === 0 ? PM.MES[0] : PM.MES[0] + '–' + WJ + PM.MES[q]) + NB + y;
    }
    function tramoLargo(y, q) {
      if (q === frec - 1) return 'todo ' + y;
      if (frec === 4) return (q === 0 ? 'el primer trimestre' : 'T1–T' + (q + 1)) + ' de ' + y;
      return (q === 0 ? 'enero' : 'enero–' + PM.MES_L[q]) + ' de ' + y;
    }
    function tramoSin(q) { return q === frec - 1 ? 'año' : frec === 4 ? (q === 0 ? 'T1' : 'T1–T' + (q + 1)) : (q === 0 ? 'ene' : 'ene–' + PM.MES[q]); }
    function NPER() { return frec === 12 ? '12 meses' : '4 trimestres'; }

    // ── Arranque ─────────────────────────────────────────────────────────────
    (cfg.cargar ? cfg.cargar() : PM.cargar([cfg.archivo]).then(function (r) { return r[0]; })).then(function (d) {
      D = d; K = D.meses;
      K.forEach(function (k, i) { idx[k] = i; });
      P = K.map(PM.periodo);
      estado.i = fin().q;
      if (cfg.series.length > 1)
        PM.botonera(document.getElementById('b-serie'), cfg.series.map(function (s) { return [s.k, s.boton]; }), estado.k, function (k) {
          estado.k = k; anios(true); botoneraExtra(); todo();
        });
      else document.getElementById('b-serie').parentNode.hidden = true;
      PM.botonera(document.getElementById('b-vista'), [['anio', 'Año por año'], ['crec', 'Crecimiento'], ['nivel', 'Nivel']], estado.vista, function (x) {
        estado.vista = x; botoneraExtra(); todo();
      });
      anios(false);
      botoneraExtra();
      cifras(); subtitulo(); pastillas();
      var el = document.getElementById('chart');
      el.innerHTML = '';
      grafico = PM.montar(el, opciones);
      grafico.chart.on('click', function (p) {
        if (p.dataIndex == null || !p.seriesName || p.seriesName.charAt(0) === '_') return;
        panel(estado.vista === 'anio' ? p.dataIndex : continua.desde + p.dataIndex);
      });
      panel(estado.i);
      PM.alCambiarTema(function () { cifras(); pastillas(); panel(estado.i); });
      PM.preguntas(document.getElementById('preguntas'), cfg.preguntas(ayuda()).concat([comoLeer()]));
      document.getElementById('fuente').innerHTML = 'Fuente: <a href="' + D.metadata.url + '" target="_blank" rel="noopener">' + cfg.fuente + '</a>' +
        ' · Elaboración: Centro de Estudios POPULI · Datos al: ' + PM.fecha(D.metadata.ultimo);
    }).catch(function (e) { console.error(e); PM.error(document.getElementById('chart')); });

    function todo() {
      cifras(); subtitulo(); pastillas(); grafico.redibujar();
      panel(estado.vista === 'anio' ? fin().q : ultimo());
    }
    // años base que se ofrecen: 2017 (el año de referencia del INE), 2019 (prepandemia) y el anterior al último dato
    function bases() {
      var f = fin(), out = [];
      [2017, BASE_REF, f.y - 1].forEach(function (b) { if (out.indexOf(b) < 0 && b < f.y && baseVal(b) != null) out.push(b); });
      return out.length ? out : [P[primero()].a + 1];
    }
    // La tercera botonera cambia con la vista: «Medida» en el año por año, «Año base» en el nivel, nada en el crecimiento
    function botoneraExtra() {
      var item = document.getElementById('b-extra').parentNode, lbl = item.querySelector('.hz-lbl');
      if (estado.vista === 'anio') {
        item.hidden = false; lbl.textContent = 'Medida';
        PM.botonera(document.getElementById('b-extra'), [['mes', frec === 12 ? 'Del mes' : 'Del trimestre'], ['acum', flujo() ? 'Acumulado' : 'Promedio del año']], estado.medida,
          function (x) { estado.medida = x; subtitulo(); grafico.redibujar(); panel(estado.i); });
      } else if (estado.vista === 'nivel' && !tasa()) {
        var bs = bases();
        if (bs.indexOf(estado.base) < 0) estado.base = bs.indexOf(BASE_REF) >= 0 ? BASE_REF : bs[0];
        item.hidden = false; lbl.textContent = 'Año base';
        PM.botonera(document.getElementById('b-extra'), bs.map(function (b) { return [String(b), String(b)]; }), String(estado.base),
          function (x) { estado.base = +x; subtitulo(); pastillas(); grafico.redibujar(); panel(estado.i); });
      } else item.hidden = true;
    }

    // ── Años a la vista ──────────────────────────────────────────────────────
    function anios(conservar) {
      var f = fin(), lista = [String(f.y)];
      rangoAnios().anios.concat([f.y - 1]).sort(function (a, b) { return b - a; }).forEach(function (y) { if (lista.indexOf(String(y)) < 0) lista.push(String(y)); });
      var antes = activos;
      activos = new Set([String(f.y), String(f.y - 1), 'rango']);
      if (conservar && antes) {
        lista.forEach(function (y) { if (+y < f.y - 1 && antes.has(y)) activos.add(y); });
        if (!antes.has('rango')) activos.delete('rango');
        if (!antes.has(String(f.y - 1))) activos.delete(String(f.y - 1));
      }
      anios.lista = lista;
    }
    function rol(y) { var f = fin().y; return y === f ? 'cur' : y === f - 1 ? 'prev' : 'ctx'; }
    function colorAnio(y) { var r = rol(y); return r === 'cur' ? ROJO : r === 'prev' ? PREVIO : CONTEXTO; }
    function colorTip(y) { var r = rol(y); return r === 'cur' ? ROJO : r === 'prev' ? PREVIO : CTX_TIP; }
    // Lo normal: los cinco años anteriores al último, sin 2020 (el cierre por la pandemia no es un año normal:
    // estiraba la franja hacia abajo y hacía que cualquier año pareciera bueno). Si 2020 cae en el tramo, se dice.
    function rangoAnios() {
      var f = fin().y, a0 = P[primero()].a, out = [];
      for (var y = f - 1; y >= a0 && out.length < NORMAL; y--) if (y !== 2020) out.push(y);
      var lo = out[out.length - 1], hi = out[0];
      var sin = lo < 2020 && hi > 2020;
      // txt para la página (no se parte); plano para el CSV y la leyenda de la imagen (sin caracteres invisibles)
      return { anios: out, txt: lo + '–' + WJ + hi + (sin ? NB + 'sin' + NB + '2020' : ''), plano: lo + '–' + hi + (sin ? ' sin 2020' : '') };
    }

    function pastillas() {
      var el = document.getElementById('pastillas');
      el.innerHTML = '';
      if (estado.vista !== 'anio') {
        // vistas continuas: una clave fija (no son series que se apaguen) que explica la línea y los dos colores
        var cre = estado.vista === 'crec', dk = PM.dk();
        var arriba = cre ? 'Sube' : tasa() ? 'Sobre su promedio' : 'Sobre ' + estado.base, abajo = cre ? 'Baja' : tasa() ? 'Bajo su promedio' : 'Bajo ' + estado.base;
        el.innerHTML = '<span class="ref"><span class="lp" style="--pc:' + PM.col(TINTA) + '"></span>' + nombreContinua() + '</span>' +
          '<span class="ref"><span class="lp area" style="--pc:' + PM.tinte(true) + '"></span>' + arriba + '</span>' +
          '<span class="ref"><span class="lp area" style="--pc:' + PM.tinte(false) + '"></span>' + abajo + '</span>';
        return;
      }
      var f = fin(), ra = rangoAnios(), items = [{ k: String(f.y), nombre: String(f.y), fija: true }, { k: String(f.y - 1), nombre: String(f.y - 1) },
        { k: 'rango', nombre: 'Rango ' + ra.txt, area: true }];
      anios.lista.forEach(function (y) { if (+y < f.y - 1) items.push({ k: y, nombre: y }); });
      items.forEach(function (it) {
        var b = document.createElement('button');
        b.type = 'button'; b.className = 'pill'; b.dataset.k = it.k;
        b.innerHTML = '<span class="lp' + (it.area ? ' area' : '') + '"></span>' + it.nombre;
        if (it.fija) { b.setAttribute('aria-disabled', 'true'); b.title = 'El año en curso siempre está a la vista'; }
        if (it.area) b.title = 'El valor más bajo y el más alto de cada ' + (frec === 12 ? 'mes' : 'trimestre') + ' en ' + ra.txt;
        b.addEventListener('click', function () {
          if (it.fija) return;
          if (activos.has(it.k)) activos.delete(it.k); else activos.add(it.k);
          pintarPastillas(); subtitulo(); grafico.redibujar(); panel(estado.i);
        });
        el.appendChild(b);
      });
      pintarPastillas();
    }
    function pintarPastillas() {
      var dk = PM.dk();
      document.querySelectorAll('#pastillas .pill').forEach(function (b) {
        var k = b.dataset.k, col = k === 'rango' ? BANDA : colorAnio(+k);
        b.style.setProperty('--pc', k === 'rango' ? PM.rgba(BANDA, dk ? 0.55 : 0.4) : PM.col(col));
        b.style.setProperty('--pf', PM.rgba(col, dk ? 0.16 : 0.09));
        b.style.setProperty('--pb', PM.rgba(col, dk ? 0.6 : 0.45));
        b.setAttribute('aria-pressed', activos.has(k));
      });
    }

    // ── Subtítulo en lenguaje llano: qué se ve y cómo leerlo ─────────────────
    function subtitulo() {
      var f = fin(), s = serie(), ra = rangoAnios(), txt, per = frec === 12 ? 'mes' : 'trimestre';
      if (estado.vista === 'anio') {
        var que = estado.medida === 'mes' ? s.mide + ' en cada ' + per
          : flujo() ? s.mide + ', acumulado desde enero' : s.mide + ', promedio desde ' + (frec === 12 ? 'enero' : 'el primer trimestre');
        txt = que + '. Cada línea es un año: ' + f.y + ', en rojo, ' + (f.q === frec - 1 ? 'está completo' : 'llega hasta ' + nomP(f.q, true)) +
          (activos && activos.has('rango') ? '; la franja gris es el rango de ' + ra.txt : '') + '.';
      } else if (estado.vista === 'crec') {
        txt = (tasa() ? s.mide + ': promedio de los últimos ' + NPER() + ' menos el de los ' + NPER() + ' anteriores, en puntos'
          : s.mide + ': variación de los últimos ' + NPER() + ' frente a los ' + NPER() + ' anteriores') +
          ', que quita la estacionalidad y el ruido de cada ' + per + '. Turquesa = sube; rojo = baja.';
      } else {
        txt = tasa() ? s.mide + ', promedio de los últimos ' + NPER() + '. Turquesa = por encima de su promedio desde ' + P[primero()].a + '; rojo = por debajo.'
          : s.mide + ': ' + (flujo() ? 'suma' : 'promedio') + ' de los últimos ' + NPER() + ', con ' + estado.base + ' = 100. Turquesa = por encima de ' + estado.base + '; rojo = por debajo.';
      }
      document.getElementById('sub').textContent = txt;
    }

    // ── Cifras: pocas, grandes, con su significado en la línea chica ──────────
    function cifras() {
      var s = S(), iU = ultimo(), f = fin(), u = U(), sr = serie(), h = hace(iU);
      var y1 = dif(s[iU], h == null ? null : s[h]);
      var a = agregado(f.y, f.q), b = agregado(f.y - 1, f.q), ya = dif(a, b);
      // lo normal: el promedio del mismo tramo en los cinco años anteriores
      var ra = rangoAnios(), norm = [], hist = [];
      ra.anios.forEach(function (y) { var x = agregado(y, f.q); if (x != null) norm.push(x); });
      for (var z = P[primero()].a; z <= f.y; z++) hist.push(agregado(z, f.q));
      var prom = norm.length ? norm.reduce(suma, 0) / norm.length : null, yn = dif(a, prom);
      var yb = dif(doce(iU), baseVal(BASE_REF));
      var serieN = function (fn, n) { var o = []; for (var i = Math.max(0, iU - n + 1); i <= iU; i++) o.push(fn(i)); return o; };
      var nper = frec === 12 ? 24 : 12;
      var fechaH = h == null ? fechaP(f.y - 1, f.q) : PM.fecha(K[h]);
      document.getElementById('kpis').innerHTML =
        PM.kpi({ color: signoCol(y1), rotulo: sr.corto + ' · ' + PM.fecha(K[iU]), valor: cifra(s[iU], u),
          delta: y1 == null ? 'sin dato de ' + (f.y - 1) : fdif(y1) + ' vs ' + fechaH, tono: tono(y1),
          serie: serieN(function (i) { return s[i]; }, nper) }) +
        PM.kpi({ color: signoCol(ya), rotulo: f.q === frec - 1 ? 'En ' + f.y : 'En lo que va de ' + f.y, valor: fdif(ya),
          delta: tramoCorto(f.y, f.q) + ' vs ' + tramoCorto(f.y - 1, f.q), tono: tono(ya),
          serie: serieN(function (i) { var p = P[i]; return dif(agregado(p.a, pos(i)), agregado(p.a - 1, pos(i))); }, nper) }) +
        PM.kpi({ color: signoCol(yn), rotulo: tramoCorto(f.y, f.q) + ' frente a lo normal', valor: fdif(yn),
          delta: 'vs promedio ' + ra.txt, tono: tono(yn),
          serie: hist }) +
        PM.kpi({ color: signoCol(yb), rotulo: 'Frente a ' + BASE_REF, valor: fdif(yb),
          delta: 'últimos ' + NPER() + ' vs ' + BASE_REF, tono: tono(yb),
          serie: serieN(function (i) { return doce(i); }, frec === 12 ? 36 : 12) });
    }

    // ── Gráfico ──────────────────────────────────────────────────────────────
    function valorAnio(y, q) {
      if (estado.medida === 'mes') return v(y, q);
      var t = tramo(y, q); if (!t) return null;
      var s = t.reduce(suma, 0); return flujo() ? s : s / t.length;
    }
    function banda() {                     // mínimo, máximo y promedio de cada período en los años del rango normal
      var ra = rangoAnios(), out = [];
      for (var q = 0; q < frec; q++) {
        var xs = [];
        ra.anios.forEach(function (y) { var x = valorAnio(y, q); if (x != null) xs.push(x); });
        out.push(xs.length < 3 ? null : { min: Math.min.apply(null, xs), max: Math.max.apply(null, xs), prom: xs.reduce(suma, 0) / xs.length });
      }
      return out;
    }
    function fmtNivel(x) { return cifra(x, U()); }
    function corto(x) {
      var u = U();
      return u === 'indice' ? PM.num(x, 1) : u === 'pct' ? PM.pct(x, 1) : u === 'permisos' ? PM.num(x, 0) : PM.auto(x / UNIDAD[u].div);
    }
    // la unidad arranca sobre el eje y se extiende hacia el trazado; el aire arriba y abajo deja lugar a la cifra
    // del punto final cuando el último dato es el mínimo o el máximo (no pisa los años del eje)
    var NOMBRE_EJE = { nameTextStyle: { align: 'left' }, boundaryGap: ['8%', '8%'] };
    function ejeYnivel() {
      var u = UNIDAD[U()];
      return PM.ejeY({ unidad: u.eje, fmt: U() === 'pct' ? 'pct' : function (x) { return PM.tick(x / u.div); }, extra: PM.mezclar({ scale: true }, NOMBRE_EJE) });
    }
    // Cifra junto al punto final: del lado contrario a la línea que llega (arriba si sube, abajo si baja);
    // cerca del borde derecho se recuesta hacia la izquierda. En las vistas continuas el punto es hueco (deck).
    function etiqueta(color, txt, n, x, total, previo, hueco) {
      var mp = hueco ? { symbol: 'circle', symbolSize: 8, silent: true, animation: false, data: [{ coord: [n, x] }],
            itemStyle: { color: PM.var('--card') || '#fff', borderColor: PM.col(color), borderWidth: 2 } }
        : PM.puntoFinal(color, n, x);
      // la cifra va en una capa propia y con FONDO del color de la tarjeta: el halo (textBorder) sólo contornea
      // cada letra y la línea del año anterior se veía entre los dígitos («33‡4,4», medido en el teléfono)
      mp.zlevel = 1;
      var baja = previo != null && previo > x, borde = n >= total * 0.85;
      mp.label = { show: true, position: baja ? 'bottom' : 'top', distance: 8, formatter: txt, color: hueco ? PM.col(color) : PM.tx(color), fontFamily: PM.mono(), fontWeight: 700,
        fontSize: PM.pequeno() ? 10.5 : 11.5, backgroundColor: PM.var('--card') || '#fff', padding: [1, 3] };
      if (borde) { mp.label.align = 'right'; mp.label.offset = [4, 0]; }
      return mp;
    }
    function ejeX(el) {
      var dk = PM.dk(), fs = PM.pequeno() ? 10 : 10.5, linea = dk ? '#3A4549' : '#C9CDCE';
      var cats = frec === 12 ? PM.MES.slice() : ['T1', 'T2', 'T3', 'T4'];
      // meses: todos si entran; si no, uno de cada dos. Misma letra que el eje de tiempo del molde.
      var util = Math.max(140, el.clientWidth - 70), px = util / (frec - 1), necesita = 3 * (fs * 0.6 + 0.15) + (PM.pequeno() ? 12 : 14);
      var paso = frec === 4 || px >= necesita ? 0 : 2 * px >= necesita ? 1 : 2;
      return {
        type: 'category', data: cats, boundaryGap: false,
        axisLine: { onZero: false, lineStyle: { color: linea } },
        axisTick: { show: true, length: 4, interval: 0, lineStyle: { color: linea } },
        axisLabel: { interval: paso, margin: 9, fontFamily: PM.mono(), fontSize: fs, fontWeight: 500, color: PM.var('--pizarra') },
        splitLine: { show: false }
      };
    }
    function opciones() {
      var el = document.getElementById('chart');
      return estado.vista === 'anio' ? opcAnio(el) : opcContinua(el);
    }

    // Líneas CURVAS (pedido de Carlos): spline con monotonía en x, la misma regla de las áreas del Monetario.
    // Sin smoothMonotone la curva sobrepasa en los quiebres e inventa picos y valles que el dato no tiene.
    // Todo lo que se dibuja junto (franja, línea, relleno por signo) lleva la MISMA curva o no calzan.
    var SUAVE = { smooth: 0.35, smoothMonotone: 'x' };

    function opcAnio(el) {
      var dk = PM.dk(), f = fin(), ra = rangoAnios(), bd = banda(), series = [];
      var lista = anios.lista.slice().reverse().filter(function (y) { return activos.has(y); });   // de más viejo a más nuevo
      if (activos.has('rango')) {
        // la franja: un piso invisible y el ancho apilado encima (nombres con «_»: fuera de la leyenda y del CSV del molde)
        series.push(PM.mezclar({ type: 'line', name: '_min', data: bd.map(function (x) { return x ? r2(x.min) : null; }), stack: 'rango', symbol: 'none', silent: true,
          lineStyle: { width: 0, opacity: 0 }, emphasis: { disabled: true }, z: 1 }, SUAVE));
        series.push(PM.mezclar({ type: 'line', name: '_rango', data: bd.map(function (x) { return x ? r2(x.max - x.min) : null; }), stack: 'rango', symbol: 'none', silent: true,
          lineStyle: { width: 0, opacity: 0 }, areaStyle: { color: PM.rgba(BANDA, dk ? 0.3 : 0.2) }, emphasis: { disabled: true }, z: 1 }, SUAVE));
      }
      lista.forEach(function (ys) {
        var y = +ys, r = rol(y), w = r === 'cur' ? 2.5 : r === 'prev' ? 1.75 : 1.25;
        var data = [];
        for (var q = 0; q < frec; q++) data.push(r2(valorAnio(y, q)));
        var o = PM.linea(colorAnio(y), { ancho: w, punto: r === 'cur' ? 8 : 6, extra: {
          name: ys, data: data, connectNulls: false, z: r === 'cur' ? 10 : r === 'prev' ? 6 : 3, smooth: SUAVE.smooth, smoothMonotone: SUAVE.smoothMonotone,
          // el cursor por eje «resalta» todas las series: el énfasis no cambia color ni grosor (rompería el Foco)
          emphasis: { focus: 'series', lineStyle: { width: w } } } });
        // Punto HUECO en cada mes (o trimestre), sutil: sólo en las dos líneas protagonistas. En los años de contexto
        // y en las vistas continuas (cientos de meses) sería una hilera que ensucia. El último dato del año en curso
        // conserva su punto lleno con halo (markPoint), que marca dónde termina.
        if (r !== 'ctx') {
          o.showSymbol = true;
          o.symbolSize = r === 'cur' ? 6 : 5;
          o.itemStyle = { color: PM.var('--card') || '#fff', borderColor: PM.col(colorAnio(y)), borderWidth: r === 'cur' ? 1.6 : 1.3 };
        }
        if (r === 'cur') {
          o.areaStyle = { color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [{ offset: 0, color: PM.rgba(ROJO, dk ? 0.2 : 0.13) }, { offset: 1, color: PM.rgba(ROJO, 0) }]) };
          var n = data.length - 1;
          while (n > 0 && data[n] == null) n--;
          if (data[n] != null) o.markPoint = etiqueta(ROJO, corto(data[n]), n, data[n], frec - 1, n ? data[n - 1] : null);
        }
        series.push(o);
      });
      // descargas: la leyenda de la imagen y la tabla de datos de esta vista
      PM.leyendaImagen = function () {
        var out = [];
        lista.slice().reverse().forEach(function (y) { if (rol(+y) !== 'ctx') out.push({ name: y, itemStyle: { color: PM.col(colorAnio(+y)) }, forma: 'linea' }); });
        if (activos.has('rango')) out.push({ name: 'Rango ' + ra.plano, itemStyle: { color: PM.rgba(BANDA, 0.4) } });
        var ctx = lista.filter(function (y) { return rol(+y) === 'ctx'; });
        if (ctx.length) out.push({ name: ctx.join(', '), itemStyle: { color: PM.col(CONTEXTO) }, forma: 'linea' });
        return out;
      };
      PM.tablaDatos = function () {
        var cols = [frec === 12 ? 'Mes' : 'Trimestre'].concat(lista.slice().reverse());
        if (activos.has('rango')) cols.push('Mínimo ' + ra.plano, 'Máximo ' + ra.plano, 'Promedio ' + ra.plano);
        var filas = [];
        for (var q = 0; q < frec; q++) {
          var fila = [frec === 12 ? PM.MES_L[q] : 'T' + (q + 1)];
          lista.slice().reverse().forEach(function (y) { fila.push(r2(valorAnio(+y, q))); });
          if (activos.has('rango')) { var b = bd[q]; fila.push(b ? r2(b.min) : null, b ? r2(b.max) : null, b ? r2(b.prom) : null); }
          filas.push(fila);
        }
        return { cols: cols, filas: filas };
      };
      return {
        grid: PM.grid({ right: PM.pequeno() ? 16 : 18 }),
        xAxis: ejeX(el),
        yAxis: ejeYnivel(),
        tooltip: PM.tooltip(function (ps) {
          if (!ps.length) return '';
          var q = ps[0].dataIndex;
          panel(q);
          var h = PM.ttTitulo(nomP(q, true), estado.medida === 'mes' ? serie().corto : (flujo() ? 'Acumulado desde enero' : 'Promedio en el año'));
          lista.slice().reverse().forEach(function (y) { var x = valorAnio(+y, q); if (x != null) h += PM.ttFila(colorTip(+y), y, fmtNivel(x), 'linea'); });
          if (activos.has('rango') && bd[q]) h += PM.ttFila(BANDA, 'Rango ' + ra.txt, corto(bd[q].min) + '–' + fmtNivel(bd[q].max), 'area');
          if (q > f.q) h += PM.ttPie(f.y + ' todavía no tiene dato de ' + nomP(q, true) + '.');
          return h;
        }),
        series: series
      };
    }

    // Crecimiento y Nivel: serie continua, línea en tinta, relleno por signo contra la referencia (0 o 100)
    var continua = { desde: 0 };
    // El nivel se mira en los últimos once años (como el deck): contra 2019, toda la expansión de los noventa
    // y dos mil saldría «por debajo» y taparía lo que importa. El crecimiento sí va con la serie entera.
    function inicioNivel() {
      var a = Math.max(P[primero()].a, fin().y - 11, P[primero()].a), i = idx[clave(a, 0)];
      return Math.max(primero(), i == null ? primero() : i);
    }
    function nombreContinua() {
      return estado.vista === 'crec' ? (tasa() ? 'Diferencia de ' + NPER() : 'Crecimiento de ' + NPER()) : tasa() ? 'Promedio de ' + NPER() : 'Nivel (' + estado.base + ' = 100)';
    }
    function opcContinua(el) {
      var dk = PM.dk(), cre = estado.vista === 'crec', i0 = cre ? primero() : inicioNivel(), iU = ultimo();
      var ref = cre ? 0 : referenciaNivel(), vals = [];
      for (var i = i0; i <= iU; i++) vals.push(r2(cre ? crec(i) : nivel(i)));
      var j0 = 0;   // la serie arranca en su primer valor: los primeros 12 o 24 meses no tienen con qué compararse
      while (j0 < vals.length - 1 && vals[j0] == null) j0++;
      vals = vals.slice(j0);
      var desde = continua.desde = i0 + j0, KV = K.slice(desde, iU + 1);
      var nombre = nombreContinua();
      // el relleno: una serie sólo de área, coloreada por el visualMap (rojo bajo la referencia, turquesa arriba)
      var area = PM.mezclar({ type: 'line', name: '_signo', data: vals, symbol: 'none', silent: true, connectNulls: false,
        lineStyle: { width: 0, opacity: 0 }, areaStyle: { origin: ref, opacity: 1 }, emphasis: { disabled: true }, z: 1 }, SUAVE);   // tintes oficiales, sólidos
      var linea = PM.linea(TINTA, { ancho: 2, extra: PM.mezclar({ name: nombre, data: vals, connectNulls: false, z: 5, emphasis: { focus: 'none', lineStyle: { width: 2 } } }, SUAVE) });
      var n = vals.length - 1;
      while (n > 0 && vals[n] == null) n--;
      var txt = cre ? fdif(vals[n]) : tasa() ? PM.pct(vals[n], 1) : PM.num(vals[n], 1);
      // de qué lado llega la línea: se mira unos puntos atrás (el rótulo ocupa ese tramo), no sólo el anterior
      var atras = Math.max(1, Math.round(vals.length * 0.04)), previo = null;
      for (var k = atras; k >= 1 && previo == null; k--) if (n - k >= 0) previo = vals[n - k];
      linea.markPoint = etiqueta(TINTA, txt, n, vals[n], vals.length - 1, previo, true);
      linea.markLine = cre ? PM.cero() : { silent: true, symbol: 'none', animation: false, data: [{ yAxis: r2(ref) }], label: { show: false },
        lineStyle: { color: dk ? 'rgba(226,232,240,.55)' : 'rgba(0,18,25,.5)', width: 1.25, type: 'dashed' } };
      PM.leyendaImagen = function () {
        var arriba = cre ? 'Sube' : tasa() ? 'Sobre su promedio' : 'Sobre ' + estado.base, abajo = cre ? 'Baja' : tasa() ? 'Bajo su promedio' : 'Bajo ' + estado.base;
        return [{ name: nombre, itemStyle: { color: PM.col(TINTA) }, forma: 'linea' },
          { name: arriba, itemStyle: { color: PM.tinte(true) } }, { name: abajo, itemStyle: { color: PM.tinte(false) } }];
      };
      PM.tablaDatos = function () {
        var cols = ['Período', nombre + (cre ? (tasa() ? ' (pp)' : ' (%)') : ''), (flujo() ? 'Suma de ' : 'Promedio de ') + NPER(), serie().corto + ' del período'];
        return { cols: cols, filas: KV.map(function (k, j) { var i = desde + j; return [k, vals[j], r2(doce(i)), r2(S()[i])]; }) };
      };
      var conDato = vals.filter(function (x) { return x != null; });
      var lo = Math.min(ref, Math.min.apply(null, conDato)) - 1, hi = Math.max(ref, Math.max.apply(null, conDato)) + 1;
      var ejeV = cre
        ? PM.ejeY({ unidad: tasa() ? 'pp' : '%', fmt: tasa() ? PM.tick : 'pct', extra: NOMBRE_EJE })
        : tasa() ? PM.ejeY({ unidad: '% de la capacidad', fmt: 'pct', extra: PM.mezclar({ scale: true }, NOMBRE_EJE) })
        : PM.ejeY({ unidad: estado.base + ' = 100', extra: PM.mezclar({ scale: true }, NOMBRE_EJE) });
      return {
        grid: PM.grid(),
        // tramos FINITOS: con dos tramos abiertos ECharts no arma paradas de color y el render aborta («coord»)
        visualMap: [{ type: 'piecewise', show: false, seriesIndex: 0, dimension: 1,
          pieces: [{ gte: lo, lt: ref, color: PM.tinte(false) }, { gte: ref, lte: hi, color: PM.tinte(true) }] }],
        xAxis: PM.ejeTiempo(KV, el),
        yAxis: ejeV,
        tooltip: PM.tooltip(function (ps) {
          if (!ps.length) return '';
          var j = ps[0].dataIndex, i = desde + j, x = vals[j];
          panel(i);
          var h = PM.ttTitulo(PM.fecha(KV[j]), cre ? 'Últimos ' + NPER() + ' vs los ' + NPER() + ' anteriores' : tasa() ? 'Promedio de los últimos ' + NPER() : 'Últimos ' + NPER() + ', ' + estado.base + ' = 100');
          if (x != null) h += PM.ttFila(TINTA, nombre, cre ? fdif(x) : tasa() ? PM.pct(x, 1) : PM.num(x, 1), 'linea');
          h += PM.ttFila(null, (flujo() ? 'Suma de ' : 'Promedio de ') + NPER(), fmtNivel(doce(i)));
          h += PM.ttFila(null, serie().corto + ' del ' + (frec === 12 ? 'mes' : 'trimestre'), fmtNivel(S()[i]));
          return h;
        }),
        series: [area, linea]
      };
    }

    // ── Panel que sigue al cursor: el período bajo el cursor y la conclusión calculada ──
    function umbral(x) {
      var um = serie().umbrales || cfg.umbrales, a = um[0], b = um[1], ab = Math.abs(x), un = tasa() ? ' puntos' : '%';
      if (ab < a) return 'prácticamente sin cambio (menos de ' + PM.num(a, 0) + un + ')';
      if (ab < b) return (x > 0 ? 'un aumento moderado' : 'una caída moderada') + ' (entre ' + PM.num(a, 0) + ' y ' + PM.num(b, 0) + un + ')';
      return (x > 0 ? 'un aumento fuerte' : 'una caída fuerte') + ' (más de ' + PM.num(b, 0) + un + ')';
    }
    function masMenos(x) { return tasa() ? (x >= 0 ? PM.num(x, 1) + ' puntos más' : PM.num(-x, 1) + ' puntos menos') : x >= 0 ? PM.pct(x, 1) + ' más' : PM.pct(-x, 1) + ' menos'; }
    function panel(i) {
      if (i == null || !D) return;
      estado.i = i;
      if (estado.vista === 'anio') panelAnio(Math.max(0, Math.min(frec - 1, i))); else panelContinua(Math.max(primero(), Math.min(ultimo(), i)));
    }
    function posicion() {                  // dónde queda el tramo del año en curso entre los mismos tramos de toda la serie
      var f = fin(), hist = [];
      for (var y = P[primero()].a; y <= f.y; y++) { var x = agregado(y, f.q); if (x != null) hist.push({ y: y, v: x }); }
      var n = hist.length, cur = hist[n - 1], prev = hist[n - 2], j, t = tramoSin(f.q);
      if (!cur || cur.y !== f.y || !prev) return '';
      if (cur.v >= prev.v) {
        for (j = n - 2; j >= 0 && hist[j].v < cur.v; j--);
        return j < 0 ? ' Es el ' + t + ' más alto de la serie (desde ' + hist[0].y + ').' : ' Es el ' + t + ' más alto desde ' + hist[j].y + '.';
      }
      for (j = n - 2; j >= 0 && hist[j].v > cur.v; j--);
      return j < 0 ? ' Es el ' + t + ' más bajo de la serie (desde ' + hist[0].y + ').' : ' Es el ' + t + ' más bajo desde ' + hist[j].y + '.';
    }
    function lecturaAnio() {
      var f = fin(), a = agregado(f.y, f.q), b = agregado(f.y - 1, f.q), va = dif(a, b), sr = serie();
      if (va == null) return '';
      return 'En ' + tramoLargo(f.y, f.q) + ' ' + sr.acum.replace('{v}', '<strong>' + fmtNivel(a) + '</strong>') + ', <strong>' + masMenos(va) +
        '</strong> que ' + (f.q === frec - 1 ? 'en ' : 'en el mismo tramo de ') + (f.y - 1) + ': ' + umbral(va) + '.' + posicion();
    }
    function panelAnio(q) {
      var f = fin(), ra = rangoAnios(), bd = banda()[q], html = '';
      document.getElementById('panel-fecha').textContent = nomP(q, true);
      var cur = valorAnio(f.y, q), prev = valorAnio(f.y - 1, q), prev2 = valorAnio(f.y - 2, q);
      var etq = function (y) { return estado.medida === 'mes' ? fechaP(y, q) : tramoCorto(y, q); };
      html += cur == null
        ? PM.pb(ROJO, String(f.y), '—', 'Todavía sin dato de ' + nomP(q, true) + ': ' + f.y + ' llega hasta ' + nomP(f.q, true) + '.')
        : PM.pb(ROJO, etq(f.y), fmtNivel(cur), prev == null ? '' : fdif(dif(cur, prev)) + ' vs ' + etq(f.y - 1));
      if (prev != null) html += PM.pb(PREVIO, etq(f.y - 1), fmtNivel(prev), prev2 == null ? '' : fdif(dif(prev, prev2)) + ' vs ' + etq(f.y - 2));
      if (bd) {
        html += PM.pb(BANDA, 'Rango ' + ra.txt, corto(bd.min) + '–' + fmtNivel(bd.max), 'Promedio de esos años: ' + fmtNivel(bd.prom));
        var filas = '';
        ra.anios.forEach(function (y) { var x = valorAnio(y, q); if (x != null) filas += '<div class="anio-f"><span class="anio-a">' + y + '</span><span class="anio-v">' + fmtNivel(x) + '</span></div>'; });
        html += '<div class="anios">' + filas + '</div>';
      }
      // conclusión: el período bajo el cursor contra el año anterior y contra lo normal; después, el año en curso
      var txt = '';
      if (cur != null) {
        var d1 = dif(cur, prev), dn = bd ? dif(cur, bd.prom) : null;
        txt += 'En ' + (estado.medida === 'mes' ? (frec === 4 ? 'el ' : '') + nomP(q, true) + ' de ' + f.y : tramoLargo(f.y, q)) + ': ' + (d1 == null ? '' : '<strong>' + masMenos(d1) + '</strong> que en ' + (f.y - 1)) +
          (bd ? (cur > bd.max ? ' y <strong>por encima de todo el rango</strong> de ' + ra.txt
            : cur < bd.min ? ' y <strong>por debajo de todo el rango</strong> de ' + ra.txt
            : ', dentro del rango de ' + ra.txt + ' (' + fdif(dn, 0) + ' frente a su promedio)') : '') + '. ';
      } else txt += f.y + ' todavía no llega a ' + nomP(q, true) + '. ';
      txt += lecturaAnio();
      document.getElementById('panel').innerHTML = html + PM.ctx(txt);
    }
    function racha(i, fn) {                // períodos seguidos con el mismo signo hasta i
      var x = fn(i); if (x == null) return 0;
      var n = 0, s = x >= 0;
      for (var j = i; j >= 0; j--) { var y = fn(j); if (y == null || (y >= 0) !== s) break; n++; }
      return n;
    }
    function panelContinua(i) {
      var cre = estado.vista === 'crec', s = S(), h = hace(i), html = '', per = frec === 12 ? 'meses' : 'trimestres';
      document.getElementById('panel-fecha').textContent = PM.fecha(K[i]);
      var c = crec(i), d = doce(i), nv = nivel(i), ref = cre ? 0 : referenciaNivel();
      var dm = dif(s[i], h == null ? null : s[h]);
      if (cre) html += PM.pb(TINTA, tasa() ? 'Diferencia de ' + NPER() : 'Crecimiento de ' + NPER(), fdif(c), (flujo() ? 'Suma' : 'Promedio') + ' de ' + NPER() + ': ' + fmtNivel(d));
      else html += PM.pb(TINTA, tasa() ? 'Promedio de ' + NPER() : 'Nivel (' + estado.base + ' = 100)', tasa() ? PM.pct(nv, 1) : PM.num(nv, 1),
        tasa() ? 'Promedio desde ' + P[primero()].a + ': ' + PM.pct(ref, 1) : (flujo() ? 'Suma' : 'Promedio') + ' de ' + NPER() + ': ' + fmtNivel(d));
      html += PM.pb(CTX_TIP, serie().corto + ' del ' + (frec === 12 ? 'mes' : 'trimestre'), fmtNivel(s[i]), dm == null ? '' : fdif(dm) + ' vs ' + PM.fecha(K[h]));
      var txt;
      if (cre && c != null) {
        var rn = racha(i, crec);
        txt = 'A ' + PM.fecha(K[i]) + ', los últimos ' + NPER() + ' ' + (flujo() ? 'suman ' : 'promedian ') + '<strong>' + fmtNivel(d) + '</strong>: <strong>' + masMenos(c) + '</strong> que los ' + NPER() + ' anteriores, ' + umbral(c) + '. ' +
          (rn > 1 ? 'Lleva <strong>' + rn + ' ' + per + ' seguidos ' + (c >= 0 ? 'en alza' : 'en baja') + '</strong>.' : 'Es el primer ' + (frec === 12 ? 'mes' : 'trimestre') + ' ' + (c >= 0 ? 'en alza' : 'en baja') + ' después de uno de signo contrario.');
      } else if (!cre && nv != null) {
        var mx = null;
        for (var j = inicioNivel(); j <= ultimo(); j++) { var x = nivel(j); if (x != null && (!mx || x > mx.v)) mx = { v: x, i: j }; }
        txt = tasa()
          ? 'A ' + PM.fecha(K[i]) + ', la industria usó en promedio el <strong>' + PM.pct(nv, 1) + '</strong> de su capacidad en los últimos ' + NPER() + ', ' +
            (Math.abs(nv - ref) < 1 ? 'en línea con' : nv > ref ? PM.num(nv - ref, 1) + ' puntos <strong>por encima</strong> de' : PM.num(ref - nv, 1) + ' puntos <strong>por debajo</strong> de') + ' su promedio desde ' + P[primero()].a + ' (' + PM.pct(ref, 1) + ').'
          : 'A ' + PM.fecha(K[i]) + ', el volumen de los últimos ' + NPER() + ' está ' + (Math.abs(nv - 100) < 0.05 ? 'igual que en ' + estado.base
            : '<strong>' + PM.pct(Math.abs(nv - 100), 1) + (nv > 100 ? ' por encima' : ' por debajo') + '</strong> del de ' + estado.base) + '.';
        if (mx) txt += ' Máximo desde ' + P[inicioNivel()].a + ': ' + (tasa() ? PM.pct(mx.v, 1) : PM.num(mx.v, 1)) + ' (' + PM.fecha(K[mx.i]) + ')' + (mx.i === i ? ', este mismo período.' : '.');
      } else txt = 'Para este período todavía no hay ' + NPER() + ' anteriores con qué comparar.';
      document.getElementById('panel').innerHTML = html + PM.ctx(txt);
    }

    // ── Preguntas ────────────────────────────────────────────────────────────
    function ayuda() {
      return { D: D, fin: fin, prelim: D.metadata.preliminar_desde,
        primeroAnio: function (k) { var s = D.series[k]; for (var i = 0; i < s.length; i++) if (s[i] != null) return P[i].a; return null; } };
    }
    function comoLeer() {
      var f = fin(), ra = rangoAnios(), per = frec === 12 ? 'mes' : 'trimestre';
      return ['¿Cómo leer las tres vistas?',
        '<p><strong>Año por año</strong>: cada línea es un año y el eje va de ' + (frec === 12 ? 'enero a diciembre' : 'T1 a T4') + '. ' + f.y + ' va en rojo' +
          (f.q === frec - 1 ? '' : ' y llega hasta ' + nomP(f.q, true)) + '; ' + (f.y - 1) + ', en turquesa; la franja gris es el rango de ' + ra.txt +
          ' (el valor más bajo y el más alto de cada ' + per + '): arriba de la franja es más de lo normal, abajo, menos. Las cifras comparan siempre el <strong>mismo tramo</strong> del año.</p>' +
        '<p><strong>Crecimiento</strong>: compara los últimos ' + NPER() + ' con los ' + NPER() + ' anteriores. Al tomar un año entero se cancelan la estacionalidad y los ' + per + (frec === 12 ? 'es' : 's') + ' atípicos; turquesa cuando sube, rojo cuando baja.</p>' +
        '<p><strong>Nivel</strong>: los últimos ' + NPER() + ' con el año base elegido = 100' + (tasa() ? '; en el uso de la capacidad, la referencia es su promedio histórico.' : ': 90 es 10% menos que en el año base.') + '</p>'];
    }
  }

  window.IND = { armar: armar, cifra: cifra };
})();
