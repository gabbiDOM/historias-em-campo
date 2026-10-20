/* Histórias em Campo — lê dados/dados.xlsx e monta o portal no navegador. */
(() => {
  'use strict';

  const DATA_URL = 'dados/dados.xlsx';
  const VIDEOS_URL = 'dados/videos.json'; // gerado pela automação do GitHub a partir do feed do canal
  const FOTOS_URL = 'dados/fotos.json'; // miniaturas da EA por nome; a coluna Foto do Elenco tem prioridade
  const $app = document.getElementById('app');
  const $avisos = document.getElementById('avisos');

  // ---------- utilidades ----------
  const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const norm = s => String(s ?? '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/[^a-z0-9]/g, '');
  const num = v => {
    if (v === '' || v == null) return null;
    const n = typeof v === 'number' ? v : Number(String(v).replace(',', '.'));
    return Number.isFinite(n) ? n : null;
  };
  const isOwnGoal = name => norm(name) === 'golcontra';
  const sim = v => ['sim', 's', 'yes', 'true', '1', 'x'].includes(norm(v));
  // "26/27", "26-27" e "2026/2027" viram "2026/27", para todas as abas usarem a mesma temporada.
  const season = v => {
    const s = String(v ?? '').trim();
    const m = s.match(/^(?:20)?(\d{2})\s*[\/-]\s*(?:20)?(\d{2})$/);
    return m ? `20${m[1]}/${m[2]}` : s;
  };
  const plural =(n, um, varios) => `${n} ${n === 1 ? um : varios}`;
  const ordinal = n => `${n}º`;

  function fmtDate(v) {
    if (typeof v === 'number') {
      const d = new Date(Math.round((v - 25569) * 86400000));
      return d.toLocaleDateString('pt-BR', { timeZone: 'UTC' });
    }
    return String(v || '');
  }
  // Aceita 23, "23'", "45+2" ou "90 + 3"; ordem coloca acréscimos depois do minuto base.
  function parseMinute(v) {
    const m = String(v ?? '').replace(/\s|'/g, '').match(/^(\d{1,3})(?:\+(\d{1,2}))?$/);
    if (!m) return null;
    const base = Number(m[1]), extra = m[2] ? Number(m[2]) : 0;
    return { texto: extra ? `${base}+${extra}'` : `${base}'`, ordem: base + extra / 100 };
  }
  function ytId(url) {
    const m = String(url || '').match(/(?:v=|youtu\.be\/|shorts\/|embed\/|live\/)([\w-]{11})/);
    return m ? m[1] : null;
  }

  // Lê uma aba como lista de objetos, com chaves normalizadas (sem acento, minúsculas).
  function sheet(wb, name) {
    const key = wb.SheetNames.find(n => norm(n) === norm(name));
    if (!key) return [];
    return XLSX.utils.sheet_to_json(wb.Sheets[key], { defval: '' })
      .map(r => {
        const o = {};
        for (const k in r) o[norm(k)] = typeof r[k] === 'string' ? r[k].trim() : r[k];
        return o;
      })
      .filter(o => Object.values(o).some(v => v !== ''));
  }

  // Completa os episódios com os vídeos do canal: título, data e link vêm do YouTube
  // quando a planilha deixa esses campos vazios. Vídeo sem linha na planilha vira episódio novo.
  // Episódio 0 é a INTRO: aparece na lista antes do EP 01, mas não tem jogos nem entra na contagem.
  const isIntro = e => e.n === 0;
  const cleanTitle = t => String(t || '').split('|')[0].trim() || String(t || '');
  function mergeVideos(episodes, videos, temporadaPadrao) {
    if (!videos || !Array.isArray(videos.videos)) return episodes;
    const byN = new Map(episodes.map(e => [e.n, e]));
    videos.videos.forEach(v => {
      if (v.episodio == null || !v.id) return;
      let e = byN.get(v.episodio);
      if (!e) {
        e = { n: v.episodio, temporada: temporadaPadrao, titulo: '', data: '', link: '', resumo: '' };
        byN.set(e.n, e);
      }
      if (!e.link) e.link = `https://www.youtube.com/watch?v=${v.id}`;
      if (!e.titulo) e.titulo = isIntro(e) ? cleanTitle(v.titulo).replace(/^hist[oó]rias em campo\s*[:\-–]\s*/i, '') : cleanTitle(v.titulo);
      if (!e.data && v.publicado) e.data = new Date(v.publicado).toLocaleDateString('pt-BR');
    });
    return [...byN.values()].sort((a, b) => a.n - b.n);
  }

  // ---------- modelo ----------
  function build(wb, videos, fotos) {
    const cfg = {};
    sheet(wb, 'Clube').forEach(r => { if (r.chave) cfg[norm(r.chave)] = r.valor; });

    const clube = String(cfg.clube || 'Meu clube');
    const liga = String(cfg.competicaoliga || '');

    const episodes = mergeVideos(sheet(wb, 'Episodios')
      .map(r => ({ n: num(r.episodio), temporada: season(r.temporada), titulo: cleanTitle(r.titulo), data: fmtDate(r.data), link: String(r.linkyoutube || ''), resumo: String(r.resumo || '') }))
      .filter(e => e.n != null)
      .sort((a, b) => a.n - b.n), videos, season(cfg.temporadaatual));
    const epSeason = new Map(episodes.map(e => [e.n, e.temporada]));

    const games = sheet(wb, 'Jogos')
      .map(r => {
        const gp = num(r.golspro), gc = num(r.golscontra), ep = num(r.episodio);
        return {
          id: num(r.jogo), ep, temporada: season(r.temporada) || epSeason.get(ep) || '',
          comp: String(r.competicao || ''), rodada: r.rodada, adv: String(r.adversario || ''),
          casa: norm(r.local) !== 'fora', gp, gc,
          res: gp == null || gc == null ? null : gp > gc ? 'V' : gp === gc ? 'E' : 'D',
          gols: []
        };
      })
      .filter(g => g.id != null && g.res)
      .sort((a, b) => a.id - b.id);

    const byId = new Map(games.map(g => [g.id, g]));
    const warnings = [];
    sheet(wb, 'Gols').forEach(r => {
      const g = byId.get(num(r.jogo));
      if (!g) { if (r.autor) warnings.push(`Gol de ${r.autor} aponta para o jogo ${r.jogo}, que não existe na aba Jogos.`); return; }
      if (r.autor) g.gols.push({ autor: String(r.autor), assist: String(r.assistencia || ''), minuto: parseMinute(r.minuto) });
    });
    games.forEach(g => {
      g.gols.sort((a, b) => (a.minuto ? a.minuto.ordem : 999) - (b.minuto ? b.minuto.ordem : 999));
      if (g.gols.length !== g.gp) warnings.push(`Jogo ${g.id} (${g.adv}): placar com ${plural(g.gp, 'gol', 'gols')}, mas ${g.gols.length} registrado(s) na aba Gols.`);
    });

    const tables = new Map();
    sheet(wb, 'Classificacao').forEach(r => {
      const ep = num(r.episodio);
      if (ep == null || !r.time) return;
      const row = { pos: num(r.posicao), time: String(r.time), j: num(r.j) || 0, v: num(r.v) || 0, e: num(r.e) || 0, d: num(r.d) || 0, gp: num(r.gp) || 0, gc: num(r.gc) || 0 };
      row.pts = row.v * 3 + row.e;
      row.sg = row.gp - row.gc;
      if (!tables.has(ep)) tables.set(ep, []);
      tables.get(ep).push(row);
    });
    tables.forEach(t => t.sort((a, b) => (a.pos ?? 99) - (b.pos ?? 99)));
    const tableEps = [...tables.keys()].sort((a, b) => a - b);

    const elenco = new Map();
    // Chegada/Saida são números de episódio; o Sheets às vezes converte "1/10" em data (número grande).
    const epNum = (v, jogador, coluna) => {
      const n = num(v);
      if (n == null) return null;
      if (n >= 1000) { warnings.push(`Elenco: ${jogador} tem uma data em ${coluna}. Use só o número do episódio (ex.: 3).`); return null; }
      return n;
    };
    // Situacao: "Emprestado" ou saída ("Vendido", "Saiu", "Dispensado"...). Saida preenchida sem situação = saiu.
    const situacao = r => {
      const n = norm(r.situacao), label = String(r.situacao || '').trim();
      if (n.startsWith('empre')) return { situacao: 'emprestado', situacaoLabel: label };
      if (/^(vend|saiu|transf|dispens|liber|aposent|rescin)/.test(n) || (!n && num(r.saida) != null)) return { situacao: 'saiu', situacaoLabel: label || 'Saiu' };
      return { situacao: '', situacaoLabel: '' };
    };
    const fotosMap = new Map(Object.entries(fotos || {}).map(([k, v]) => [norm(k), String(v)]));
    // Origem: "Base" (cria da base), "Contratação" (reforço) ou vazio/"Elenco inicial".
    const origemTipo = v => { const n = norm(v); return n.startsWith('base') ? 'base' : /^(contrat|reforc|emprest)/.test(n) ? 'reforco' : ''; };
    sheet(wb, 'Elenco').forEach(r => {
      if (r.jogador) elenco.set(norm(r.jogador), {
        nome: String(r.jogador), pos: String(r.posicao || ''), nac: String(r.nacionalidade || ''),
        origem: origemTipo(r.origem), chegada: epNum(r.chegada, r.jogador, 'Chegada'), saida: epNum(r.saida, r.jogador, 'Saida'),
        ...situacao(r), foto: String(r.foto || '').trim() || fotosMap.get(norm(r.jogador)) || ''
      });
    });

    const recordes = sheet(wb, 'Recordes').map(r => ({
      chave: norm(r.chave), titulo: String(r.recorde || r.chave), valor: num(r.valor),
      detentor: String(r.detentor || ''), quando: String(r.quando || ''), obs: String(r.observacao || ''), confirmado: sim(r.confirmado)
    }));
    const metas = sheet(wb, 'Metas').map(r => ({ meta: String(r.meta || ''), status: String(r.status || 'Em andamento'), obs: String(r.observacao || '') })).filter(m => m.meta);
    const titulos = sheet(wb, 'Titulos').map(r => ({ temporada: season(r.temporada), comp: String(r.competicao || ''), obs: String(r.observacao || '') })).filter(t => t.comp);

    const seasons = [...new Set([...episodes.map(e => e.temporada), ...games.map(g => g.temporada)].filter(Boolean))];
    const temporadaAtual = season(cfg.temporadaatual) || seasons[seasons.length - 1] || '';

    return {
      cfg, clube, liga, episodes, games, tables, tableEps, elenco, fotosMap, recordes, metas, titulos, seasons, temporadaAtual, warnings,
      // A tabela pode ter só os primeiros colocados; o gráfico usa o total da liga.
      totalTimes: num(cfg.totaltimes) ?? 18,
      zonas: {
        acesso: num(cfg.zonaacesso) ?? 2,
        playoff: num(cfg.zonaplayoffacesso) ?? 3,
        playoffReb: num(cfg.zonaplayoffrebaixamento) ?? 16,
        reb: num(cfg.zonarebaixamento) ?? 17
      }
    };
  }

  // ---------- estatísticas ----------
  function playerStats(M, games) {
    const m = new Map();
    const get = name => {
      const k = norm(name);
      if (!m.has(k)) m.set(k, { ...(M.elenco.get(k) || { nome: name, pos: '', nac: '', origem: '', situacao: '', foto: M.fotosMap.get(k) || '' }), g: 0, a: 0 });
      return m.get(k);
    };
    games.forEach(g => g.gols.forEach(x => {
      if (!isOwnGoal(x.autor)) get(x.autor).g++;
      if (x.assist) get(x.assist).a++;
    }));
    return m;
  }

  function streak(games, ok) {
    let cur = 0, max = 0, start = 0, best = null;
    games.forEach((g, i) => {
      if (ok(g)) {
        if (cur === 0) start = i;
        cur++;
        if (cur > max) { max = cur; best = [games[start], games[i]]; }
      } else cur = 0;
    });
    return { max, cur, best };
  }

  const scoreline = (M, g) => g.casa ? `${M.clube} ${g.gp} x ${g.gc} ${g.adv}` : `${g.adv} ${g.gc} x ${g.gp} ${M.clube}`;
  const rangeLabel = best => best ? (best[0].ep === best[1].ep ? `EP ${best[0].ep}` : `EP ${best[0].ep} ao ${best[1].ep}`) : '';

  function topBy(map, key) {
    let top = null;
    map.forEach(p => { if (p[key] > 0 && (!top || p[key] > top[key])) top = p; });
    return top;
  }

  // Valor atual da carreira para cada chave de recorde.
  function careerRecords(M) {
    const out = {};
    const w = streak(M.games, g => g.res === 'V');
    out.seqvitorias = { valor: w.max, label: rangeLabel(w.best), atual: w.cur };
    const u = streak(M.games, g => g.res !== 'D');
    out.invencibilidade = { valor: u.max, label: rangeLabel(u.best), atual: u.cur };

    let goleada = null;
    M.games.forEach(g => {
      if (g.res !== 'V') return;
      const s = g.gp - g.gc;
      if (!goleada || s > goleada.gp - goleada.gc || (s === goleada.gp - goleada.gc && g.gp > goleada.gp)) goleada = g;
    });
    out.maiorgoleada = { valor: goleada ? goleada.gp - goleada.gc : 0, label: goleada ? `${scoreline(M, goleada)} · EP ${goleada.ep}` : '' };

    let bestG = null, bestA = null;
    M.seasons.forEach(s => {
      const st = playerStats(M, M.games.filter(g => g.temporada === s));
      const tg = topBy(st, 'g'), ta = topBy(st, 'a');
      if (tg && (!bestG || tg.g > bestG.p.g)) bestG = { p: tg, s };
      if (ta && (!bestA || ta.a > bestA.p.a)) bestA = { p: ta, s };
    });
    out.golstemporada = { valor: bestG ? bestG.p.g : 0, label: bestG ? `${bestG.p.nome} · ${bestG.s}` : '' };
    out.assisttemporada = { valor: bestA ? bestA.p.a : 0, label: bestA ? `${bestA.p.nome} · ${bestA.s}` : '' };

    const all = topBy(playerStats(M, M.games), 'g');
    out.artilheirohistorico = { valor: all ? all.g : 0, label: all ? all.nome : '' };
    out.titulos = { valor: M.titulos.length, label: M.titulos.length ? M.titulos.map(t => t.comp).join(', ') : 'Nenhum ainda' };
    return out;
  }

  function recordStatus(real, carreira) {
    if (real == null) return { pct: 0, texto: 'Recorde real a definir na planilha', quebrado: false };
    if (carreira > real) return { pct: 100, texto: 'Recorde quebrado!', quebrado: true };
    if (carreira === real) return { pct: 100, texto: 'Recorde igualado! Falta 1 para quebrar', quebrado: false };
    const falta = real - carreira + 1;
    return { pct: real > 0 ? Math.round(carreira / real * 100) : 0, texto: `Faltam ${falta} para quebrar`, quebrado: false };
  }

  function clubRow(M, ep) {
    const t = M.tables.get(ep);
    return t ? t.find(r => norm(r.time) === norm(M.clube)) : null;
  }
  function zoneClass(M, pos) {
    const z = M.zonas;
    if (pos <= z.acesso) return 'z-acesso';
    if (pos === z.playoff) return 'z-playoff';
    if (pos >= z.reb) return 'z-reb';
    if (pos === z.playoffReb) return 'z-playoff-reb';
    return '';
  }

  // ---------- componentes ----------
  function gameCard(M, g) {
    const home = g.casa ? M.clube : g.adv, away = g.casa ? g.adv : M.clube;
    const hs = g.casa ? g.gp : g.gc, as = g.casa ? g.gc : g.gp;
    const rodada = g.rodada === '' ? '' : (num(g.rodada) != null ? `Rodada ${g.rodada}` : esc(g.rodada));
    const gols = g.gols.length ? `<ul class="gols-lista">${g.gols.map(x => `<li><span class="ic-gol"></span>${x.minuto ? `<span class="min">${x.minuto.texto}</span>` : ''}${esc(x.autor)}${x.assist ? ` <small>assist. ${esc(x.assist)}</small>` : ''}</li>`).join('')}</ul>` : '';
    return `<div class="jogo">
      <div class="jogo-meta">${esc(g.comp)}${rodada ? ' · ' + rodada : ''}</div>
      <div class="placar">
        <span class="time ${g.casa ? 'eu' : ''}">${esc(home)}</span>
        <span class="gols">${hs}<i> x </i>${as}</span>
        <span class="time ${g.casa ? '' : 'eu'}">${esc(away)}</span>
        <span class="res res-${g.res}" title="${{ V: 'Vitória', E: 'Empate', D: 'Derrota' }[g.res]}">${g.res}</span>
      </div>${gols}
    </div>`;
  }

  function episodeCard(M, e) {
    const id = ytId(e.link);
    const intro = isIntro(e);
    const pad = String(e.n).padStart(2, '0');
    const thumb = id
      ? `<a class="ep-thumb" href="${esc(e.link)}" target="_blank" rel="noopener"><img src="https://i.ytimg.com/vi/${id}/hqdefault.jpg" alt="Miniatura ${intro ? 'da intro' : `do episódio ${e.n}`}" loading="lazy"><span class="play">▶ Assistir</span></a>`
      : `<div class="ep-thumb"><span class="ep-num">${intro ? 'INTRO' : `EP ${pad}`}</span>${e.link ? '' : '<span class="play">Vídeo em breve</span>'}</div>`;
    const games = M.games.filter(g => g.ep === e.n);
    const me = clubRow(M, e.n);
    const prevEp = M.tableEps.filter(x => x < e.n).pop();
    const prev = prevEp != null ? clubRow(M, prevEp) : null;
    const delta = me && prev ? prev.pos - me.pos : 0;
    const pos = me ? `<a class="ep-pos" href="#tabela/${e.n}">Fim do episódio: <b>${ordinal(me.pos)} lugar</b> · ${me.pts} pts${delta ? ` <span class="delta ${delta > 0 ? 'sobe' : 'desce'}">${delta > 0 ? '▲' : '▼'}${Math.abs(delta)}</span>` : ''}</a>` : '';
    return `<article class="card ep" id="ep-${e.n}">
      ${thumb}
      <div>
        <div class="ep-cabecalho"><small>${intro ? 'INTRO' : `EPISÓDIO ${pad}`}${e.temporada ? ' · ' + esc(e.temporada) : ''}</small><h3>${esc(e.titulo || (intro ? 'Intro da série' : `Episódio ${pad}`))}</h3></div>
        ${e.data ? `<div class="ep-data">Lançado em ${esc(e.data)}</div>` : ''}
        ${e.resumo ? `<p class="ep-resumo">${esc(e.resumo)}</p>` : ''}
        ${intro && !games.length ? '' : `<div class="jogos">${games.map(g => gameCard(M, g)).join('') || '<p class="mudo">Sem jogos registrados.</p>'}</div>`}
        ${pos}
      </div>
    </article>`;
  }

  function positionChart(M) {
    const pts = M.tableEps.map(ep => ({ ep, row: clubRow(M, ep), n: M.tables.get(ep).length })).filter(p => p.row);
    if (!pts.length) return '<p class="mudo">A evolução aparece quando houver classificação na planilha.</p>';
    const nTeams = Math.max(M.totalTimes, ...pts.map(p => p.n), ...pts.map(p => p.row.pos));
    const W = 640, H = 240, l = 40, r = 20, t = 18, b = 32;
    const x = i => pts.length === 1 ? (l + W - r) / 2 : l + i * (W - l - r) / (pts.length - 1);
    const y = pos => t + (pos - 1) * (H - t - b) / Math.max(1, nTeams - 1);
    const z = M.zonas;
    const ticks = [1, z.acesso, z.reb, nTeams].filter((v, i, a) => v >= 1 && v <= nTeams && a.indexOf(v) === i);
    const path = pts.map((p, i) => `${i ? 'L' : 'M'}${x(i).toFixed(1)},${y(p.row.pos).toFixed(1)}`).join(' ');
    return `<svg viewBox="0 0 ${W} ${H}" role="img" aria-label="Posição do ${esc(M.clube)} ao final de cada episódio">
      <rect class="faixa-up" x="${l}" y="${t - 8}" width="${W - l - r}" height="${y(z.acesso) - t + 16}" rx="6"/>
      <rect class="faixa-down" x="${l}" y="${y(z.reb) - 8}" width="${W - l - r}" height="${y(nTeams) - y(z.reb) + 16}" rx="6"/>
      ${ticks.map(v => `<line class="grade" x1="${l}" x2="${W - r}" y1="${y(v)}" y2="${y(v)}"/><text class="eixo" x="${l - 10}" y="${y(v) + 4}" text-anchor="end">${v}º</text>`).join('')}
      <path class="linha" d="${path}"/>
      ${pts.map((p, i) => `<circle class="ponto" cx="${x(i)}" cy="${y(p.row.pos)}" r="6"><title>EP ${p.ep}: ${p.row.pos}º · ${p.row.pts} pts</title></circle>
        <text class="rotulo" x="${x(i)}" y="${y(p.row.pos) + (p.row.pos > nTeams / 2 ? -14 : 22)}" text-anchor="middle">${p.row.pos}º</text>
        <text class="eixo" x="${x(i)}" y="${H - 8}" text-anchor="middle">EP ${p.ep}</text>`).join('')}
    </svg>`;
  }

  function recordCard(M, rec, car) {
    const c = car[rec.chave];
    if (!c) return '';
    const st = recordStatus(rec.valor, c.valor);
    const real = [rec.detentor, rec.quando].filter(Boolean).join(' · ');
    const atual = c.atual != null && rec.chave !== 'titulos' ? `<span class="mudo">Sequência atual: ${c.atual}</span>` : '';
    return `<article class="card recorde ${st.quebrado ? 'quebrado' : ''}">
      <header><h3>${esc(rec.titulo)}</h3>${st.quebrado ? '<span class="tag ok">quebrado</span>' : rec.confirmado ? '' : '<span class="tag" title="Valor ainda não conferido">a confirmar</span>'}</header>
      <div class="rec-numeros">
        <div><small>Recorde do clube</small><strong>${rec.valor ?? '—'}</strong><span>${esc(real || rec.obs)}</span></div>
        <div><small>Na carreira</small><strong>${c.valor}</strong><span>${esc(c.label || '—')}</span></div>
      </div>
      <div class="barra" role="progressbar" aria-valuemin="0" aria-valuemax="100" aria-valuenow="${st.pct}"><span style="width:${st.pct}%"></span></div>
      <p class="rec-status">${esc(st.texto)}</p>${atual}
    </article>`;
  }

  // ---------- páginas ----------
  function pageInicio(M) {
    const season = M.games.filter(g => g.temporada === M.temporadaAtual);
    const v = season.filter(g => g.res === 'V').length, e = season.filter(g => g.res === 'E').length, d = season.filter(g => g.res === 'D').length;
    const gp = season.reduce((s, g) => s + g.gp, 0), gc = season.reduce((s, g) => s + g.gc, 0);
    const lastTableEp = M.tableEps[M.tableEps.length - 1];
    const me = lastTableEp != null ? clubRow(M, lastTableEp) : null;
    const st = playerStats(M, season);
    const art = topBy(st, 'g'), gar = topBy(st, 'a');
    const forma = season.slice(-5).map(g => `<span class="res res-${g.res}" title="${esc(scoreline(M, g))}">${g.res}</span>`).join('');
    const last = M.episodes[M.episodes.length - 1];

    const car = careerRecords(M);
    const destaques = M.recordes
      .filter(r => r.valor != null && car[r.chave])
      .map(r => ({ r, pct: recordStatus(r.valor, car[r.chave].valor).pct }))
      .sort((a, b) => b.pct - a.pct)
      .slice(0, 3);

    const link = M.cfg.linkcanal ? `<a class="botao" href="${esc(M.cfg.linkcanal)}" target="_blank" rel="noopener">▶ Assistir no YouTube</a>` : '';
    return `
      <section class="hero">
        <h1>${esc(M.cfg.nomeportal || 'Histórias em Campo')}</h1>
        <p>${esc(M.clube)} · Temporada ${esc(M.temporadaAtual)}${M.liga ? ' · ' + esc(M.liga) : ''}. Episódios, resultados, artilharia e os recordes que a série quer quebrar.</p>
        <div class="botoes">${link}<a class="botao secundario" href="#episodios">Ver episódios</a></div>
      </section>

      <section class="secao grid grid-4">
        <div class="card stat"><small>Posição atual</small><strong>${me ? ordinal(me.pos) : '—'}</strong><span>${me ? `${me.pts} pts · ${esc(M.liga)}` : 'sem tabela'}</span></div>
        <div class="card stat"><small>Campanha na temporada</small><strong>${v}V ${e}E ${d}D</strong><div class="forma" aria-label="Últimos jogos">${forma}</div></div>
        <div class="card stat"><small>Gols marcados / sofridos</small><strong>${gp} / ${gc}</strong><span>${plural(season.length, 'jogo', 'jogos')}</span></div>
        <div class="card stat"><small>Artilheiro da temporada</small><strong>${art ? art.g : 0}</strong><span>${art ? esc(art.nome) : '—'}${gar ? ` · Garçom: ${esc(gar.nome)} (${gar.a})` : ''}</span></div>
      </section>

      ${last ? `<section class="secao"><div class="secao-topo"><h2>Último episódio</h2><a href="#episodios">Todos os episódios →</a></div>${episodeCard(M, last)}</section>` : ''}

      <section class="secao grid grid-2">
        <div class="card grafico"><div class="secao-topo"><h3>Evolução na tabela</h3><a href="#tabela">Tabela completa →</a></div>${positionChart(M)}</div>
        <div class="card"><div class="secao-topo"><h3>Recordes na mira</h3><a href="#recordes">Todos →</a></div>
          <ul class="lista-simples">${destaques.map(({ r }) => {
            const c = car[r.chave], s = recordStatus(r.valor, c.valor);
            return `<li><div><b>${esc(r.titulo)}</b><small>${c.valor} de ${r.valor} · ${esc(s.texto)}</small><div class="barra" style="margin-top:6px"><span style="width:${s.pct}%"></span></div></div></li>`;
          }).join('') || '<li class="mudo">Cadastre os recordes na aba Recordes.</li>'}</ul>
        </div>
      </section>`;
  }

  function pageEpisodios(M) {
    const eps = [...M.episodes].reverse();
    return `<div class="secao-topo"><h2>Episódios</h2><span class="mudo">${plural(eps.filter(e => !isIntro(e)).length, 'episódio', 'episódios')}</span></div>
      ${eps.map(e => episodeCard(M, e)).join('') || '<p class="estado">Nenhum episódio cadastrado.</p>'}`;
  }

  function pageTabela(M, arg) {
    if (!M.tableEps.length) return '<p class="estado">Nenhuma classificação cadastrada.</p>';
    let ep = num(arg);
    if (!M.tables.has(ep)) ep = M.tableEps[M.tableEps.length - 1];
    const prevEp = M.tableEps.filter(x => x < ep).pop();
    const prevPos = new Map((prevEp != null ? M.tables.get(prevEp) : []).map(r => [norm(r.time), r.pos]));
    const table = M.tables.get(ep);
    const rows = table.map((r, i) => {
      const p = prevPos.get(norm(r.time));
      const dlt = p != null ? p - r.pos : 0;
      const delta = dlt ? `<span class="delta ${dlt > 0 ? 'sobe' : 'desce'}">${dlt > 0 ? '▲' : '▼'}${Math.abs(dlt)}</span>` : '';
      // Linha "…" quando há posições puladas (ex.: só os 6 primeiros e o clube mais abaixo).
      const prevRow = table[i - 1];
      const gap = prevRow && r.pos - prevRow.pos > 1 ? '<tr class="lacuna"><td colspan="10">…</td></tr>' : '';
      return `${gap}<tr class="${zoneClass(M, r.pos)} ${norm(r.time) === norm(M.clube) ? 'eu' : ''}">
        <td class="pos">${r.pos}</td><td class="esq">${esc(r.time)}${delta}</td><td class="pts">${r.pts}</td>
        <td>${r.j}</td><td>${r.v}</td><td>${r.e}</td><td>${r.d}</td><td>${r.gp}</td><td>${r.gc}</td><td>${r.sg > 0 ? '+' : ''}${r.sg}</td></tr>`;
    }).join('');
    const ep0 = M.episodes.find(e => e.n === ep);
    return `<div class="secao-topo">
        <div><h2>${esc(M.liga || 'Classificação')}</h2><span class="mudo">Ao final do episódio ${ep}${ep0 && ep0.titulo ? ' · ' + esc(ep0.titulo) : ''}${table.length < M.totalTimes ? ` · ${plural(table.length, 'time', 'times')} de ${M.totalTimes}` : ''}</span></div>
        <label>Episódio <select id="sel-ep">${M.tableEps.map(x => `<option value="${x}" ${x === ep ? 'selected' : ''}>EP ${String(x).padStart(2, '0')}</option>`).join('')}</select></label>
      </div>
      <div class="tabela-wrap"><table>
        <thead><tr><th>#</th><th class="esq">Time</th><th>Pts</th><th>J</th><th>V</th><th>E</th><th>D</th><th>GP</th><th>GC</th><th>SG</th></tr></thead>
        <tbody>${rows}</tbody>
      </table></div>
      <div class="legenda">${[
        ['l-acesso', 'Acesso direto', r => r.pos <= M.zonas.acesso],
        ['l-playoff', 'Playoff', r => r.pos === M.zonas.playoff || r.pos === M.zonas.playoffReb],
        ['l-reb', 'Rebaixamento', r => r.pos >= M.zonas.reb]
      ].filter(([, , has]) => table.some(has)).map(([cls, label]) => `<span class="${cls}">${label}</span>`).join('')}</div>`;
  }

  function playerTags(p) {
    const tags = [];
    if (p.origem === 'base') tags.push(`<span class="tag base"${p.chegada != null ? ` title="Subiu no EP ${p.chegada}"` : ''}>Base</span>`);
    if (p.origem === 'reforco') tags.push(`<span class="tag"${p.chegada != null ? ` title="Chegou no EP ${p.chegada}"` : ''}>Reforço</span>`);
    if (p.situacao === 'emprestado') tags.push(`<span class="tag emprestado">${esc(p.situacaoLabel || 'Emprestado')}${p.saida != null ? ` · EP ${p.saida}` : ''}</span>`);
    if (p.situacao === 'saiu') tags.push(`<span class="tag">${esc(p.situacaoLabel)}${p.saida != null ? ` no EP ${p.saida}` : ''}</span>`);
    return tags.length ? ` ${tags.join(' ')}` : '';
  }

  // Foto da EA quando existe; sem foto (ou se a imagem falhar) ficam as iniciais.
  function avatar(p) {
    const ini = p.nome.split(/\s+/).map(w => w[0] || '').join('').slice(0, 2).toUpperCase();
    const img = p.foto ? `<img src="${esc(p.foto)}" alt="" loading="lazy" referrerpolicy="no-referrer" onerror="this.remove()">` : '';
    return `<span class="avatar" aria-hidden="true">${esc(ini)}${img}</span>`;
  }

  const MOSTRAR = { todos: ['Todos', () => true], atual: ['Elenco atual', p => !p.situacao], emprestados: ['Emprestados', p => p.situacao === 'emprestado'], sairam: ['Saíram', p => p.situacao === 'saiu'] };
  const ui = { sort: 'g', dir: -1, temporada: null, mostrar: 'todos' };
  function pageJogadores(M) {
    if (ui.temporada == null) ui.temporada = M.temporadaAtual || 'todas';
    const games = ui.temporada === 'todas' ? M.games : M.games.filter(g => g.temporada === ui.temporada);
    const st = playerStats(M, games);
    M.elenco.forEach((e, k) => { if (!st.has(k)) st.set(k, { ...e, g: 0, a: 0 }); });
    const list = [...st.values()].map(p => ({ ...p, ga: p.g + p.a }));
    const key = ui.sort, dir = ui.dir;
    list.sort((a, b) => key === 'nome' ? dir * a.nome.localeCompare(b.nome, 'pt-BR') : dir * (a[key] - b[key]) || b.g - a.g || b.a - a.a || a.nome.localeCompare(b.nome, 'pt-BR'));
    const top = k => { const p = [...list].sort((a, b) => b[k] - a[k])[0]; return p && p[k] > 0 ? p : null; };
    const [tg, ta, tga] = [top('g'), top('a'), top('ga')];
    const base = list.filter(p => p.origem === 'base');
    const th = (k, label, cls = '') => `<th data-sort="${k}" class="${cls} ${ui.sort === k ? 'ordenado' : ''}">${label}${ui.sort === k ? (ui.dir < 0 ? ' ▾' : ' ▴') : ''}</th>`;
    return `<div class="secao-topo"><h2>Jogadores</h2>
        <label>Temporada <select id="sel-temp"><option value="todas" ${ui.temporada === 'todas' ? 'selected' : ''}>Todas</option>${M.seasons.map(s => `<option ${s === ui.temporada ? 'selected' : ''}>${esc(s)}</option>`).join('')}</select></label>
        <label>Mostrar <select id="sel-mostrar">${Object.entries(MOSTRAR).map(([k, [label]]) => `<option value="${k}" ${k === ui.mostrar ? 'selected' : ''}>${label}</option>`).join('')}</select></label>
      </div>
      <div class="grid grid-4" style="margin-bottom:16px">
        <div class="card stat"><small>Artilheiro</small><strong>${tg ? tg.g : 0}</strong><span>${tg ? esc(tg.nome) : '—'}</span></div>
        <div class="card stat"><small>Garçom</small><strong>${ta ? ta.a : 0}</strong><span>${ta ? esc(ta.nome) : '—'}</span></div>
        <div class="card stat"><small>Participações em gols</small><strong>${tga ? tga.ga : 0}</strong><span>${tga ? esc(tga.nome) : '—'}</span></div>
        ${base.length ? `<div class="card stat"><small>Crias da base</small><strong>${plural(base.length, 'jogador', 'jogadores')}</strong><span>${plural(base.reduce((s, p) => s + p.g, 0), 'gol', 'gols')} · ${plural(base.reduce((s, p) => s + p.a, 0), 'assistência', 'assistências')}</span></div>` : ''}
      </div>
      <div class="tabela-wrap"><table>
        <thead><tr>${th('nome', 'Jogador', 'esq')}<th>Pos</th><th class="esq">País</th>${th('g', 'Gols')}${th('a', 'Assist.')}${th('ga', 'G+A')}</tr></thead>
        <tbody>${list.filter(MOSTRAR[ui.mostrar][1]).map(p => `<tr class="${p.situacao ? 'fora' : ''}"><td class="esq"><span class="jogador">${avatar(p)}<span>${esc(p.nome)}${playerTags(p)}</span></span></td><td class="mudo">${esc(p.pos)}</td><td class="esq mudo">${esc(p.nac)}</td><td class="pts">${p.g}</td><td>${p.a}</td><td>${p.ga}</td></tr>`).join('')}</tbody>
      </table></div>`;
  }

  function pageRecordes(M) {
    const car = careerRecords(M);
    const reais = M.recordes.find(r => r.chave === 'titulos');
    const statusCls = s => { const n = norm(s); return n.startsWith('conclu') ? 'concluida' : n.startsWith('falh') ? 'falhou' : 'andamento'; };
    const trofeuIc = '<span class="trofeu-ic" aria-hidden="true"><svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor"><path d="M7 3h10v2h3v3a4 4 0 0 1-4 4h-.3A5 5 0 0 1 13 14.9V18h3v3H8v-3h3v-3.1A5 5 0 0 1 8.3 12H8a4 4 0 0 1-4-4V5h3V3zm0 4H6v1a2 2 0 0 0 1 1.7V7zm10 0v2.7A2 2 0 0 0 18 8V7h-1z"/></svg></span>';
    return `<div class="secao-topo"><h2>Recordes históricos</h2><span class="mudo">Recordes reais do clube até o início da série vs. a carreira</span></div>
      <div class="grid grid-3">${M.recordes.map(r => recordCard(M, r, car)).join('')}</div>

      <section class="secao grid grid-2">
        <div class="card"><h3>Metas da carreira</h3>
          <ul class="lista-simples" style="margin-top:8px">${M.metas.map(m => `<li><div><b>${esc(m.meta)}</b>${m.obs ? `<small>${esc(m.obs)}</small>` : ''}</div><span class="status ${statusCls(m.status)}">${esc(m.status)}</span></li>`).join('') || '<li class="mudo">Cadastre metas na aba Metas.</li>'}</ul>
        </div>
        <div class="card"><h3>Sala de troféus</h3>
          <p class="mudo" style="margin:6px 0 4px">Na carreira</p>
          <ul class="lista-simples">${M.titulos.map(t => `<li><div class="trofeu">${trofeuIc}<div><b>${esc(t.comp)}</b><small>${esc(t.temporada)}${t.obs ? ' · ' + esc(t.obs) : ''}</small></div></div></li>`).join('') || '<li class="mudo">Nenhum título ainda. A história está sendo escrita.</li>'}</ul>
          ${reais && reais.obs ? `<p class="mudo" style="margin:14px 0 4px">História real do clube</p><ul class="lista-simples">${reais.obs.split(/·|;|\n/).map(s => s.trim()).filter(Boolean).map(s => `<li><div class="trofeu">${trofeuIc}<b>${esc(s)}</b></div></li>`).join('')}</ul>` : ''}
        </div>
      </section>`;
  }

  // ---------- roteador ----------
  let MODEL = null;
  const pages = { inicio: pageInicio, episodios: pageEpisodios, tabela: pageTabela, jogadores: pageJogadores, recordes: pageRecordes };
  let lastPage = null;

  function render() {
    const [p, arg] = location.hash.replace(/^#/, '').split('/');
    const page = pages[p] ? p : 'inicio';
    document.querySelectorAll('#nav a').forEach(a => a.classList.toggle('ativo', a.dataset.page === page));
    $app.innerHTML = pages[page](MODEL, arg);
    if (page !== lastPage) window.scrollTo(0, 0);
    lastPage = page;
  }

  $app.addEventListener('change', ev => {
    if (ev.target.id === 'sel-ep') location.hash = `#tabela/${ev.target.value}`;
    if (ev.target.id === 'sel-temp') { ui.temporada = ev.target.value; render(); }
    if (ev.target.id === 'sel-mostrar') { ui.mostrar = ev.target.value; render(); }
  });
  $app.addEventListener('click', ev => {
    const th = ev.target.closest('th[data-sort]');
    if (!th) return;
    const k = th.dataset.sort;
    if (ui.sort === k) ui.dir *= -1; else { ui.sort = k; ui.dir = k === 'nome' ? 1 : -1; }
    render();
  });

  function applyBranding(M) {
    const hex = v => /^#[0-9a-f]{6}$/i.test(String(v || '').trim()) ? String(v).trim() : null;
    const root = document.documentElement.style;
    if (hex(M.cfg.corprimaria)) root.setProperty('--primary', hex(M.cfg.corprimaria));
    if (hex(M.cfg.corsecundaria)) root.setProperty('--secondary', hex(M.cfg.corsecundaria));
    const nome = String(M.cfg.nomeportal || 'Histórias em Campo');
    document.title = nome;
    document.getElementById('nome-portal').textContent = nome;
    document.getElementById('rodape-portal').textContent = nome;
    document.getElementById('sub-portal').textContent = `${M.clube} · Modo Carreira FC27`;
    if (M.cfg.canal) document.getElementById('rodape-canal').textContent = M.cfg.canal;
    if (M.cfg.linkcanal) document.getElementById('rodape-canal').href = M.cfg.linkcanal;
    const escudo = document.getElementById('escudo');
    if (M.cfg.escudo) escudo.innerHTML = `<img src="${esc(M.cfg.escudo)}" alt="Escudo do ${esc(M.clube)}">`;
    else escudo.textContent = (M.clube.split(/\s+/).pop() || 'C')[0].toUpperCase();

    const avisos = [];
    if (sim(M.cfg.dadosexemplo)) avisos.push('<div class="aviso"><b>Dados de exemplo.</b> Jogadores, placares e tabelas são fictícios, só para mostrar o formato do portal.</div>');
    if (M.warnings.length) avisos.push(`<details class="aviso"><summary>Conferência da planilha: ${plural(M.warnings.length, 'ponto', 'pontos')} para revisar</summary><ul>${M.warnings.map(w => `<li>${esc(w)}</li>`).join('')}</ul></details>`);
    $avisos.innerHTML = avisos.join('');
  }

  async function init() {
    try {
      if (typeof XLSX === 'undefined') throw new Error('A biblioteca de leitura de planilhas não carregou. Verifique a conexão.');
      const bust = `?v=${Date.now()}`;
      // Os vídeos são opcionais: sem o arquivo, o site usa só a planilha.
      const optional = url => fetch(url + bust, { cache: 'no-store' }).then(r => r.ok ? r.json() : null).catch(() => null);
      const videosReq = optional(VIDEOS_URL), fotosReq = optional(FOTOS_URL);
      const res = await fetch(DATA_URL + bust, { cache: 'no-store' });
      if (!res.ok) throw new Error(`Não foi possível baixar ${DATA_URL} (HTTP ${res.status}).`);
      const wb = XLSX.read(await res.arrayBuffer(), { type: 'array' });
      MODEL = build(wb, await videosReq, await fotosReq);
      applyBranding(MODEL);
      window.addEventListener('hashchange', render);
      render();
    } catch (err) {
      console.error(err);
      $app.innerHTML = `<div class="estado erro"><b>Não foi possível carregar os dados.</b><br>${esc(err.message)}</div>`;
    }
  }

  init();
})();
