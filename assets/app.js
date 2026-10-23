/* Histórias em Campo — lê dados/dados.xlsx e monta o portal no navegador. */
(() => {
  'use strict';

  const DATA_URL = 'dados/dados.xlsx';
  const VIDEOS_URL = 'dados/videos.json'; // gerado pela automação do GitHub a partir do feed do canal
  const FOTOS_URL = 'dados/fotos.json'; // miniaturas da EA por nome; a coluna Foto da planilha tem prioridade
  const CLUBES_URL = 'dados/clubes.json'; // escudos, estádios e logos de campeonatos (Wikidata); a aba Clubes tem prioridade
  const FOTOS_CORPO_URL = 'dados/fotos-corpo.json'; // foto de corpo inteiro, usada no card do artilheiro
  const TREINADOR_URL = 'dados/treinador.json';
  const TEMPORADAS_URL = 'dados/temporadas.json'; // nome e logo de cada temporada da série
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
  const plural = (n, um, varios) => `${n} ${n === 1 ? um : varios}`;
  const ordinal = n => `${n}º`;
  const pad2 = n => String(n).padStart(2, '0');
  const RES_NOME = { V: 'Vitória', E: 'Empate', D: 'Derrota' };

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
  // Valores em euro: aceita número da planilha ou texto como "12 mi", "500 mil", "€ 1.200.000".
  function parseMoney(v) {
    if (typeof v === 'number') return v;
    let s = String(v ?? '').toLowerCase().replace(/€|eur|\s/g, '');
    if (!s) return null;
    const mult = /mi|m$/.test(s) ? 1e6 : /mil|k$/.test(s) ? 1e3 : 1;
    s = s.replace(/[a-z]/g, '');
    if (/^\d{1,3}(\.\d{3})+(,\d+)?$/.test(s)) s = s.replace(/\./g, '');
    const n = Number(s.replace(',', '.'));
    return Number.isFinite(n) ? n * mult : null;
  }
  function fmtMoney(n) {
    if (n == null) return '—';
    if (n === 0) return 'Sem custo';
    const a = Math.abs(n), f = (x, d) => x.toLocaleString('pt-BR', { maximumFractionDigits: d });
    return a >= 1e6 ? `€ ${f(n / 1e6, 1)} mi` : a >= 1e3 ? `€ ${f(n / 1e3, 0)} mil` : `€ ${f(n, 0)}`;
  }
  function ytId(url) {
    const m = String(url || '').match(/(?:v=|youtu\.be\/|shorts\/|embed\/|live\/)([\w-]{11})/);
    return m ? m[1] : null;
  }
  const initials = (name, n = 2) => String(name || '').split(/\s+/).map(w => w[0] || '').join('').slice(0, n).toUpperCase();
  // Sigla do clube sem prefixos genéricos ("1. FC Nürnberg" -> "NÜR").
  const clubInitials = name => {
    const words = String(name || '').split(/\s+/).filter(w => !/^(\d+\.?|fc|vfl|sv|sc|spvgg|bsc|tsg|fsv|ssv|vfb|afc|cf|cd|ac|as|ec|se|club|clube)$/i.test(w));
    return (words[0] || name || '?').slice(0, 3).toUpperCase();
  };

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
  function build(wb, extras) {
    const { videos, fotos, clubesJson, treinador, fotosCorpo, temporadas } = extras;
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

    // Escudos e estádios: aba Clubes (se existir) tem prioridade sobre o que a automação encontrou.
    const clubes = new Map();
    Object.entries((clubesJson && clubesJson.clubes) || {}).forEach(([k, v]) => clubes.set(norm(k), { escudo: v.escudo || '', estadio: v.estadio || '' }));
    // Link de página (ex.: artigo da Wikipédia) não é imagem: nesse caso vale o escudo da automação.
    const ehImagem = u => /\.(png|jpe?g|svg|webp|gif)(\?|$)|special:filepath|upload\.wikimedia|gstatic|googleusercontent|ytimg/i.test(u) && !/wikipedia\.org\/wiki\//i.test(u);
    sheet(wb, 'Clubes').forEach(r => {
      if (!r.clube) return;
      const atual = clubes.get(norm(r.clube)) || {};
      const escudo = String(r.escudo || '').trim();
      if (escudo && !ehImagem(escudo)) warnings.push(`Clubes: o escudo do ${r.clube} é um link de página, não de imagem. Use o link direto da imagem ou deixe vazio.`);
      clubes.set(norm(r.clube), { escudo: (ehImagem(escudo) ? escudo : '') || atual.escudo || '', estadio: String(r.estadio || '') || atual.estadio || '' });
    });
    const competicoes = new Map(Object.entries((clubesJson && clubesJson.competicoes) || {}).map(([k, v]) => [norm(k), v.logo || '']));

    const fotosMap = new Map(Object.entries(fotos || {}).map(([k, v]) => [norm(k), String(v || '')]));
    const fotosCorpoMap = new Map(Object.entries(fotosCorpo || {}).map(([k, v]) => [norm(k), String(v || '')]));
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
    // Origem: "Base" (cria da base), "Contratação" (reforço) ou vazio/"Elenco inicial".
    const origemTipo = v => { const n = norm(v); return n.startsWith('base') ? 'base' : /^(contrat|reforc|emprest)/.test(n) ? 'reforco' : ''; };

    // Elenco = jogadores do elenco atual (até o último episódio).
    const elenco = new Map();
    sheet(wb, 'Elenco').forEach(r => {
      if (r.jogador) elenco.set(norm(r.jogador), {
        nome: String(r.jogador), pos: String(r.posicao || ''), nac: String(r.nacionalidade || ''),
        origem: origemTipo(r.origem), chegada: epNum(r.chegada, r.jogador, 'Chegada'), saida: epNum(r.saida, r.jogador, 'Saida'),
        ...situacao(r), situacaoManual: !!situacao(r).situacao, doElenco: true, foto: String(r.foto || '').trim() || fotosMap.get(norm(r.jogador)) || ''
      });
    });

    // Transferências: chegadas e saídas por temporada.
    // Posição, país e foto vêm da própria aba se existirem; senão, do Elenco.
    const transfers = sheet(wb, 'Transferencias')
      .map(r => {
        const ep = num(r.episodio), e = elenco.get(norm(r.jogador)) || {};
        return {
          temporada: season(r.temporada) || epSeason.get(ep) || season(cfg.temporadaatual), ep,
          tipo: norm(r.tipo).startsWith('sa') ? 'saida' : 'chegada',
          jogador: String(r.jogador || ''), pos: String(r.posicao || e.pos || ''), idade: num(r.idade), nac: String(r.nacionalidade || e.nac || ''),
          clube: String(r.clube || ''), valor: parseMoney(r.valor), valorTexto: String(r.valor ?? ''), obs: String(r.observacao || ''),
          foto: String(r.foto || '').trim() || e.foto || fotosMap.get(norm(r.jogador)) || ''
        };
      })
      .filter(t => t.jogador)
      .sort((a, b) => (a.ep ?? 0) - (b.ep ?? 0));

    // Cada transferência, em ordem de episódio, atualiza a situação do jogador (reforço, emprestado, vendido).
    // A coluna Situacao do Elenco, quando preenchida, tem a palavra final.
    transfers.forEach(t => {
      const k = norm(t.jogador);
      const p = elenco.get(k) || { nome: t.jogador, pos: t.pos, nac: t.nac, origem: '', situacao: '', situacaoLabel: '', foto: t.foto };
      if (!p.pos) p.pos = t.pos;
      if (!p.nac) p.nac = t.nac;
      if (!p.foto) p.foto = t.foto;
      if (t.tipo === 'chegada') {
        if (!p.origem) p.origem = 'reforco';
        if (p.chegada == null) p.chegada = t.ep;
        if (!p.situacaoManual) { p.situacao = ''; p.situacaoLabel = ''; p.saida = null; } // voltou ao clube
      } else if (!p.situacaoManual) {
        const emprestimo = norm(t.obs).startsWith('empre');
        p.situacao = emprestimo ? 'emprestado' : 'saiu';
        p.situacaoLabel = emprestimo ? 'Emprestado' : t.valor ? 'Vendido' : 'Saiu';
        p.saida = t.ep;
      }
      elenco.set(k, p);
    });

    const recordes = sheet(wb, 'Recordes').map(r => ({
      chave: norm(r.chave), titulo: String(r.recorde || r.chave), valor: num(r.valor),
      detentor: String(r.detentor || ''), quando: String(r.quando || ''), obs: String(r.observacao || ''), confirmado: sim(r.confirmado)
    }));
    // Metas: Imagem (arquivo em assets/metas/ ou link), Atual/Alvo para a barra de progresso e Temporada opcionais.
    // Tipo "Regra" = regra da carreira (sem status). Chave liga contagens automáticas (ex.: sulamericanos).
    const metas = sheet(wb, 'Metas').map(r => ({
      meta: String(r.meta || ''), status: String(r.status || 'Em andamento'), obs: String(r.observacao || ''),
      imagem: String(r.imagem || ''), atual: num(r.atual), alvo: num(r.alvo), temporada: season(r.temporada),
      regra: norm(r.tipo).startsWith('regra'), chave: norm(r.chave)
    })).filter(m => m.meta);
    const titulos = sheet(wb, 'Titulos').map(r => ({ temporada: season(r.temporada), comp: String(r.competicao || ''), obs: String(r.observacao || '') })).filter(t => t.comp);

    const seasons = [...new Set([...episodes.map(e => e.temporada), ...games.map(g => g.temporada), ...transfers.map(t => t.temporada)].filter(Boolean))];
    const temporadaAtual = season(cfg.temporadaatual) || seasons[seasons.length - 1] || '';
    const temporadaInfo = Object.entries(temporadas || {}).find(([k]) => season(k) === temporadaAtual)?.[1] || {};

    return {
      cfg, clube, liga, episodes, games, tables, tableEps, elenco, fotosMap, fotosCorpoMap, clubes, competicoes, transfers, treinador,
      recordes, metas, titulos, seasons, temporadaAtual, temporadaInfo, warnings,
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

  // ---------- componentes ----------
  const clubInfo = (M, name) => M.clubes.get(norm(name)) || {};
  // Escudo com sigla por baixo: se a imagem não existir ou falhar, fica a sigla.
  function crest(M, name, cls = '') {
    const c = clubInfo(M, name);
    const img = c.escudo ? `<img src="${esc(c.escudo)}" alt="" loading="lazy" referrerpolicy="no-referrer" onerror="this.remove()">` : '';
    return `<span class="escudo-clube ${cls}" title="${esc(name)}" aria-hidden="true"><i>${esc(clubInitials(name))}</i>${img}</span>`;
  }
  function compLogo(M, comp) {
    const url = M.competicoes.get(norm(comp));
    return url ? `<img class="logo-comp" src="${esc(url)}" alt="" loading="lazy" referrerpolicy="no-referrer" onerror="this.remove()">` : '';
  }
  // Foto da EA quando existe; sem foto (ou se a imagem falhar) ficam as iniciais.
  function avatar(p, cls = '') {
    const img = p.foto ? `<img src="${esc(p.foto)}" alt="" loading="lazy" referrerpolicy="no-referrer" onerror="this.remove()">` : '';
    return `<span class="avatar ${cls}" aria-hidden="true">${esc(initials(p.nome || p.jogador))}${img}</span>`;
  }
  const rodadaLabel = g => g.rodada === '' ? '' : (num(g.rodada) != null ? `Rodada ${g.rodada}` : esc(g.rodada));
  function golsLista(g) {
    return g.gols.length ? `<ul class="gols-lista">${g.gols.map(x => `<li><span class="ic-gol"></span>${x.minuto ? `<span class="min">${x.minuto.texto}</span>` : ''}${esc(x.autor)}${x.assist ? ` <small>assist. ${esc(x.assist)}</small>` : ''}</li>`).join('')}</ul>` : '';
  }

  // Jogo compacto (dentro do card do episódio).
  function gameCard(M, g) {
    const home = g.casa ? M.clube : g.adv, away = g.casa ? g.adv : M.clube;
    const hs = g.casa ? g.gp : g.gc, as = g.casa ? g.gc : g.gp;
    const rodada = rodadaLabel(g);
    return `<div class="jogo">
      <div class="jogo-meta">${esc(g.comp)}${rodada ? ' · ' + rodada : ''}</div>
      <div class="placar">
        <span class="time ${g.casa ? 'eu' : ''}">${esc(home)} ${crest(M, home, 'mini')}</span>
        <span class="gols">${hs}<i> x </i>${as}</span>
        <span class="time ${g.casa ? '' : 'eu'}">${crest(M, away, 'mini')} ${esc(away)}</span>
        <span class="res res-${g.res}" title="${RES_NOME[g.res]}">${g.res}</span>
      </div>${golsLista(g)}
    </div>`;
  }

  // Tema visual do card pelo nome do campeonato: liga, copa, europa ou padrão.
  function compTema(comp) {
    const n = norm(comp);
    if (/champions|europa|conference|uefa/.test(n)) return 'europa';
    if (/pokal|copa|cup|taca|supercup|supercopa/.test(n)) return 'copa';
    if (/liga|bundesliga|league|serie|division/.test(n)) return 'liga';
    return 'padrao';
  }

  // Jogo completo (página Jogos): escudos, estádio, campeonato e resultado.
  function matchCard(M, g) {
    const home = g.casa ? M.clube : g.adv, away = g.casa ? g.adv : M.clube;
    const hs = g.casa ? g.gp : g.gc, as = g.casa ? g.gc : g.gp;
    const estadio = clubInfo(M, home).estadio;
    const rodada = rodadaLabel(g);
    return `<article class="card partida borda-${g.res} tema-${compTema(g.comp)}">
      <div class="partida-meta">${compLogo(M, g.comp)}<span>${esc(g.comp)}${rodada ? ' · ' + rodada : ''}</span><span class="res-texto res-${g.res}-texto">${RES_NOME[g.res]}</span></div>
      <div class="partida-placar">
        <div class="lado ${g.casa ? 'eu' : ''}">${crest(M, home, 'grande')}<span>${esc(home)}</span></div>
        <div class="centro"><strong>${hs}<i>x</i>${as}</strong></div>
        <div class="lado ${g.casa ? '' : 'eu'}">${crest(M, away, 'grande')}<span>${esc(away)}</span></div>
      </div>
      <div class="partida-local">${g.casa ? 'Em casa' : 'Fora de casa'}${estadio ? ` · ${esc(estadio)}` : ''}</div>
      ${golsLista(g)}
    </article>`;
  }

  function episodeCard(M, e) {
    const id = ytId(e.link);
    const intro = isIntro(e);
    const pad = pad2(e.n);
    const thumb = id
      ? `<a class="ep-thumb" href="${esc(e.link)}" target="_blank" rel="noopener"><img src="https://i.ytimg.com/vi/${id}/maxresdefault.jpg" alt="Miniatura ${intro ? 'da intro' : `do episódio ${e.n}`}" loading="lazy" onload="if(this.naturalWidth<200)this.src='https://i.ytimg.com/vi/${id}/mqdefault.jpg'" onerror="this.onerror=null;this.src='https://i.ytimg.com/vi/${id}/mqdefault.jpg'"><span class="play">▶ Assistir</span></a>`
      : `<div class="ep-thumb"><span class="ep-num">${intro ? 'INTRO' : `EP ${pad}`}</span>${e.link ? '' : '<span class="play">Vídeo em breve</span>'}</div>`;
    const games = M.games.filter(g => g.ep === e.n);
    const me = clubRow(M, e.n);
    const prevEp = M.tableEps.filter(x => x < e.n).pop();
    const prev = prevEp != null ? clubRow(M, prevEp) : null;
    const delta = me && prev ? prev.pos - me.pos : 0;
    const pos = me ? `<span class="ep-pos">Fim do episódio: <b>${ordinal(me.pos)} lugar</b> · ${me.pts} pts${delta ? ` <span class="delta ${delta > 0 ? 'sobe' : 'desce'}">${delta > 0 ? '▲' : '▼'}${Math.abs(delta)}</span>` : ''}</span>` : '';
    // Resultados escondidos por padrão: quem ainda não assistiu não leva spoiler.
    const resultados = games.length || pos
      ? `<details class="spoiler"><summary><span>Mostrar resultados do episódio</span><small>${plural(games.length, 'jogo', 'jogos')} · contém spoilers</small></summary>
          <div class="jogos">${games.map(g => gameCard(M, g)).join('')}</div>${pos}
        </details>`
      : (intro ? '' : '<p class="mudo">Sem jogos registrados.</p>');
    return `<article class="card ep" id="ep-${e.n}">
      ${thumb}
      <div>
        <div class="ep-cabecalho"><small>${intro ? 'INTRO' : `EPISÓDIO ${pad}`}${e.temporada ? ' · ' + esc(e.temporada) : ''}</small><h3>${esc(e.titulo || (intro ? 'Intro da série' : `Episódio ${pad}`))}</h3></div>
        ${e.data ? `<div class="ep-data">Lançado em ${esc(e.data)}</div>` : ''}
        ${e.resumo ? `<p class="ep-resumo">${esc(e.resumo)}</p>` : ''}
        ${resultados}
      </div>
    </article>`;
  }

  // Tabela no estilo do jogo: faixas de Acesso / Mata-mata / Rebaixamento.
  function tableCard(M) {
    const ep = M.tableEps[M.tableEps.length - 1];
    if (ep == null) return '<div class="card"><h3>Classificação</h3><p class="mudo">A tabela aparece quando houver classificação na planilha.</p></div>';
    const z = M.zonas, table = M.tables.get(ep);
    const zona = pos => pos <= z.acesso ? 'Acesso' : pos === z.playoff ? 'Mata-mata: acesso' : pos >= z.reb ? 'Rebaixamento' : pos === z.playoffReb ? 'Mata-mata: rebaixamento' : '';
    let ultimaZona = null;
    const rows = table.map((r, i) => {
      const zn = zona(r.pos), prev = table[i - 1];
      let sep = '';
      if (zn !== ultimaZona) { sep = `<tr class="fc-zona"><td colspan="10">${esc(zn)}</td></tr>`; ultimaZona = zn; }
      else if (prev && r.pos - prev.pos > 1) sep = '<tr class="fc-zona"><td colspan="10">…</td></tr>';
      return `${sep}<tr class="${norm(r.time) === norm(M.clube) ? 'eu' : ''}">
        <td class="pos">${r.pos}</td><td class="esq"><span class="fc-clube">${crest(M, r.time, 'mini')}${esc(r.time)}</span></td>
        <td>${r.j}</td><td>${r.v}</td><td>${r.e}</td><td>${r.d}</td><td>${r.gp}</td><td>${r.gc}</td><td>${r.sg > 0 ? '+' : ''}${r.sg}</td><td class="pts">${r.pts}</td></tr>`;
    }).join('');
    return `<div class="tabela-fc">
      <div class="fc-topo">${compLogo(M, M.liga)}<b>${esc(M.liga || 'Classificação')}</b><span>${esc(M.temporadaAtual)} · fim do EP ${pad2(ep)}</span></div>
      <div class="fc-scroll"><table>
        <thead><tr><th>Pos</th><th class="esq">Clube</th><th>J</th><th>V</th><th>E</th><th>D</th><th>GP</th><th>GS</th><th>SG</th><th>Pts</th></tr></thead>
        <tbody>${rows}</tbody>
      </table></div>
    </div>`;
  }

  // Artilheiro da temporada: usa a foto de corpo inteiro (fotos-corpo.json) quando existir.
  function scorerCard(M, art, jogos) {
    if (!art) return '<div class="card"><h3>Artilheiro da temporada</h3><p class="mudo">Aparece quando houver gols registrados.</p></div>';
    const corpo = M.fotosCorpoMap.get(norm(art.nome));
    const media = jogos ? (art.g / jogos).toLocaleString('pt-BR', { maximumFractionDigits: 2 }) : '0';
    return `<div class="card artilheiro">
      <span class="artilheiro-marca" aria-hidden="true">${art.g}</span>
      <div class="artilheiro-texto">
        <small class="sobretitulo">Artilheiro da temporada</small>
        <h2>${esc(art.nome)}</h2>
        <span class="artilheiro-info">${[art.pos, art.nac].filter(Boolean).map(esc).join(' · ')}</span>
        <div class="artilheiro-numeros">
          <div><strong>${art.g}</strong><small>gols</small></div>
          <div><strong>${art.a}</strong><small>assist.</small></div>
          <div><strong>${media}</strong><small>gols por jogo</small></div>
        </div>
        <a class="botao secundario" href="#jogadores">Todos os jogadores →</a>
      </div>
      <div class="artilheiro-foto">${corpo ? `<img src="${esc(corpo)}" alt="${esc(art.nome)}" onerror="this.remove()">` : avatar(art, 'gigante')}</div>
    </div>`;
  }

  // ---------- agenda de lançamentos ----------
  const DIAS = ['domingo', 'segunda', 'terça', 'quarta', 'quinta', 'sexta', 'sábado'];
  const NO_AR_MS = 60 * 60 * 1000; // por 1 hora depois do horário mostra "no ar"

  // Próximo lançamento em ms (UTC), a partir de dias da semana + hora no fuso do canal.
  function proximoLancamento(cfg, agora = Date.now()) {
    const dias = (cfg.dias || []).map(d => DIAS.findIndex(x => norm(x) === norm(d))).filter(i => i >= 0);
    if (!dias.length) return null;
    const [h, min] = String(cfg.hora || '18:00').split(':').map(Number);
    const off = (cfg.utc ?? -3) * 3600000;
    const local = new Date(agora + off); // "relógio" do canal lido via getUTC*
    for (let add = 0; add <= 7; add++) {
      const d = new Date(Date.UTC(local.getUTCFullYear(), local.getUTCMonth(), local.getUTCDate() + add, h, min || 0));
      const alvo = d.getTime() - off;
      if (dias.includes(d.getUTCDay()) && alvo + NO_AR_MS > agora) return alvo;
    }
    return null;
  }

  // Faixa fina abaixo do destaque: próximo episódio, dia/hora e contagem regressiva.
  function agendaBar(M) {
    const cfg = M.temporadaInfo.lancamentos;
    if (!cfg) return '';
    const nomes = (cfg.dias || []).join(' e ');
    const hora = String(cfg.hora || '18:00').replace(/:00$/, 'h').replace(':', 'h');
    const publicados = M.episodes.filter(e => !isIntro(e) && ytId(e.link));
    const prox = (publicados.length ? publicados[publicados.length - 1].n : 0) + 1;
    const alvo = proximoLancamento(cfg);
    const off = (cfg.utc ?? -3) * 3600000;
    const dia = alvo != null ? DIAS[new Date(alvo + off).getUTCDay()] : '';
    return `<div class="faixa-agenda">
      <span class="ponto-vivo"></span>
      <span class="agenda-proximo">Próximo episódio · <b>EP ${pad2(prox)}</b>${dia ? ` · ${esc(dia)}, ${esc(hora)}` : ''}</span>
      <span class="contagem" data-contagem aria-live="off"><span class="contagem-rotulo">faltam</span>
        <span><b data-u="d">–</b><small>d</small></span><span><b data-u="h">–</b><small>h</small></span><span><b data-u="m">–</b><small>min</small></span><span><b data-u="s">–</b><small>s</small></span>
      </span>
      <span class="agenda-regra">Novos episódios toda ${esc(nomes)}, às ${esc(hora)}</span>
    </div>`;
  }

  let contagemTimer = null;
  function startCountdown() {
    clearInterval(contagemTimer);
    const box = $app.querySelector('[data-contagem]');
    const cfg = MODEL && MODEL.temporadaInfo.lancamentos;
    if (!box || !cfg) return;
    const tick = () => {
      if (!box.isConnected) return clearInterval(contagemTimer);
      const alvo = proximoLancamento(cfg), agora = Date.now();
      if (alvo == null) return;
      if (alvo <= agora) { box.classList.add('no-ar'); box.innerHTML = '<span class="contagem-rotulo">no ar agora!</span>'; return; }
      let s = Math.floor((alvo - agora) / 1000);
      const v = { d: Math.floor(s / 86400), h: Math.floor(s % 86400 / 3600), m: Math.floor(s % 3600 / 60), s: s % 60 };
      box.querySelectorAll('[data-u]').forEach(el => { el.textContent = el.dataset.u === 'd' ? v.d : pad2(v[el.dataset.u]); });
    };
    tick();
    contagemTimer = setInterval(tick, 1000);
  }

  // Fotos do treinador alternando a cada 10 segundos (o timer é refeito a cada render).
  let slideTimer = null;
  function startSlides() {
    clearInterval(slideTimer);
    const slides = [...$app.querySelectorAll('.treinador-foto img')];
    if (slides.length < 2 || matchMedia('(prefers-reduced-motion: reduce)').matches) return;
    let i = 0;
    slideTimer = setInterval(() => {
      const vivos = slides.filter(s => s.isConnected);
      if (vivos.length < 2) return clearInterval(slideTimer);
      vivos[i % vivos.length].classList.remove('ativo');
      i++;
      vivos[i % vivos.length].classList.add('ativo');
    }, 10000);
  }

  function coachSection(M) {
    const t = M.treinador;
    if (!t || !t.nome) return '';
    const imagens = Array.isArray(t.imagens) ? t.imagens : t.imagem ? [{ src: t.imagem }] : [];
    const imgs = imagens.map((im, i) => `<img class="${i === 0 ? 'ativo' : ''}" src="${esc(im.src)}" alt="${esc(t.nome)}" style="object-position:${esc(im.posicao || 'center')}" onerror="this.remove()">`).join('');
    const marcos = (t.marcos || []).map(m => {
      const titulo = m.titulo || m.texto, sub = m.titulo ? m.texto : '';
      return `<li class="lt-item lt-${esc(m.tipo || 'clube')}${m.destaque ? ' lt-destaque' : ''}">
        <span class="lt-ano">${esc(m.ano)}</span>
        <span class="lt-ponto" aria-hidden="true">${LT_ICONES[m.tipo] || LT_ICONES.clube}</span>
        <b>${esc(titulo)}</b>${sub ? `<small>${esc(sub)}</small>` : ''}
      </li>`;
    }).join('');
    return `<section class="secao treinador-bloco">
      <span class="treinador-marca" aria-hidden="true">10</span>
      <div class="treinador">
        <div class="treinador-foto"><span>${esc(initials(t.nome))}</span>${imgs}</div>
        <div class="treinador-texto">
          <small class="sobretitulo">${esc(t.cargo || 'O treinador')}</small>
          <h2>${esc(t.nome)}</h2>
          ${(t.paragrafos || []).map((p, i) => `<p class="${i === 0 ? 'treinador-lead' : ''}">${esc(p)}</p>`).join('')}
        </div>
      </div>
      ${marcos ? `<div class="lt-bloco"><small class="sobretitulo">Trajetória</small><ol class="lt">${marcos}</ol></div>` : ''}
    </section>`;
  }

  // Ícones da linha do tempo do treinador (24x24, cor herdada).
  const LT_ICONES = {
    clube: '<svg viewBox="0 0 24 24" fill="currentColor"><path d="M12 2l8 3v6c0 5-3.5 9-8 11-4.5-2-8-6-8-11V5z"/></svg>',
    selecao: '<svg viewBox="0 0 24 24" fill="currentColor"><path d="M12 2l3 6.5 7 .8-5.2 4.8 1.5 7-6.3-3.6-6.3 3.6 1.5-7L2 9.3l7-.8z"/></svg>',
    titulo: '<svg viewBox="0 0 24 24" fill="currentColor"><path d="M7 3h10v2h3v3a4 4 0 0 1-4 4h-.3A5 5 0 0 1 13 14.9V18h3v3H8v-3h3v-3.1A5 5 0 0 1 8.3 12H8a4 4 0 0 1-4-4V5h3V3zm0 4H6v1a2 2 0 0 0 1 1.7V7zm10 0v2.7A2 2 0 0 0 18 8V7h-1z"/></svg>',
    treinador: '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"><rect x="5" y="4" width="14" height="17" rx="2"/><path d="M9 4V2.5h6V4M8.5 10h7M8.5 14h4"/></svg>'
  };

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
    // Destaque é a temporada (ex.: "Entre o Tango e os Lobos"); "Histórias em Campo" é a série.
    const temp = M.temporadaInfo;
    const serie = String(M.cfg.nomeportal || 'Histórias em Campo');
    const titulo = temp.logo
      ? `<img src="${esc(temp.logo)}" alt="${esc(temp.nome || serie)}" onerror="this.parentNode.textContent=this.alt">`
      : esc(temp.nome || serie);
    return `
      <section class="hero hero-temporada">
        <div class="hero-conteudo">
          <small class="hero-sobre">${esc(serie)} · Modo Carreira Realista</small>
          <h1 class="hero-logo">${titulo}</h1>
          ${temp.chamada ? `<p class="hero-chamada">${esc(temp.chamada)}</p>` : ''}
          <div class="botoes">${link}<a class="botao secundario" href="#episodios">Ver episódios</a></div>
        </div>
        <div class="hero-arte" aria-hidden="true">${temp.imagem ? `<img src="${esc(temp.imagem)}" alt="" onerror="this.remove()">` : ''}</div>
      </section>
      ${agendaBar(M)}

      <section class="secao grid grid-4">
        <div class="card stat stat-tile"><small>Posição atual</small><strong>${me ? ordinal(me.pos) : '—'}</strong><span>${me ? `${me.pts} pts · ${esc(M.liga)}` : 'sem tabela'}</span></div>
        <div class="card stat stat-tile"><small>Campanha na temporada</small><strong>${v}V ${e}E ${d}D</strong><div class="forma" aria-label="Últimos jogos">${forma}</div></div>
        <div class="card stat stat-tile"><small>Gols marcados / sofridos</small><strong>${gp} / ${gc}</strong><span>${plural(season.length, 'jogo', 'jogos')}</span></div>
        <div class="card stat stat-tile"><small>Garçom da temporada</small><strong>${gar ? gar.a : 0}</strong><span>${gar ? `${esc(gar.nome)} · assistências` : '—'}</span></div>
      </section>

      <section class="secao grid grid-2">
        ${tableCard(M)}
        ${scorerCard(M, art, season.length)}
      </section>

      ${last ? `<section class="secao"><div class="secao-topo"><h2>Último episódio</h2><a href="#episodios">Todos os episódios →</a></div><div class="ep-destaque">${episodeCard(M, last)}</div></section>` : ''}

      ${coachSection(M)}

      <section class="secao card"><div class="secao-topo"><h3>Recordes na mira</h3><a href="#recordes">Todos →</a></div>
        <ul class="lista-simples">${destaques.map(({ r }) => {
          const c = car[r.chave], s = recordStatus(r.valor, c.valor);
          return `<li><div><b>${esc(r.titulo)}</b><small>${c.valor} de ${r.valor} · ${esc(s.texto)}</small><div class="barra" style="margin-top:6px"><span style="width:${s.pct}%"></span></div></div></li>`;
        }).join('') || '<li class="mudo">Cadastre os recordes na aba Recordes.</li>'}</ul>
      </section>`;
  }

  function pageEpisodios(M) {
    const eps = [...M.episodes].reverse();
    return `<div class="secao-topo"><h2>Episódios</h2><span class="mudo">${plural(eps.filter(e => !isIntro(e)).length, 'episódio', 'episódios')}</span></div>
      ${eps.map(e => episodeCard(M, e)).join('') || '<p class="estado">Nenhum episódio cadastrado.</p>'}`;
  }

  // Do último jogo para o primeiro, agrupados por episódio.
  function pageJogos(M) {
    if (!M.games.length) return '<p class="estado">Nenhum jogo cadastrado.</p>';
    const games = [...M.games].reverse();
    const v = M.games.filter(g => g.res === 'V').length, e = M.games.filter(g => g.res === 'E').length, d = M.games.filter(g => g.res === 'D').length;
    let html = '', epAtual;
    games.forEach(g => {
      if (g.ep !== epAtual) {
        epAtual = g.ep;
        const ep = M.episodes.find(x => x.n === g.ep);
        html += `<h3 class="grupo-ep"><span>EP ${pad2(g.ep)}</span>${ep && ep.titulo ? esc(ep.titulo) : ''}</h3>`;
      }
      html += matchCard(M, g);
    });
    return `<div class="secao-topo"><h2>Jogos</h2><span class="mudo">${plural(M.games.length, 'jogo', 'jogos')} · ${v}V ${e}E ${d}D</span></div>
      <p class="aviso-spoiler">Esta página mostra todos os resultados até o último episódio lançado.</p>
      <div class="lista-partidas">${html}</div>`;
  }

  function playerTags(p) {
    const tags = [];
    if (p.origem === 'base') tags.push(`<span class="tag base"${p.chegada != null ? ` title="Subiu no EP ${p.chegada}"` : ''}>Base</span>`);
    if (p.origem === 'reforco') tags.push(`<span class="tag"${p.chegada != null ? ` title="Chegou no EP ${p.chegada}"` : ''}>Reforço</span>`);
    if (p.situacao === 'emprestado') tags.push(`<span class="tag emprestado"${p.saida != null ? ` title="Emprestado no EP ${p.saida}"` : ''}>Emprestado</span>`);
    if (p.situacao === 'saiu') tags.push(`<span class="tag">${esc(p.situacaoLabel)}${p.saida != null ? ` no EP ${p.saida}` : ''}</span>`);
    return tags.length ? ` ${tags.join(' ')}` : '';
  }

  const POS_ORDEM = ['gol', 'ld', 'zag', 'le', 'ala', 'vol', 'mc', 'md', 'me', 'mei', 'pd', 'pe', 'sa', 'ata'];
  const MOSTRAR = { todos: ['Todos', () => true], atual: ['Elenco atual', p => !p.situacao], emprestados: ['Emprestados', p => p.situacao === 'emprestado'], sairam: ['Saíram', p => p.situacao === 'saiu'] };
  const ui = { sort: 'g', dir: -1, temporada: null, mostrar: 'todos', tempTransf: null };
  function seasonSelect(id, value, M) {
    return `<label>Temporada <select id="${id}"><option value="todas" ${value === 'todas' ? 'selected' : ''}>Todas</option>${M.seasons.map(s => `<option ${s === value ? 'selected' : ''}>${esc(s)}</option>`).join('')}</select></label>`;
  }

  function pageJogadores(M) {
    if (ui.temporada == null) ui.temporada = M.temporadaAtual || 'todas';
    const games = ui.temporada === 'todas' ? M.games : M.games.filter(g => g.temporada === ui.temporada);
    const st = playerStats(M, games);
    M.elenco.forEach((e, k) => { if (!st.has(k)) st.set(k, { ...e, g: 0, a: 0 }); });
    const list = [...st.values()].map(p => ({ ...p, ga: p.g + p.a }));
    const key = ui.sort, dir = ui.dir;
    // Posição segue a ordem do campo (goleiro → ataque); vazios vão sempre para o fim.
    const ordemPos = p => { const i = POS_ORDEM.indexOf(norm(p)); return i < 0 ? 99 : i; };
    const compara = (a, b) => {
      if (key === 'pos') return dir * (ordemPos(a.pos) - ordemPos(b.pos));
      if (key === 'nome' || key === 'nac') {
        if (!a[key] !== !b[key]) return a[key] ? -1 : 1;
        return dir * String(a[key]).localeCompare(String(b[key]), 'pt-BR');
      }
      return dir * (a[key] - b[key]);
    };
    list.sort((a, b) => compara(a, b) || b.g - a.g || b.a - a.a || a.nome.localeCompare(b.nome, 'pt-BR'));
    const top = k => { const p = [...list].sort((a, b) => b[k] - a[k])[0]; return p && p[k] > 0 ? p : null; };
    const [tg, ta, tga] = [top('g'), top('a'), top('ga')];
    const base = list.filter(p => p.origem === 'base');
    const th = (k, label, cls = '') => `<th data-sort="${k}" class="${cls} ${ui.sort === k ? 'ordenado' : ''}">${label}${ui.sort === k ? (ui.dir < 0 ? ' ▾' : ' ▴') : ''}</th>`;
    return `<div class="secao-topo"><h2>Jogadores</h2>
        <div class="filtros">${seasonSelect('sel-temp', ui.temporada, M)}
        <label>Mostrar <select id="sel-mostrar">${Object.entries(MOSTRAR).map(([k, [label]]) => `<option value="${k}" ${k === ui.mostrar ? 'selected' : ''}>${label}</option>`).join('')}</select></label></div>
      </div>
      <div class="grid grid-4" style="margin-bottom:16px">
        ${destaque('Artilheiro', tg, 'g')}
        ${destaque('Garçom', ta, 'a')}
        ${destaque('Participações em gols', tga, 'ga')}
        ${base.length ? `<div class="card stat"><small>Crias da base</small><strong>${plural(base.length, 'jogador', 'jogadores')}</strong><span>${plural(base.reduce((s, p) => s + p.g, 0), 'gol', 'gols')} · ${plural(base.reduce((s, p) => s + p.a, 0), 'assistência', 'assistências')}</span></div>` : ''}
      </div>
      <div class="tabela-wrap"><table>
        <thead><tr>${th('nome', 'Jogador', 'esq')}${th('pos', 'Pos')}${th('nac', 'País', 'esq')}${th('g', 'Gols')}${th('a', 'Assist.')}${th('ga', 'G+A')}</tr></thead>
        <tbody>${list.filter(MOSTRAR[ui.mostrar][1]).map(p => `<tr class="${p.situacao === 'saiu' ? 'fora' : ''}"><td class="esq"><span class="jogador">${avatar(p)}<span>${esc(p.nome)}${playerTags(p)}</span></span></td><td class="mudo">${esc(p.pos)}</td><td class="esq mudo">${esc(p.nac)}</td><td class="pts">${p.g}</td><td>${p.a}</td><td>${p.ga}</td></tr>`).join('')}</tbody>
      </table></div>
      ${nationalities(M)}`;
  }

  function destaque(rotulo, p, key) {
    return `<div class="card stat destaque"><small>${rotulo}</small>
      <div class="destaque-linha">${p ? avatar(p, 'medio') : ''}<div><strong>${p ? p[key] : 0}</strong><span>${p ? esc(p.nome) : '—'}</span></div></div>
    </div>`;
  }

  // Elenco atual (aba Elenco, sem quem saiu) agrupado por nacionalidade.
  function nationalities(M) {
    const atuais = [...M.elenco.values()].filter(p => p.doElenco && !p.situacao);
    if (!atuais.length) return '';
    const cont = new Map();
    atuais.forEach(p => { const n = p.nac || 'Não informado'; cont.set(n, (cont.get(n) || 0) + 1); });
    const lista = [...cont.entries()].sort((a, b) => b[1] - a[1] || a[0].localeCompare(b[0], 'pt-BR'));
    const max = lista[0][1];
    return `<section class="secao card">
      <div class="secao-topo"><h3>Elenco por nacionalidade</h3><span class="mudo">${plural(atuais.length, 'jogador', 'jogadores')} · ${plural(lista.length, 'país', 'países')}</span></div>
      <ul class="nacionalidades">${lista.map(([n, c]) => `<li><span class="nac-nome">${esc(n)}</span><span class="nac-barra"><span style="width:${Math.round(c / max * 100)}%"></span></span><b>${c}</b></li>`).join('')}</ul>
    </section>`;
  }

  function pageTransferencias(M) {
    if (ui.tempTransf == null) ui.tempTransf = M.temporadaAtual || 'todas';
    const lista = ui.tempTransf === 'todas' ? M.transfers : M.transfers.filter(t => t.temporada === ui.tempTransf);
    const chegadas = lista.filter(t => t.tipo === 'chegada').reverse();
    const saidas = lista.filter(t => t.tipo === 'saida').reverse();
    const soma = l => l.reduce((s, t) => s + (t.valor || 0), 0);
    const gasto = soma(chegadas), receita = soma(saidas), saldo = receita - gasto;
    const valor = t => t.valor != null ? fmtMoney(t.valor) : esc(t.valorTexto || '—');
    const obs = t => t.obs ? `<span class="tag${norm(t.obs).startsWith('empre') ? ' emprestado' : ''}">${esc(t.obs)}</span>` : '';
    // "Sem clube", "Livre", "Base"... não são clubes: só texto, sem escudo.
    const semClube = nome => ['semclube', 'livre', 'base', 'aposentado', 'aposentou', 'dispensado'].includes(norm(nome));
    const clubeLinha = nome => !nome ? '<span class="mudo">—</span>'
      : semClube(nome) ? `<span class="mudo">${esc(nome)}</span>`
      : `<span class="clube-linha">${crest(M, nome, 'mini')}${esc(nome)}</span>`;

    const cards = chegadas.map(t => `<article class="card chegada">
      <div class="chegada-topo">${avatar(t, 'grande')}<div><h3>${esc(t.jogador)}</h3>
        <span class="mudo">${[t.pos, t.idade != null ? `${t.idade} anos` : '', t.nac].filter(Boolean).map(esc).join(' · ')}</span></div></div>
      <div class="chegada-linha"><small>Veio do</small>${clubeLinha(t.clube)}</div>
      <div class="chegada-linha"><small>Valor pago</small><strong class="valor">${valor(t)}</strong>${obs(t)}</div>
      ${t.ep != null ? `<span class="ep-tag">EP ${pad2(t.ep)}</span>` : ''}
    </article>`).join('');

    const linhas = saidas.map(t => `<tr>
      <td class="esq"><span class="jogador">${avatar(t)}<span>${esc(t.jogador)} ${obs(t)}</span></span></td>
      <td class="mudo">${esc(t.pos)}</td>
      <td class="esq">${clubeLinha(t.clube)}</td>
      <td class="pts">${valor(t)}</td>
      <td class="mudo">${t.ep != null ? `EP ${pad2(t.ep)}` : ''}</td></tr>`).join('');

    return `<div class="secao-topo"><h2>Transferências</h2>${seasonSelect('sel-transf', ui.tempTransf, M)}</div>
      <div class="grid grid-3" style="margin-bottom:8px">
        <div class="card stat"><small>Investido em chegadas</small><strong>${fmtMoney(gasto)}</strong><span>${plural(chegadas.length, 'chegada', 'chegadas')}</span></div>
        <div class="card stat"><small>Arrecadado com saídas</small><strong>${fmtMoney(receita)}</strong><span>${plural(saidas.length, 'saída', 'saídas')}</span></div>
        <div class="card stat"><small>Saldo</small><strong class="${saldo >= 0 ? 'positivo' : 'negativo'}">${saldo < 0 ? '−' : ''}${fmtMoney(Math.abs(saldo))}</strong><span>${saldo >= 0 ? 'lucro no mercado' : 'investimento líquido'}</span></div>
      </div>
      <section class="secao"><h3 class="subtitulo">Chegadas</h3>
        ${cards ? `<div class="grid grid-3">${cards}</div>` : '<p class="mudo">Nenhuma chegada registrada nesta temporada.</p>'}
      </section>
      <section class="secao"><h3 class="subtitulo">Saídas</h3>
        ${linhas ? `<div class="tabela-wrap"><table>
          <thead><tr><th class="esq">Jogador</th><th>Pos</th><th class="esq">Destino</th><th>Valor recebido</th><th>Ep.</th></tr></thead>
          <tbody>${linhas}</tbody></table></div>` : '<p class="mudo">Nenhuma saída registrada nesta temporada.</p>'}
      </section>`;
  }

  const statusCls = s => { const n = norm(s); return n.startsWith('conclu') ? 'concluida' : n.startsWith('falh') ? 'falhou' : 'andamento'; };
  const ALVO_IC = '<svg viewBox="0 0 24 24" width="44" height="44" fill="none" stroke="currentColor" stroke-width="1.8" aria-hidden="true"><circle cx="12" cy="12" r="9"/><circle cx="12" cy="12" r="5"/><circle cx="12" cy="12" r="1.5" fill="currentColor"/></svg>';

  const SUL_AMERICA = ['argentina', 'bolivia', 'brasil', 'chile', 'colombia', 'equador', 'paraguai', 'peru', 'uruguai', 'venezuela'];
  // Contagens automáticas pela coluna Chave da aba Metas.
  function metaAtual(M, m) {
    if (m.chave === 'sulamericanos') return [...M.elenco.values()].filter(p => p.doElenco && !p.situacao && SUL_AMERICA.includes(norm(p.nac))).length;
    return m.atual;
  }

  function metaCard(M, m) {
    const src = !m.imagem ? '' : /^(https?:)?\/\/|\//.test(m.imagem) ? m.imagem : `assets/metas/${m.imagem}`;
    const atual = metaAtual(M, m);
    const batida = m.alvo != null && atual != null && atual >= m.alvo;
    const status = batida && !m.regra ? 'Concluída' : m.status;
    const cls = m.regra ? 'regra' : statusCls(status);
    const pct = m.alvo ? Math.min(100, Math.round((atual || 0) / m.alvo * 100)) : null;
    // Observação com ";" ou quebra de linha vira lista (ex.: faixas da regra de transferências).
    const itens = m.obs.split(/\s*;\s*|\n/).map(s => s.trim()).filter(Boolean);
    const texto = itens.length > 1 ? `<ul class="meta-lista">${itens.map(i => `<li>${esc(i)}</li>`).join('')}</ul>` : itens.length ? `<p>${esc(itens[0])}</p>` : '';
    return `<article class="card meta meta-${cls}">
      <div class="meta-img">${ALVO_IC}${src ? `<img src="${esc(src)}" alt="" loading="lazy" onerror="this.remove()">` : ''}</div>
      <div class="meta-corpo">
        <span class="status ${cls}">${m.regra ? 'Regra da carreira' : esc(status)}</span>
        <h3>${esc(m.meta)}</h3>
        ${texto}
        ${pct != null ? `<div class="barra"><span style="width:${pct}%"></span></div><small class="mudo">${atual ?? 0} de ${m.alvo}${m.chave ? ' · contagem automática do elenco' : ''}</small>` : ''}
      </div>
    </article>`;
  }

  function pageMetas(M) {
    // Metas sem temporada valem para a temporada atual.
    const todas = M.metas.filter(m => !m.temporada || m.temporada === M.temporadaAtual);
    const metas = todas.filter(m => !m.regra), regras = todas.filter(m => m.regra);
    const statusDe = m => { const a = metaAtual(M, m); return m.alvo != null && a != null && a >= m.alvo ? 'concluida' : statusCls(m.status); };
    const conta = c => metas.filter(m => statusDe(m) === c).length;
    const nome = M.temporadaInfo.nome ? `${esc(M.temporadaInfo.nome)} · ` : '';
    return `<div class="secao-topo"><div><h2>Metas da temporada</h2><span class="mudo">${nome}${esc(M.temporadaAtual)}</span></div>
        <span class="mudo">${conta('concluida')} concluídas · ${conta('andamento')} em andamento · ${conta('falhou')} não alcançadas</span></div>
      ${metas.length ? `<div class="grid grid-3">${metas.map(m => metaCard(M, m)).join('')}</div>` : '<p class="estado">Cadastre as metas na aba Metas da planilha.</p>'}
      ${regras.length ? `<section class="secao"><div class="secao-topo"><h2>Regras da carreira</h2><span class="mudo">O que vale em toda a série</span></div>
        <div class="grid grid-3">${regras.map(m => metaCard(M, m)).join('')}</div></section>` : ''}`;
  }

  function pageRecordes(M) {
    const car = careerRecords(M);
    const reais = M.recordes.find(r => r.chave === 'titulos');
    const trofeuIc = '<span class="trofeu-ic" aria-hidden="true"><svg width="20" height="20" viewBox="0 0 24 24" fill="currentColor"><path d="M7 3h10v2h3v3a4 4 0 0 1-4 4h-.3A5 5 0 0 1 13 14.9V18h3v3H8v-3h3v-3.1A5 5 0 0 1 8.3 12H8a4 4 0 0 1-4-4V5h3V3zm0 4H6v1a2 2 0 0 0 1 1.7V7zm10 0v2.7A2 2 0 0 0 18 8V7h-1z"/></svg></span>';
    return `<div class="pagina-historia">
      <header class="historia-topo">
        <small>Arquivo do clube · desde 1945</small>
        <h2>Recordes históricos</h2>
        <p>Os marcos reais do VfL Wolfsburg até o início da série, frente ao que a carreira já escreveu.</p>
      </header>
      <div class="grid grid-3">${M.recordes.map(r => recordCard(M, r, car)).join('')}</div>

      <section class="secao">
        <div class="card"><h3>Sala de troféus</h3>
          <p class="mudo" style="margin:6px 0 4px">Na carreira</p>
          <ul class="lista-simples">${M.titulos.map(t => `<li><div class="trofeu">${trofeuIc}<div><b>${esc(t.comp)}</b><small>${esc(t.temporada)}${t.obs ? ' · ' + esc(t.obs) : ''}</small></div></div></li>`).join('') || '<li class="mudo">Nenhum título ainda. A história está sendo escrita.</li>'}</ul>
          ${reais && reais.obs ? `<p class="mudo" style="margin:14px 0 4px">História real do clube</p><ul class="lista-simples">${reais.obs.split(/·|;|\n/).map(s => s.trim()).filter(Boolean).map(s => `<li><div class="trofeu">${trofeuIc}<b>${esc(s)}</b></div></li>`).join('')}</ul>` : ''}
        </div>
      </section>
    </div>`;
  }

  // ---------- roteador ----------
  let MODEL = null;
  const pages = { inicio: pageInicio, episodios: pageEpisodios, metas: pageMetas, jogos: pageJogos, jogadores: pageJogadores, transferencias: pageTransferencias, recordes: pageRecordes };
  let lastPage = null;

  function render() {
    const [p, arg] = location.hash.replace(/^#/, '').split('/');
    const page = pages[p] ? p : 'inicio';
    document.querySelectorAll('#nav a').forEach(a => a.classList.toggle('ativo', a.dataset.page === page));
    $app.innerHTML = pages[page](MODEL, arg);
    startSlides();
    startCountdown();
    if (page !== lastPage) window.scrollTo(0, 0);
    lastPage = page;
  }

  $app.addEventListener('change', ev => {
    if (ev.target.id === 'sel-temp') { ui.temporada = ev.target.value; render(); }
    if (ev.target.id === 'sel-mostrar') { ui.mostrar = ev.target.value; render(); }
    if (ev.target.id === 'sel-transf') { ui.tempTransf = ev.target.value; render(); }
  });
  $app.addEventListener('click', ev => {
    const th = ev.target.closest('th[data-sort]');
    if (!th) return;
    const k = th.dataset.sort;
    // Texto e posição começam em ordem crescente; números começam do maior.
    if (ui.sort === k) ui.dir *= -1; else { ui.sort = k; ui.dir = ['nome', 'pos', 'nac'].includes(k) ? 1 : -1; }
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
    document.getElementById('sub-portal').textContent = `${M.clube} · Modo Carreira Realista · FC 27`;
    if (M.cfg.canal) document.getElementById('rodape-canal').textContent = M.cfg.canal;
    if (M.cfg.linkcanal) document.getElementById('rodape-canal').href = document.getElementById('link-youtube').href = M.cfg.linkcanal;
    // Escudo do topo: o da aba Clube; senão o encontrado pela automação; senão a inicial.
    const escudo = document.getElementById('escudo');
    const src = String(M.cfg.escudo || '') || clubInfo(M, M.clube).escudo;
    escudo.textContent = (M.clube.split(/\s+/).pop() || 'C')[0].toUpperCase();
    if (src) escudo.insertAdjacentHTML('beforeend', `<img src="${esc(src)}" alt="Escudo do ${esc(M.clube)}" referrerpolicy="no-referrer" onerror="this.remove()">`);

    const avisos = [];
    if (sim(M.cfg.dadosexemplo)) avisos.push('<div class="aviso"><b>Dados de exemplo.</b> Jogadores, placares e tabelas são fictícios, só para mostrar o formato do portal.</div>');
    if (M.warnings.length) avisos.push(`<details class="aviso"><summary>Conferência da planilha: ${plural(M.warnings.length, 'ponto', 'pontos')} para revisar</summary><ul>${M.warnings.map(w => `<li>${esc(w)}</li>`).join('')}</ul></details>`);
    $avisos.innerHTML = avisos.join('');
  }

  async function init() {
    try {
      if (typeof XLSX === 'undefined') throw new Error('A biblioteca de leitura de planilhas não carregou. Verifique a conexão.');
      const bust = `?v=${Date.now()}`;
      // Arquivos de apoio são opcionais: sem eles, o site usa só a planilha.
      const optional = url => fetch(url + bust, { cache: 'no-store' }).then(r => r.ok ? r.json() : null).catch(() => null);
      const extrasReq = Promise.all([VIDEOS_URL, FOTOS_URL, CLUBES_URL, TREINADOR_URL, FOTOS_CORPO_URL, TEMPORADAS_URL].map(optional));
      const res = await fetch(DATA_URL + bust, { cache: 'no-store' });
      if (!res.ok) throw new Error(`Não foi possível baixar ${DATA_URL} (HTTP ${res.status}).`);
      const wb = XLSX.read(await res.arrayBuffer(), { type: 'array' });
      const [videos, fotos, clubesJson, treinador, fotosCorpo, temporadas] = await extrasReq;
      MODEL = build(wb, { videos, fotos, clubesJson, treinador, fotosCorpo, temporadas });
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
