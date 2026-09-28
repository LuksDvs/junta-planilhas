/* Núcleo: leitura, limpeza e junção das planilhas. Depende do XLSX (SheetJS). */
(function (raiz) {
  const LIMITE_LINHAS_EXCEL = 1048576;

  const semAcento = (s) => s.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  const vazio = (v) => v === '' || v === null || v === undefined || (typeof v === 'string' && v.trim() === '');
  const pausa = () => new Promise((r) => setTimeout(r, 0));

  function normalizar(v, flex) {
    let s = String(v ?? '').replace(/\s+/g, ' ').trim();
    if (flex) s = semAcento(s).toLowerCase();
    return s;
  }

  function semExtensao(nome) {
    return nome.replace(/\.[^.]+$/, '');
  }

  /* ---------- CSV ---------- */

  function contarFora(linha, sep) {
    let n = 0, aspas = false;
    for (const ch of linha) {
      if (ch === '"') aspas = !aspas;
      else if (ch === sep && !aspas) n++;
    }
    return n;
  }

  function detectarSeparador(texto) {
    const linhas = texto.slice(0, 100000).split(/\r?\n/).filter((l) => l.trim()).slice(0, 10);
    let melhor = ',', pontos = -1;
    for (const sep of [';', ',', '\t', '|']) {
      const cont = linhas.map((l) => contarFora(l, sep));
      const primeiro = cont[0] || 0;
      if (!primeiro) continue;
      const p = cont.filter((c) => c === primeiro).length * 1000 + primeiro;
      if (p > pontos) { pontos = p; melhor = sep; }
    }
    return melhor;
  }

  function detectarDecimalVirgula(texto, sep) {
    if (sep === ',') return false;
    const celulas = texto.slice(0, 200000).split(/\r?\n/).slice(0, 500)
      .flatMap((l) => l.split(sep)).map((c) => c.replace(/"/g, '').trim());
    let virgula = 0, ponto = 0;
    for (const c of celulas) {
      if (/^-?(\d{1,3}(\.\d{3})+|\d+),\d+$/.test(c)) virgula++;
      else {
        const m = c.match(/^-?\d+\.(\d+)$/);
        if (m && m[1].length !== 3) ponto++; // "1.234" é ambíguo, não vota
      }
    }
    return virgula > ponto || (virgula === ponto && sep === ';');
  }

  function converterValorCSV(v, decimalVirgula) {
    if (typeof v !== 'string') return v;
    const s = v.trim();
    if (s === '') return '';
    if (!/^-?[\d.,]+$/.test(s)) return v;
    const neg = s.startsWith('-');
    const corpo = neg ? s.slice(1) : s;
    if (/^0\d/.test(corpo)) return v;                 // zero à esquerda: código, CEP etc.
    if (/^\d+$/.test(corpo)) return corpo.length >= 11 ? v : Number(s); // CPF, EAN, CNPJ ficam texto
    let num = null;
    if (decimalVirgula) {
      if (/^(\d{1,3}(\.\d{3})+|\d+)(,\d+)?$/.test(corpo)) num = Number(corpo.replace(/\./g, '').replace(',', '.'));
    } else if (/^(\d{1,3}(,\d{3})+|\d+)(\.\d+)?$/.test(corpo)) {
      num = Number(corpo.replace(/,/g, ''));
    }
    if (num === null || !isFinite(num)) return v;
    return neg ? -num : num;
  }

  /* ---------- Leitura ---------- */

  function interpretar(nome, buffer) {
    const ehTexto = /\.(csv|txt)$/i.test(nome);
    let wb, decimalVirgula = false;
    if (ehTexto) {
      let texto;
      try { texto = new TextDecoder('utf-8', { fatal: true }).decode(buffer); }
      catch (e) { texto = new TextDecoder('windows-1252').decode(buffer); }
      texto = texto.replace(/^\uFEFF/, '');
      const sep = detectarSeparador(texto);
      decimalVirgula = detectarDecimalVirgula(texto, sep);
      wb = XLSX.read(texto, { type: 'string', FS: sep, raw: true });
    } else {
      wb = XLSX.read(new Uint8Array(buffer), { type: 'array', cellDates: true });
    }
    if (!wb.SheetNames.length) throw new Error('Nenhuma aba encontrada no arquivo.');
    return wb.SheetNames.map((abaNome) => {
      let aoa = XLSX.utils.sheet_to_json(wb.Sheets[abaNome], { header: 1, defval: '', blankrows: false, raw: true });
      if (ehTexto) aoa = aoa.map((l) => l.map((v) => converterValorCSV(v, decimalVirgula)));
      return { nome: ehTexto ? semExtensao(nome) : abaNome, aoa };
    });
  }

  /* ---------- Preparação de cada aba ---------- */

  function prepararBloco(arquivo, aba, aoa, opts, avisos) {
    if (!aoa.length) { avisos.push(`${arquivo} › ${aba}: aba vazia, ignorada.`); return null; }

    let inicio = 0;
    if (opts.usarCab) {
      // Pula linhas de título acima do cabeçalho (comum em exportações de ERP)
      const cont = aoa.slice(0, 15).map((l) => l.filter((v) => !vazio(v)).length);
      const max = Math.max(...cont);
      const limite = Math.max(2, Math.ceil(max * 0.6));
      const i = cont.findIndex((c) => c >= limite);
      inicio = i < 0 ? 0 : i;
      if (inicio > 0) avisos.push(`${arquivo} › ${aba}: ${inicio} linha(s) de título acima do cabeçalho foram ignoradas.`);
    }

    const cabBruto = opts.usarCab ? aoa[inicio] : null;
    let linhas = aoa.slice(opts.usarCab ? inicio + 1 : 0);

    let largura = cabBruto ? cabBruto.length : 0;
    for (const l of linhas) if (l.length > largura) largura = l.length;

    const util = new Array(largura).fill(false);
    if (cabBruto) cabBruto.forEach((v, i) => { if (!vazio(v)) util[i] = true; });
    for (const l of linhas) for (let i = 0; i < l.length; i++) if (!util[i] && !vazio(l[i])) util[i] = true;

    let indices;
    if (opts.usarCab) {
      indices = [];
      util.forEach((u, i) => { if (u) indices.push(i); });   // descarta colunas totalmente vazias
    } else {
      const ultima = util.lastIndexOf(true);                  // sem cabeçalho, a posição importa
      indices = Array.from({ length: ultima + 1 }, (_, i) => i);
    }
    if (!indices.length) { avisos.push(`${arquivo} › ${aba}: sem dados, ignorada.`); return null; }

    const cabecalho = indices.map((i) => {
      const s = cabBruto ? String(cabBruto[i] ?? '').replace(/\s+/g, ' ').trim() : '';
      return s || `Coluna ${i + 1}`;
    });
    linhas = linhas.map((l) => indices.map((i) => (l[i] === undefined ? '' : l[i])));

    let repetidos = 0;
    if (opts.usarCab && opts.removerRepetidos) {
      const cabNorm = cabecalho.map((c) => normalizar(c, true));
      const minimo = Math.min(2, cabecalho.length);
      linhas = linhas.filter((l) => {
        let preenchidas = 0, iguais = 0;
        l.forEach((v, i) => {
          if (vazio(v)) return;
          preenchidas++;
          if (normalizar(v, true) === cabNorm[i]) iguais++;
        });
        const eh = preenchidas >= minimo && iguais === preenchidas;
        if (eh) repetidos++;
        return !eh;
      });
    }

    return { arquivo, aba, cabecalho, linhas, repetidos };
  }

  /* ---------- Junção alinhada pelo nome da coluna ---------- */

  function combinar(blocos, opts) {
    const colunas = [];
    const porChave = new Map();
    const comAba = opts.origem && blocos.some((b) => b.multiplasAbas);
    const fontes = blocos.map((b) => (b.multiplasAbas ? `${b.arquivo} › ${b.aba}` : b.arquivo));

    const mapas = blocos.map((b, bi) => {
      const usados = new Set();
      return b.cabecalho.map((nome) => {
        const base = normalizar(nome, opts.flex);
        let chave = base, n = 1;
        while (usados.has(chave)) { n++; chave = `${base}#${n}`; }
        usados.add(chave);
        let idx = porChave.get(chave);
        if (idx === undefined) {
          idx = colunas.length;
          porChave.set(chave, idx);
          colunas.push({ rotulo: n > 1 ? `${nome} (${n})` : nome, variantes: new Set([nome]), fontes: new Set() });
        } else {
          colunas[idx].variantes.add(nome);
        }
        colunas[idx].fontes.add(bi);
        return idx;
      });
    });

    const prefixo = opts.origem ? (comAba ? ['Arquivo Origem', 'Aba Origem'] : ['Arquivo Origem']) : [];
    const aoa = [prefixo.concat(colunas.map((c) => c.rotulo))];
    blocos.forEach((b, bi) => {
      const mapa = mapas[bi];
      const origem = opts.origem ? (comAba ? [b.arquivo, b.aba] : [b.arquivo]) : [];
      for (const l of b.linhas) {
        const linha = new Array(colunas.length).fill('');
        for (let i = 0; i < mapa.length; i++) linha[mapa[i]] = l[i];
        aoa.push(origem.length ? origem.concat(linha) : linha);
      }
    });

    const relatorio = colunas.map((c) => ({
      rotulo: c.rotulo,
      variantes: [...c.variantes],
      ausentesEm: fontes.filter((_, i) => !c.fontes.has(i)),
    }));
    return { aoa, relatorio, totalFontes: blocos.length, linhas: aoa.length - 1 };
  }

  function nomeAbaSeguro(nome, usados) {
    const base = String(nome).replace(/[\[\]:*?\/\\]/g, '_').replace(/^'+|'+$/g, '').trim().slice(0, 31) || 'Planilha';
    let final = base, i = 2;
    while (usados.has(final.toLowerCase())) {
      const suf = `_${i++}`;
      final = base.slice(0, 31 - suf.length) + suf;
    }
    usados.add(final.toLowerCase());
    return final;
  }

  async function montarResultado(arquivos, modo, opts, aoProgredir) {
    const avisos = [];
    const blocos = [];
    const total = arquivos.reduce((s, a) => s + a.abas.length, 0) || 1;
    let feitos = 0;

    for (const arq of arquivos) {
      const multi = arq.abas.length > 1;
      for (const aba of arq.abas) {
        const b = prepararBloco(arq.nome, aba.nome, aba.aoa, opts, avisos);
        if (b) { b.multiplasAbas = multi; blocos.push(b); }
        feitos++;
        aoProgredir && aoProgredir((feitos / total) * 0.6);
        await pausa();
      }
    }
    if (!blocos.length) throw new Error('Nenhum dado encontrado nos arquivos carregados.');

    const repetidos = blocos.reduce((s, b) => s + b.repetidos, 0);
    if (repetidos) avisos.push(`${repetidos.toLocaleString('pt-BR')} linha(s) de cabeçalho repetido foram removidas do meio dos dados.`);

    const usados = new Set();
    const abas = [];
    if (modo === 'empilhar') {
      abas.push({ nome: nomeAbaSeguro('Unificado', usados), ...combinar(blocos, opts) });
    } else if (modo === 'separado') {
      for (const b of blocos) {
        const nome = b.multiplasAbas ? `${semExtensao(b.arquivo)} - ${b.aba}` : semExtensao(b.arquivo);
        abas.push({ nome: nomeAbaSeguro(nome, usados), ...combinar([b], opts) });
      }
    } else {
      const grupos = new Map();
      for (const b of blocos) {
        const chave = normalizar(b.aba, true);
        if (!grupos.has(chave)) grupos.set(chave, []);
        grupos.get(chave).push(b);
      }
      for (const lista of grupos.values()) {
        abas.push({ nome: nomeAbaSeguro(lista[0].aba, usados), ...combinar(lista, opts) });
      }
    }
    aoProgredir && aoProgredir(1);

    for (const a of abas) {
      if (a.aoa.length > LIMITE_LINHAS_EXCEL) {
        avisos.push(`A aba "${a.nome}" passa do limite do Excel (1.048.576 linhas) e será dividida em partes no arquivo final.`);
      }
    }
    return { abas, avisos, totalLinhas: abas.reduce((s, a) => s + a.linhas, 0) };
  }

  /* ---------- Exportação ---------- */

  function larguras(aoa) {
    const amostra = aoa.slice(0, 300);
    const n = aoa[0] ? aoa[0].length : 0;
    const cols = [];
    for (let c = 0; c < n; c++) {
      let max = 6;
      for (const l of amostra) {
        const v = l[c];
        const len = v instanceof Date ? 10 : String(v ?? '').length;
        if (len > max) max = len;
      }
      cols.push({ wch: Math.min(50, max + 2) });
    }
    return cols;
  }

  function adicionarAba(wb, nome, aoa) {
    const ws = XLSX.utils.aoa_to_sheet(aoa, { dateNF: 'dd/mm/yyyy' });
    ws['!cols'] = larguras(aoa);
    if (ws['!ref']) ws['!autofilter'] = { ref: ws['!ref'] };
    XLSX.utils.book_append_sheet(wb, ws, nome);
  }

  function gerarWorkbook(resultado) {
    const wb = XLSX.utils.book_new();
    const usados = new Set();
    for (const a of resultado.abas) {
      if (a.aoa.length <= LIMITE_LINHAS_EXCEL) {
        adicionarAba(wb, nomeAbaSeguro(a.nome, usados), a.aoa);
      } else {
        const cab = a.aoa[0];
        const porParte = LIMITE_LINHAS_EXCEL - 1;
        for (let i = 1, parte = 1; i < a.aoa.length; i += porParte, parte++) {
          const nome = nomeAbaSeguro(parte === 1 ? a.nome : `${a.nome.slice(0, 26)}_p${parte}`, usados);
          adicionarAba(wb, nome, [cab].concat(a.aoa.slice(i, i + porParte)));
        }
      }
    }
    return wb;
  }

  function nomeArquivoSeguro(texto, padrao) {
    let s = String(texto || '').trim().replace(/\.xlsx$/i, '');
    s = s.replace(/[<>:"/\\|?*\x00-\x1f]/g, '_').replace(/[. ]+$/, '');
    return (s || padrao) + '.xlsx';
  }

  raiz.Nucleo = {
    normalizar, detectarSeparador, detectarDecimalVirgula, converterValorCSV,
    interpretar, prepararBloco, combinar, montarResultado, gerarWorkbook,
    nomeArquivoSeguro, nomeAbaSeguro,
  };
})(typeof window !== 'undefined' ? window : globalThis);
