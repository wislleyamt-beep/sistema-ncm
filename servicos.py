"""
Módulo de classificação de SERVIÇOS para a NFS-e Nacional.

Monta, a partir das tabelas oficiais, a classificação completa de um serviço:
item da LC 116 -> código de tributação nacional (cTribNac) -> NBS -> indOp
-> cClassTrib / CST IBS-CBS -> anexo do Simples Nacional / fator R.

Fontes (pasta dados_servicos/):
- anexo_viii-tabela-correlacao-*.xlsx : Anexo VIII da NFS-e (correlação Item
  LC 116 x NBS x cIndOp x cClassTrib x atividade do Simples), publicado no
  portal gov.br/nfse. Fonte OFICIAL e principal.
- ctribnac.csv : códigos de tributação nacional (portal gov.br/nfse).
- cnae_lc116.csv : correlação CNAE x item LC 116 (referência de mercado,
  repositório público luizbasca/tabela-correlacao-fiscal-servico). Não há
  tabela nacional oficial CNAE x LC 116; cada município publica a sua.
- inss_lc116.csv : retenção de 11% de INSS por item (IN RFB 2.110/2022),
  mesma referência de mercado.
"""
import csv
import glob
import os
import re
import unicodedata

from openpyxl import load_workbook

_DIR = os.path.join(os.path.dirname(__file__), 'dados_servicos')

# ---------------------------------------------------------------------------
# Simples Nacional: códigos de atividade do Anexo VIII (cAtivSN / COD_ATIV_SN)
# ---------------------------------------------------------------------------
ATIV_SN = {
    '7':  {'anexo': 'III', 'fator_r': False, 'rotulo': 'Anexo III'},
    '8':  {'anexo': 'III', 'fator_r': False, 'rotulo': 'Anexo III (contábil com ISS fixo)'},
    '9':  {'anexo': 'III ou V', 'fator_r': True, 'rotulo': 'Fator R: Anexo III ou V'},
    '10': {'anexo': 'III', 'fator_r': False, 'rotulo': 'Anexo III (transporte municipal de passageiros)'},
    '11': {'anexo': 'III', 'fator_r': False, 'rotulo': 'Anexo III (locação de bens móveis / sem ISS)'},
    '12': {'anexo': 'III', 'fator_r': False, 'rotulo': 'Anexo III (construção civil 7.02/7.05)'},
    '13': {'anexo': 'IV', 'fator_r': False, 'rotulo': 'Anexo IV (construção civil 7.02/7.05)'},
    '14': {'anexo': 'IV', 'fator_r': False, 'rotulo': 'Anexo IV'},
    'IMP': {'anexo': None, 'fator_r': False, 'rotulo': 'Impeditiva ao Simples Nacional'},
}

# Tabelas do Simples (LC 123/2006, anexos III, IV e V): (limite RBT12, alíquota nominal, parcela a deduzir)
TABELAS_SN = {
    'III': [(180000, 0.06, 0), (360000, 0.112, 9360), (720000, 0.135, 17640),
            (1800000, 0.16, 35640), (3600000, 0.21, 125640), (4800000, 0.33, 648000)],
    'IV':  [(180000, 0.045, 0), (360000, 0.09, 8100), (720000, 0.102, 12420),
            (1800000, 0.14, 39780), (3600000, 0.22, 183780), (4800000, 0.33, 828000)],
    'V':   [(180000, 0.155, 0), (360000, 0.18, 4500), (720000, 0.195, 9900),
            (1800000, 0.205, 17100), (3600000, 0.23, 62100), (4800000, 0.305, 540000)],
}

# Descrições do cIndOp usados no Anexo VIII (Anexo VII da NFS-e, v1.03.00)
INDOP_DESC = {
    '010101': 'Bem móvel material, presencial, com retirada no estabelecimento do fornecedor',
    '010102': 'Bem móvel material, presencial, com retirada fora do estabelecimento do fornecedor',
    '010103': 'Bem móvel material, não presencial, com entrega em endereço fornecido',
    '010106': 'Locação de bem móvel material em aquisição centralizada (contribuinte do regime regular com mais de um estabelecimento)',
    '020101': 'Operação com bem imóvel ou direito relacionado a bem imóvel',
    '020102': 'Operação com bem imóvel de características especiais (BICE)',
    '020201': 'Serviços prestados fisicamente sobre bem imóvel',
    '020202': 'Serviços sobre bens imóveis de características especiais (BICE)',
    '020301': 'Administração e intermediação de bens imóveis',
    '020401': 'Administração e intermediação de bens imóveis de características especiais (BICE)',
    '030101': 'Serviço prestado fisicamente sobre a pessoa física, no estabelecimento do fornecedor',
    '030102': 'Serviço prestado fisicamente sobre a pessoa física, fora do estabelecimento do fornecedor',
    '040101': 'Planejamento, organização e administração de feiras, exposições, congressos, espetáculos e eventos',
    '050101': 'Serviço prestado fisicamente sobre bem móvel material, no estabelecimento do fornecedor',
    '050102': 'Serviço prestado fisicamente sobre bem móvel material, fora do estabelecimento do fornecedor',
    '050201': 'Serviços portuários',
    '060101': 'Transporte de passageiros',
    '070101': 'Transporte de carga',
    '080101': 'Exploração de via',
    '100101': 'Cessão de espaço para prestação de serviços',
    '100301': 'Demais serviços, em operações onerosas',
    '100501': 'Demais bens, em operações onerosas (código de bens: o Anexo VII o marca como não aplicável à NFS-e)',
    '130101': 'Transporte dutoviário de gás natural (capacidade)',
    '130201': 'Transporte dutoviário de gás natural (contratação de capacidade)',
}

# Regra do indOp para operação NÃO onerosa e/ou adquirente no EXTERIOR
# (aba "REGRA OP NÃO ONE_EXT" do Anexo VIII v1.02.00)
REGRA_INDOP = {
    '100301': {'exterior': '100302', 'nao_onerosa': '100401'},
    '100101': {'exterior': '100102', 'nao_onerosa': '100201'},
    '100501': {'exterior': '100502', 'nao_onerosa': '100601'},
}

# LC 116/2003, art. 3º: ISS devido no local da prestação (exceções à regra do estabelecimento do prestador)
ISS_LOCAL_PRESTACAO = {
    '03.05', '07.02', '07.04', '07.05', '07.09', '07.10', '07.11', '07.12', '07.16',
    '07.17', '07.18', '07.19', '11.01', '11.02', '11.03', '11.04', '16.01', '16.02',
    '17.05', '17.10', '20.01', '20.02', '20.03', '04.22', '04.23', '05.09', '15.01',
    '15.09', '10.04',
}
# LC 116/2003, art. 6º, § 2º, II: tomador PJ é responsável pela retenção do ISS
ISS_RETENCAO_TOMADOR = {
    '03.05', '07.02', '07.04', '07.05', '07.09', '07.10', '07.12', '07.14', '07.15',
    '07.16', '07.17', '07.19', '11.02', '17.05', '17.10',
}


def _norm(txt):
    txt = unicodedata.normalize('NFKD', str(txt or '')).encode('ascii', 'ignore').decode()
    return re.sub(r'\s+', ' ', txt.lower()).strip()


def _cod(v, n=6):
    if v is None or v == '':
        return None
    s = str(v).strip()
    if s.upper() == 'IMP':
        return 'IMP'
    if re.fullmatch(r'\d+(\.0)?', s):
        s = s.split('.')[0]
        return s.zfill(n) if n else s
    return s


def _limpa(v):
    return re.sub(r'\s+', ' ', str(v)).strip() if v is not None else ''


def _item_de_ctrib(ctrib):
    """'070202' -> '07.02' ; '990301' -> '99.03'."""
    return f'{ctrib[:2]}.{ctrib[2:4]}'


# ---------------------------------------------------------------------------
# Carga das tabelas
# ---------------------------------------------------------------------------
def _carrega_anexo_viii():
    arquivos = sorted(glob.glob(os.path.join(_DIR, 'anexo_viii-*.xlsx')))
    if not arquivos:
        return {}, None
    arquivo = arquivos[-1]
    versao = re.search(r'v(\d{2})-(\d{2})-(\d{2})', arquivo)
    versao = '.'.join(versao.groups()) if versao else os.path.basename(arquivo)
    ws = load_workbook(arquivo, read_only=True, data_only=True)['tabela geral']
    linhas = list(ws.iter_rows(values_only=True))[1:]

    itens = {}
    item = None
    desdobro_atual = None
    grupo = []          # NBS do grupo atual (recebem as alternativas)
    ultimo = None       # última NBS (para herdar campos em branco)
    for r in linhas:
        r = list(r) + [None] * 12
        cod_item, desc_item, nbs, desc_nbs, ativ, _desc_ativ, indop, local, ctrib, nome_ctrib = r[:10]
        if cod_item:
            cod_item = _limpa(cod_item)
            # 99.03.01 -> item 99.03 com desdobro 01
            partes = cod_item.split('.')
            chave = '.'.join(partes[:2])
            item = itens.setdefault(chave, {
                'item': chave,
                'descricao': _limpa(desc_item),
                'desdobros_anexo': [],
                'nbs': [],
            })
            if len(partes) == 3:
                item['desdobros_anexo'].append({'codigo': ''.join(partes), 'descricao': _limpa(desc_item)})
            grupo, ultimo = [], None
            desdobro_atual = ''.join(partes) if len(partes) == 3 else None
        if item is None:
            continue
        ctrib = _cod(ctrib)
        ativ = _cod(ativ, 0)
        indop = _cod(indop)
        if nbs:
            entrada = {
                'nbs': _limpa(nbs),
                'descricao': _limpa(desc_nbs),
                'ativ_sn': ativ or (ultimo or {}).get('ativ_sn'),
                'indop': indop or (ultimo or {}).get('indop'),
                'local_ibs': _limpa(local) or (ultimo or {}).get('local_ibs', ''),
                'cclasstrib': ctrib or (ultimo or {}).get('cclasstrib'),
                'nome_cclasstrib': _limpa(nome_ctrib) if ctrib else (ultimo or {}).get('nome_cclasstrib', ''),
                'cclasstrib_alternativos': [],
                'ativ_sn_alternativos': [],
                'indop_alternativos': [],
                'desdobro': desdobro_atual,
            }
            if ctrib or not grupo:
                grupo = []
            grupo.append(entrada)
            item['nbs'].append(entrada)
            ultimo = entrada
        else:
            # linha sem NBS: alternativas para o grupo atual
            for e in grupo:
                if ctrib and ctrib != e['cclasstrib'] and all(a['codigo'] != ctrib for a in e['cclasstrib_alternativos']):
                    e['cclasstrib_alternativos'].append({'codigo': ctrib, 'nome': _limpa(nome_ctrib)})
                if ativ and ativ != e['ativ_sn'] and ativ not in e['ativ_sn_alternativos']:
                    e['ativ_sn_alternativos'].append(ativ)
                if indop and indop != e['indop'] and all(a['codigo'] != indop for a in e['indop_alternativos']):
                    e['indop_alternativos'].append({'codigo': indop, 'local_ibs': _limpa(local)})
    return itens, versao


def _le_csv(nome):
    caminho = os.path.join(_DIR, nome)
    if not os.path.exists(caminho):
        return []
    with open(caminho, encoding='utf-8') as f:
        return list(csv.DictReader(f, delimiter=';'))


ITENS, VERSAO_ANEXO_VIII = _carrega_anexo_viii()

CTRIBNAC = {}
for _r in _le_csv('ctribnac.csv'):
    CTRIBNAC.setdefault(_item_de_ctrib(_r['codigo']), []).append(
        {'codigo': _r['codigo'], 'descricao': _r['descricao']})

CNAES = {}            # cnae -> {descricao, itens[]}
CNAES_POR_ITEM = {}   # item -> [cnae]
for _r in _le_csv('cnae_lc116.csv'):
    c = CNAES.setdefault(_r['cnae'], {'cnae': _r['cnae'], 'descricao': _r['descricao_cnae'], 'itens': []})
    if _r['item_lc116'] not in c['itens']:
        c['itens'].append(_r['item_lc116'])
    CNAES_POR_ITEM.setdefault(_r['item_lc116'], []).append(_r['cnae'])

INSS = {_r['item_lc116']: _r for _r in _le_csv('inss_lc116.csv')}

# Índice de busca textual por item
_INDICE = {}
for _k, _it in ITENS.items():
    partes = [_k, _it['descricao']]
    partes += [d['descricao'] for d in CTRIBNAC.get(_k, [])]
    partes += [n['nbs'] + ' ' + n['descricao'] for n in _it['nbs']]
    partes += [c + ' ' + CNAES[c]['descricao'] for c in CNAES_POR_ITEM.get(_k, [])]
    _INDICE[_k] = _norm(' '.join(partes))


# ---------------------------------------------------------------------------
# Consultas
# ---------------------------------------------------------------------------
def _info_sn(cod):
    if cod is None:
        return None
    base = ATIV_SN.get(str(cod), {'anexo': None, 'fator_r': False, 'rotulo': f'Código {cod}'})
    return dict(base, codigo=str(cod))


def _resumo_simples(item):
    codigos = []
    for n in item['nbs']:
        for c in [n['ativ_sn']] + n['ativ_sn_alternativos']:
            if c and c not in codigos:
                codigos.append(c)
    infos = [_info_sn(c) for c in codigos]
    return {
        'codigos': infos,
        'fator_r': any(i['fator_r'] for i in infos),
        'impeditiva': any(i['codigo'] == 'IMP' for i in infos),
        'varia_por_nbs': len(codigos) > 1,
    }


def detalhe_item(chave):
    chave = normaliza_item(chave)
    item = ITENS.get(chave)
    if not item:
        return None
    nbs = []
    for n in item['nbs']:
        d = dict(n)
        d['ativ_sn_info'] = _info_sn(n['ativ_sn'])
        d['ativ_sn_alternativos_info'] = [_info_sn(a) for a in n['ativ_sn_alternativos']]
        d['indop_desc'] = INDOP_DESC.get(n['indop'] or '', '')
        d['indop_regra'] = REGRA_INDOP.get(n['indop'] or '')
        d['indop_alternativos'] = [dict(a, descricao=INDOP_DESC.get(a['codigo'], '')) for a in n['indop_alternativos']]
        d['cst_ibs_cbs'] = (n['cclasstrib'] or '')[:3] or None
        nbs.append(d)
    desdobros = CTRIBNAC.get(chave) or item['desdobros_anexo'] or [
        {'codigo': chave.replace('.', '') + '01', 'descricao': item['descricao']}]
    inss = INSS.get(chave)
    return {
        'item': chave,
        'descricao': item['descricao'],
        'ctribnac': desdobros,
        'nbs': nbs,
        'simples': _resumo_simples(item),
        'cnaes': [CNAES[c] for c in CNAES_POR_ITEM.get(chave, [])],
        'iss': {
            'local_prestacao': chave in ISS_LOCAL_PRESTACAO,
            'retencao_tomador': chave in ISS_RETENCAO_TOMADOR,
        },
        'inss': ({
            'aliquota': inss['aliquota'],
            'servico_rfb': inss['servico_rfb'],
            'codigo_reinf': inss['codigo_reinf'],
            'base_legal': inss['base_legal'],
            'nota': inss['nota'],
        } if inss and '11' in (inss.get('aliquota') or '') else None),
        'versao_anexo_viii': VERSAO_ANEXO_VIII,
    }


def normaliza_item(txt):
    s = re.sub(r'[^\d]', '', str(txt or ''))
    if len(s) >= 4:
        return f'{s[:2]}.{s[2:4]}'
    if len(s) == 3:
        return f'0{s[0]}.{s[1:]}'
    return str(txt or '').strip()


# Profissões e termos do dia a dia -> palavras da lista de serviços
SINONIMOS = {
    'advogado': 'advocacia', 'advogada': 'advocacia', 'juridico': 'advocacia',
    'medico': 'medicina', 'medica': 'medicina', 'clinica': 'clinicas',
    'dentista': 'odontologia', 'odontologo': 'odontologia', 'odontologica': 'odontologia',
    'contador': 'contabilidade', 'contadora': 'contabilidade', 'contabil': 'contabilidade',
    'psicologo': 'psicologia', 'psicologa': 'psicologia', 'psicanalista': 'psicanalise',
    'fisioterapeuta': 'fisioterapia', 'fonoaudiologo': 'fonoaudiologia', 'fonoaudiologa': 'fonoaudiologia',
    'nutricionista': 'nutricao', 'enfermeiro': 'enfermagem', 'enfermeira': 'enfermagem',
    'veterinario': 'veterinaria', 'veterinaria': 'veterinaria', 'farmaceutico': 'farmaceuticos',
    'engenheiro': 'engenharia', 'engenheira': 'engenharia', 'arquiteto': 'arquitetura', 'arquiteta': 'arquitetura',
    'programador': 'programacao', 'desenvolvedor': 'desenvolvimento', 'software': 'programas',
    'site': 'paginas eletronicas', 'sites': 'paginas eletronicas',
    'cabeleireiro': 'cabeleireiros', 'cabeleireira': 'cabeleireiros', 'barbeiro': 'barbearia',
    'manicure': 'manicuros', 'pedicure': 'pedicuros', 'esteticista': 'esteticistas', 'estetica': 'esteticistas',
    'academia': 'ginastica', 'personal': 'ginastica', 'educador fisico': 'ginastica',
    'professor': 'instrucao', 'professora': 'instrucao', 'curso': 'instrucao', 'cursos': 'instrucao',
    'escola': 'ensino', 'faculdade': 'ensino regular superior',
    'corretor': 'corretagem', 'corretora': 'corretagem', 'imobiliaria': 'imoveis',
    'representante': 'representacao', 'publicitario': 'publicidade', 'marketing': 'publicidade',
    'trafego': 'publicidade', 'anuncios': 'publicidade', 'designer': 'comunicacao visual',
    'fotografo': 'fotografia', 'fotografa': 'fotografia', 'filmagem': 'cinematografia',
    'consultor': 'consultoria', 'consultora': 'consultoria', 'palestrante': 'palestras',
    'despachante': 'despachantes', 'detetive': 'detetives', 'tradutor': 'traducao',
    'pedreiro': 'construcao civil', 'reforma': 'reforma de edificios', 'obra': 'obras', 'obras': 'obras',
    'mecanico': 'conserto', 'oficina': 'conserto', 'lavajato': 'lavagem', 'lava': 'lavagem',
    'eventos': 'eventos', 'buffet': 'bufe', 'dj': 'execucao de musica', 'musico': 'execucao de musica',
    'motorista': 'transporte', 'uber': 'transporte', 'frete': 'transporte', 'motoboy': 'entrega',
    'hotel': 'hospedagem', 'pousada': 'hospedagem', 'aluguel': 'locacao', 'locacao': 'locacao',
    'limpeza': 'limpeza', 'faxina': 'limpeza', 'jardineiro': 'jardinagem', 'dedetizacao': 'dedetizacao',
    'salao': 'cabeleireiros|saloes', 'beleza': 'cabeleireiros|esteticistas|manicuros',
    'tatuador': 'tatuagens', 'tatuagem': 'tatuagens', 'informatica': 'informatica', 'ti': 'informatica',
}


_STOP = {'de', 'da', 'do', 'das', 'dos', 'e', 'em', 'para', 'com', 'por', 'pago', 'paga', 'digital',
         'online', 'servico', 'servicos', 'empresa', 'atividade', 'meu', 'minha', 'cliente'}


def _expande(termos):
    """Cada termo vira um grupo: (alternativas fortes, raiz fraca)."""
    grupos = []
    for t in termos:
        if t in _STOP:
            continue
        fortes = {t}
        if t in SINONIMOS:
            fortes.update(SINONIMOS[t].split('|'))
        raiz = t[:6] if len(t) >= 8 and t not in SINONIMOS else None
        grupos.append((fortes, raiz))
    return grupos


def _hit(grupo, texto):
    fortes, raiz = grupo
    if any(a in texto for a in fortes):
        return 2
    if raiz and raiz in texto:
        return 1
    return 0


def _casa(grupos, texto):
    return bool(grupos) and all(_hit(g, texto) for g in grupos)


def _resumo_item(chave):
    it = ITENS[chave]
    return {
        'item': chave,
        'descricao': it['descricao'],
        'ctribnac': [d['codigo'] for d in CTRIBNAC.get(chave, [])],
        'simples': _resumo_simples(it),
        'qtd_nbs': len(it['nbs']),
    }


def buscar(q, limite=40):
    """Busca por CNAE, item da LC 116, cTribNac, NBS ou texto livre."""
    q = (q or '').strip()
    if not q:
        return {'tipo': 'vazio', 'itens': [], 'cnaes': []}
    digitos = re.sub(r'\D', '', q)

    # CNAE (7 dígitos)
    if len(digitos) == 7 and not q.startswith('1.'):
        cnae_fmt = f'{digitos[:4]}-{digitos[4]}/{digitos[5:]}'
        c = CNAES.get(cnae_fmt)
        if c:
            return {'tipo': 'cnae', 'cnae': c,
                    'itens': [_resumo_item(i) for i in c['itens'] if i in ITENS], 'cnaes': []}
        return {'tipo': 'cnae', 'cnae': {'cnae': cnae_fmt, 'descricao': '', 'itens': []}, 'itens': [], 'cnaes': [],
                'aviso': 'CNAE sem correspondência com a lista de serviços da LC 116. '
                         'Provavelmente é atividade de comércio ou indústria (sujeita ao ICMS), ou não consta da tabela de referência.'}

    # NBS (1.xxxx.xx.xx)
    if re.fullmatch(r'1\.?\d{4}\.?\d{2}\.?\d{2}', q):
        alvo = f'{digitos[0]}.{digitos[1:5]}.{digitos[5:7]}.{digitos[7:9]}'
        itens = [k for k, it in ITENS.items() if any(n['nbs'] == alvo for n in it['nbs'])]
        return {'tipo': 'nbs', 'nbs': alvo, 'itens': [_resumo_item(i) for i in itens], 'cnaes': []}

    # cTribNac (6 dígitos) ou item (4 dígitos / 01.01)
    if len(digitos) in (3, 4, 6) and len(digitos) == len(re.sub(r'[\s.\-/]', '', q)):
        chave = normaliza_item(digitos[:4] if len(digitos) != 3 else digitos)
        if chave in ITENS:
            return {'tipo': 'item', 'itens': [_resumo_item(chave)], 'cnaes': []}

    # Texto livre: todos os termos precisam aparecer
    termos = [t for t in _norm(q).split() if len(t) > 1]
    grupos = _expande(termos)
    achados, parciais = [], []
    for k, txt in _INDICE.items():
        hits = [_hit(g, txt) for g in grupos]
        if not any(hits):
            continue
        desc = _norm(ITENS[k]['descricao'] + ' ' + ' '.join(d['descricao'] for d in CTRIBNAC.get(k, [])))
        peso = sum(4 * _hit(g, desc) for g in grupos) + sum(hits)
        (achados if all(hits) else parciais).append((-peso, k))
    achados.sort()
    if not achados and len(grupos) > 1:
        parciais.sort()
        achados = parciais
    cnaes = [c for c in CNAES.values() if _casa(grupos, _norm(c['descricao']))][:15]
    return {'tipo': 'texto', 'itens': [_resumo_item(k) for _, k in achados[:limite]], 'cnaes': cnaes}


def calcular_fator_r(folha12, rbt12, rpa_mes=None):
    """Fator R = folha de salários dos 12 meses (com pró-labore, CPP e FGTS) / RBT12.
    Retorna o anexo resultante e as alíquotas efetivas dos anexos III e V."""
    folha12 = float(folha12 or 0)
    rbt12 = float(rbt12 or 0)
    if rbt12 <= 0:
        return {'erro': 'Informe a receita bruta dos últimos 12 meses (RBT12).'}
    fator = folha12 / rbt12
    anexo = 'III' if fator >= 0.28 else 'V'

    def efetiva(tab):
        for limite, aliq, pd in TABELAS_SN[tab]:
            if rbt12 <= limite:
                return {'faixa_limite': limite, 'aliquota_nominal': aliq, 'parcela_deduzir': pd,
                        'aliquota_efetiva': max((rbt12 * aliq - pd) / rbt12, 0)}
        return {'erro': 'RBT12 acima do limite do Simples Nacional (R$ 4,8 milhões).'}

    res = {
        'fator_r': fator,
        'anexo': anexo,
        'anexo_iii': efetiva('III'),
        'anexo_v': efetiva('V'),
        'folha_minima_para_iii': rbt12 * 0.28,
        'falta_folha_12m': max(rbt12 * 0.28 - folha12, 0),
    }
    if rpa_mes:
        rpa = float(rpa_mes)
        sel = res['anexo_iii'] if anexo == 'III' else res['anexo_v']
        if 'aliquota_efetiva' in sel:
            res['das_estimado'] = rpa * sel['aliquota_efetiva']
            outro = res['anexo_v'] if anexo == 'III' else res['anexo_iii']
            if 'aliquota_efetiva' in outro:
                res['das_outro_anexo'] = rpa * outro['aliquota_efetiva']
    return res


def estatisticas():
    return {
        'itens': len(ITENS),
        'nbs': sum(len(i['nbs']) for i in ITENS.values()),
        'cnaes': len(CNAES),
        'versao_anexo_viii': VERSAO_ANEXO_VIII,
    }
