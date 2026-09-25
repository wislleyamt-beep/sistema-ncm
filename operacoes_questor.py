"""
Sugestão de Operação do Questor (grupo 27 - SIMPLES Nacional) por NCM.

A operação do Questor combina duas dimensões que o PGDAS-D segrega:
  1. PIS/COFINS: tributação monofásica (LC 123/2006, art. 18, §4º-A, I) ou normal;
  2. ICMS no Piauí: tributação normal, substituição tributária, isenção ou imunidade.

A sugestão vem de duas fontes, nesta ordem de prioridade:
  a) Cadastro manual (tabela editável NCM -> operação), feito pelo usuário na
     aba "Operações Questor". Aceita NCM completo (8 dígitos) ou prefixo
     (4 a 7 dígitos); vale o cadastro mais específico (prefixo mais longo).
  b) Regra automática: situação de ICMS-PI por NCM x flag de monofásico.

Persistência do cadastro manual:
  - Se a variável de ambiente DATABASE_URL existir (Postgres, ex.: Neon ou
    Supabase), grava no banco — sobrevive a reinícios/deploys do Render.
  - Caso contrário, grava em JSON local (DATA_DIR/operacoes_ncm.json). No
    plano gratuito do Render esse arquivo é APAGADO a cada reinício/deploy.
"""

import json
import os
import re
import threading
from datetime import datetime

# ---------------------------------------------------------------------------
# Tabela de operações do Questor — grupo 27 (SIMPLES Nacional)
# Fonte: tela "Consulta Operações Diversas" do Questor Cloud.
# "sugerivel": pode ser escolhida pela regra automática. As demais só entram
# por cadastro manual (ex.: reduções de alíquota, exterior, legado até 2017).
# ---------------------------------------------------------------------------
OPERACOES_QUESTOR = [
    {"codigo": "2601",  "operacao": "1.01",    "descricao": "Revenda de Mercadoria Sem Sub. Tributária", "sugerivel": True},
    {"codigo": "2603",  "operacao": "1.02.01", "descricao": "Revenda de Mercadoria Com Sub. Tributária - ICMS", "sugerivel": True},
    {"codigo": "2604",  "operacao": "1.02.02", "descricao": "Revenda de Mercadoria Com Sub. Tributária - PIS e ICMS", "sugerivel": False},
    {"codigo": "2605",  "operacao": "1.02.03", "descricao": "Revenda de Mercadoria Com Sub. Tributária - COFINS e ICMS", "sugerivel": False},
    {"codigo": "2606",  "operacao": "1.02.04", "descricao": "Revenda de Mercadoria Com Sub. Tributária - PIS, COFINS e ICMS", "sugerivel": False},
    {"codigo": "2607",  "operacao": "1.02.05", "descricao": "Revenda de Mercadoria Com Sub. Tributária - PIS", "sugerivel": False},
    {"codigo": "2608",  "operacao": "1.02.06", "descricao": "Revenda de Mercadoria Com Sub. Tributária - COFINS", "sugerivel": False},
    {"codigo": "2609",  "operacao": "1.02.07", "descricao": "Revenda de Mercadoria Com Sub. Tributária - PIS e COFINS", "sugerivel": False},
    {"codigo": "70075", "operacao": "1.02.11", "descricao": "Revenda de Mercadoria Com Sub. Tributária de ICMS e Tributação Monofásica PIS e COFINS", "sugerivel": True},
    {"codigo": "2610",  "operacao": "1.03",    "descricao": "Revenda de mercadorias para o Exterior", "sugerivel": False},
    {"codigo": "2616",  "operacao": "1.04.01", "descricao": "Revenda de Mercadoria Com Isenção - ICMS", "sugerivel": True},
    {"codigo": "70070", "operacao": "1.04.02", "descricao": "Revenda de Mercadoria Com Imunidade - ICMS", "sugerivel": True},
    {"codigo": "2618",  "operacao": "1.05.01", "descricao": "Revenda de Mercadoria Tributação Monofásica PIS e COFINS", "sugerivel": True},
    {"codigo": "70076", "operacao": "1.05.02", "descricao": "Revenda de Mercadoria com Tributação Monofásica PIS e COFINS e Isenção de ICMS.", "sugerivel": True},
    {"codigo": "2755",  "operacao": "1.98.01", "descricao": "Revenda de Mercadoria Sem Sub. Tributária ICMS (Redução 40% na aliq.)", "sugerivel": False},
    {"codigo": "2751",  "operacao": "1.99.01", "descricao": "Revenda de Mercadoria Sem Sub. Tributária ICMS (Redução 60% na aliq.)", "sugerivel": False},
]
OPERACOES_IDX = {o["codigo"]: o for o in OPERACOES_QUESTOR}

# Matriz (monofásico?, situação ICMS) -> código interno da operação
_MATRIZ = {
    (False, "normal"): "2601",
    (False, "st"):     "2603",
    (False, "isento"): "2616",
    (False, "imune"):  "70070",
    (True,  "normal"): "2618",
    (True,  "st"):     "70075",
    (True,  "isento"): "70076",
    # (True, "imune") não existe no Questor — tratado como aviso.
}

SITUACOES_ICMS = {
    "normal": "ICMS tributado normalmente",
    "st":     "ICMS com substituição tributária",
    "isento": "ICMS isento",
    "imune":  "ICMS imune",
}

# ---------------------------------------------------------------------------
# Regras automáticas de situação do ICMS no Piauí (mais específica primeiro).
# Refinam o indicador por capítulo de get_icms_piaui() nos pontos em que a
# classificação muda dentro do capítulo.
# ---------------------------------------------------------------------------
_AVISO_ISENCAO_MEDICAMENTO = (
    "Alguns medicamentos das posições 30.03/30.04 são ISENTOS de ICMS pelo "
    "princípio ativo, não pelo NCM (ex.: oncológicos — Convênio ICMS 162/94; "
    "antirretrovirais/AIDS — Convênio ICMS 10/02; lista do Convênio ICMS "
    "140/01). Se o produto estiver numa dessas listas, cadastre-o como "
    "1.05.02 (70076) na aba Operações Questor."
)

_ICMS_PI_REGRAS = [
    # Imunidade constitucional — livros, jornais e periódicos (CF, art. 150, VI, "d")
    {"prefixos": ("4901", "4902", "4903"), "situacao": "imune",
     "base_legal": "CF/88, art. 150, VI, \"d\" — livros, jornais e periódicos",
     "confianca": "alta", "aviso": None},
    # Medicamentos — ST de produtos farmacêuticos no RICMS-PI
    {"prefixos": ("3003", "3004"), "situacao": "st",
     "base_legal": "RICMS-PI (Decreto 21.866/2023) — ST de produtos farmacêuticos",
     "confianca": "media", "aviso": _AVISO_ISENCAO_MEDICAMENTO},
]


def _ncm_digits(ncm):
    return re.sub(r"\D", "", str(ncm or ""))


def situacao_icms_pi(ncm_clean, icms_pi):
    """
    Retorna dict {situacao, base_legal, confianca, aviso}.
    icms_pi: resultado de get_icms_piaui(ncm_clean) (indicador por capítulo).
    """
    for regra in _ICMS_PI_REGRAS:
        if any(ncm_clean.startswith(p) for p in regra["prefixos"]):
            return {
                "situacao": regra["situacao"],
                "base_legal": regra["base_legal"],
                "confianca": regra["confianca"],
                "aviso": regra["aviso"],
            }

    if icms_pi.get("tem_st"):
        return {
            "situacao": "st",
            "base_legal": icms_pi.get("obs_st") or "RICMS-PI",
            "confianca": "baixa",
            "aviso": ("A ST foi indicada pelo capítulo do NCM. Nem todo item do "
                      "capítulo está no anexo de ST do RICMS-PI — confira o CEST "
                      "e a descrição do produto."),
        }

    aviso = None
    if ncm_clean[:2] in ("07", "08"):
        aviso = ("Hortifrutigranjeiros in natura costumam ter isenção de ICMS "
                 "(Convênio ICMS 44/75). Se for o caso, cadastre como 1.04.01 (2616).")
    return {
        "situacao": "normal",
        "base_legal": "Regra geral — Lei 8.558/2024",
        "confianca": "media",
        "aviso": aviso,
    }


# ---------------------------------------------------------------------------
# Cadastro manual (overrides)
# ---------------------------------------------------------------------------
class _Store:
    """Armazena overrides NCM(prefixo) -> {codigo, observacao, atualizado_em}."""

    def __init__(self):
        self._lock = threading.Lock()
        self.db_url = os.environ.get("DATABASE_URL", "").strip()
        data_dir = os.environ.get("DATA_DIR") or os.path.join(os.path.dirname(__file__), "data")
        self.json_path = os.path.join(data_dir, "operacoes_ncm.json")
        self._cache = None
        if self.db_url:
            self._init_db()

    # ---- Postgres ----
    def _conn(self):
        import psycopg
        return psycopg.connect(self.db_url, autocommit=True)

    def _init_db(self):
        with self._conn() as c:
            c.execute("""
                CREATE TABLE IF NOT EXISTS operacoes_ncm (
                    ncm TEXT PRIMARY KEY,
                    codigo TEXT NOT NULL,
                    observacao TEXT,
                    atualizado_em TEXT
                )
            """)

    @property
    def persistente(self):
        return bool(self.db_url)

    def all(self):
        if self.db_url:
            with self._conn() as c:
                rows = c.execute(
                    "SELECT ncm, codigo, observacao, atualizado_em FROM operacoes_ncm"
                ).fetchall()
            return {r[0]: {"codigo": r[1], "observacao": r[2] or "", "atualizado_em": r[3] or ""}
                    for r in rows}
        with self._lock:
            if self._cache is None:
                try:
                    with open(self.json_path, encoding="utf-8") as f:
                        self._cache = json.load(f)
                except (FileNotFoundError, json.JSONDecodeError):
                    self._cache = {}
            return dict(self._cache)

    def upsert(self, ncm, codigo, observacao=""):
        agora = datetime.now().strftime("%Y-%m-%d %H:%M")
        if self.db_url:
            with self._conn() as c:
                c.execute("""
                    INSERT INTO operacoes_ncm (ncm, codigo, observacao, atualizado_em)
                    VALUES (%s, %s, %s, %s)
                    ON CONFLICT (ncm) DO UPDATE SET
                      codigo = EXCLUDED.codigo,
                      observacao = EXCLUDED.observacao,
                      atualizado_em = EXCLUDED.atualizado_em
                """, (ncm, codigo, observacao, agora))
            return
        data = self.all()
        data[ncm] = {"codigo": codigo, "observacao": observacao, "atualizado_em": agora}
        self._save_json(data)

    def delete(self, ncm):
        if self.db_url:
            with self._conn() as c:
                c.execute("DELETE FROM operacoes_ncm WHERE ncm = %s", (ncm,))
            return
        data = self.all()
        data.pop(ncm, None)
        self._save_json(data)

    def _save_json(self, data):
        with self._lock:
            os.makedirs(os.path.dirname(self.json_path), exist_ok=True)
            tmp = self.json_path + ".tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2, sort_keys=True)
            os.replace(tmp, self.json_path)
            self._cache = data


_store = None


def store():
    global _store
    if _store is None:
        _store = _Store()
    return _store


def validar_cadastro(ncm, codigo):
    """Normaliza e valida; retorna (ncm, codigo, erro)."""
    ncm = _ncm_digits(ncm)
    codigo = _ncm_digits(codigo) or str(codigo or "").strip()
    # aceita também o código da operação (ex.: "1.02.11")
    if codigo not in OPERACOES_IDX:
        por_op = {o["operacao"].replace(".", ""): o["codigo"] for o in OPERACOES_QUESTOR}
        codigo = por_op.get(codigo, codigo)
    if not (4 <= len(ncm) <= 8):
        return None, None, "Informe o NCM com 4 a 8 dígitos."
    if codigo not in OPERACOES_IDX:
        return None, None, f"Operação '{codigo}' não existe na tabela do Questor (grupo 27)."
    return ncm, codigo, None


def _override_para(ncm_clean, overrides):
    for tam in range(8, 3, -1):
        chave = ncm_clean[:tam]
        if chave in overrides:
            return chave, overrides[chave]
    return None, None


def _op_dict(codigo):
    op = OPERACOES_IDX.get(codigo)
    if not op:
        return None
    return {
        "codigo": op["codigo"],
        "operacao": op["operacao"],
        "descricao": op["descricao"],
        "rotulo": f'{op["operacao"]} - {op["descricao"]}',
    }


def sugerir_operacao(ncm_clean, is_monofasico, icms_pi, aliquota_zero=None, overrides=None):
    """
    Monta a sugestão de operação do Questor para um NCM.
    overrides: dict opcional já carregado (evita reler o banco em lote).
    """
    if overrides is None:
        try:
            overrides = store().all()
        except Exception:
            overrides = {}

    icms = situacao_icms_pi(ncm_clean, icms_pi)
    avisos = []
    if icms.get("aviso"):
        avisos.append(icms["aviso"])
    if aliquota_zero and not is_monofasico:
        avisos.append(
            "Alíquota zero de PIS/COFINS não é segregada no PGDAS-D como "
            "monofásica — no Simples Nacional a receita é tributada normalmente."
        )

    chave, manual = _override_para(ncm_clean, overrides)
    codigo_auto = _MATRIZ.get((bool(is_monofasico), icms["situacao"]))
    if codigo_auto is None:
        avisos.append("Combinação monofásico + imunidade de ICMS não tem operação "
                      "própria no Questor — defina manualmente.")

    resultado = {
        "pis_cofins": "monofasico" if is_monofasico else "normal",
        "pis_cofins_label": ("Monofásico (LC 123/2006, art. 18, §4º-A, I)"
                             if is_monofasico else "Tributação normal no DAS"),
        "icms_situacao": icms["situacao"],
        "icms_label": SITUACOES_ICMS[icms["situacao"]],
        "icms_base_legal": icms["base_legal"],
        "confianca": icms["confianca"],
        "automatica": _op_dict(codigo_auto) if codigo_auto else None,
        "avisos": avisos,
        "origem": "automatica",
        "manual": None,
    }

    if manual and manual.get("codigo") in OPERACOES_IDX:
        resultado["origem"] = "cadastro"
        resultado["manual"] = {
            "ncm_cadastrado": chave,
            "observacao": manual.get("observacao", ""),
            "atualizado_em": manual.get("atualizado_em", ""),
        }
        resultado["operacao"] = _op_dict(manual["codigo"])
        if codigo_auto and manual["codigo"] != codigo_auto:
            resultado["divergencia"] = (
                f"O cadastro manual difere da regra automática "
                f"({resultado['automatica']['rotulo']})."
            )
    else:
        resultado["operacao"] = resultado["automatica"]

    return resultado
