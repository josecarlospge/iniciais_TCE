#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
enviador_peticao.py  --  v1.0
==============================
Módulo de envio de petição inicial de execução fiscal ao PJe (TJPI).
Baseado na lógica do sistema Dívida Ativa / PGE-PI.

Fluxo (idêntico ao sistema de referência):
  1. Usuário seleciona uma certidão ainda não protocolizada
  2. Sistema exibe os réus (responsáveis não excluídos) com endereços
  3. Usuário seleciona a comarca de destino
  4. Pergunta se deseja redirecionar ao Núcleo de Justiça 4.0
  5. Sistema solicita o arquivo PDF assinado (p7s/pdf)
  6. Confirmação final "Deseja realmente enviar?"
  7. Envio ao PJe via SOAP intercomunicação 2.2.2
  8. Número do processo criado é salvo na tabela peticoes_enviadas

Banco: certidoes_tce.db (mesmo do extrator_certidao.py)
Novas tabelas criadas automaticamente:
  - cod_comarcas      (comarca TEXT UNIQUE, cod_comarca TEXT)
  - peticoes_enviadas (log de protocolos)
"""

import base64
import hashlib
import os
import re
import sqlite3
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import requests

# ============================================================
# CONSTANTES PJe  (idênticas ao sistema de referência)
# ============================================================

PJE_URL           = "https://pje.tjpi.jus.br/1g/intercomunicacao?wsdl"
CLASSE_PROCESSUAL = "12154"      # Execução Fiscal
CODIGO_ASSUNTO    = "7703"     # Execução Fiscal (CNJ)
TIPO_DOC_PJE      = "8180008"   # Petição Inicial

# Polo Ativo fixo – Estado do Piauí / PGE-PI
AT_NOME   = "ESTADO DO PIAUI"
AT_CNPJ   = "06553481000149"
AT_CEP    = "64049110"
AT_LOGR   = "Av. Sen. Area Leão"
AT_NUM    = "1650"
AT_BAIRRO = "Jóquei"
AT_CIDADE = "Teresina"
AT_UF     = "PI"
AT_PAIS   = "BR"

ADV_NOME  = "Procuradoria Geral do Estado do Piauí"
ADV_CNPJ  = "06553481000491"

DB_PATH   = "certidoes_tce.db"

# ============================================================
# BANCO DE DADOS
# ============================================================

DDL_COMARCAS = """
CREATE TABLE IF NOT EXISTS cod_comarcas (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    comarca     TEXT NOT NULL UNIQUE,
    cod_comarca TEXT NOT NULL
);"""

DDL_PETICOES = """
CREATE TABLE IF NOT EXISTS peticoes_enviadas (
    id                   INTEGER PRIMARY KEY AUTOINCREMENT,
    certidao_id          INTEGER REFERENCES certidoes(id),
    numero_processo_pje  TEXT,
    comarca              TEXT,
    cod_comarca          TEXT,
    caminho_pdf_assinado TEXT,
    data_envio           TEXT DEFAULT CURRENT_TIMESTAMP,
    status               TEXT DEFAULT 'enviada',
    resposta_pje         TEXT
);"""

# Comarcas TJPI — edite/complemente via interface "Gerenciar comarcas"
COMARCAS_PADRAO = [
    ("ALTOS",                 "8743"),
    ("BARRAS",                "8748"),
    ("CAMPO MAIOR",           "8753"),
    ("CORRENTE",              "8760"),
    ("ESPERANTINA",           "8767"),
    ("FLORIANO",              "8771"),
    ("OEIRAS",                "8789"),
    ("PARNAÍBA",              "8792"),
    ("PICOS",                 "8795"),
    ("PIRIPIRI",              "8797"),
    ("SÃO RAIMUNDO NONATO",   "8810"),
    ("TERESINA",              "8815"),
    ("UNIÃO",                 "8818"),
    ("URUÇUÍ",                "8819"),
    ("NUCLEO DE JUSTICA 4.0", "9999"),
]


def abrir_banco(caminho: str = DB_PATH) -> sqlite3.Connection:
    conn = sqlite3.connect(caminho)
    conn.row_factory = sqlite3.Row
    conn.executescript(DDL_COMARCAS + DDL_PETICOES)
    conn.commit()
    cur = conn.cursor()
    cur.execute("SELECT COUNT(*) FROM cod_comarcas")
    if cur.fetchone()[0] == 0:
        conn.executemany(
            "INSERT OR IGNORE INTO cod_comarcas (comarca, cod_comarca) VALUES (?,?)",
            COMARCAS_PADRAO,
        )
        conn.commit()
    return conn


# ============================================================
# UTILITÁRIOS  (equivalentes ao sistema de referência)
# ============================================================

def converter_para_base64(caminho: str) -> str:
    with open(caminho, "rb") as f:
        return base64.b64encode(f.read()).decode("utf-8")


def hash_file(caminho: str) -> str:
    h = hashlib.sha256()
    with open(caminho, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def valor_float(valor_str: str) -> float:
    """Converte 'R$ 167.406,32' → 167406.32"""
    s = re.sub(r"[R$\s]", "", str(valor_str or "0"))
    s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def _esc(s: str) -> str:
    return (str(s or "")
            .replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;"))


# ============================================================
# MONTAGEM DO XML SOAP
# ============================================================

def _buscar_endereco(conn: sqlite3.Connection,
                     numero_doc: str, endereco_texto: str) -> dict:
    """
    Busca endereço estruturado em enderecos_responsavel.
    Fallback: parseia responsaveis.endereco (texto livre).
    """
    if numero_doc:
        cur = conn.cursor()
        cur.execute(
            "SELECT * FROM enderecos_responsavel WHERE numero_doc=? LIMIT 1",
            (numero_doc,))
        row = cur.fetchone()
        if row and row["logradouro"]:
            cep  = (row["cep"] or "").replace("-", "").replace(" ", "")
            logr = f"{row['tipo_logradouro'] or ''} {row['logradouro'] or ''}".strip()
            return {
                "cep":         cep or "00000000",
                "logradouro":  logr,
                "numero":      row["numero"] or "S/N",
                "complemento": row["complemento"] or "",
                "bairro":      row["bairro"] or "",
                "cidade":      row["municipio"] or "",
                "uf":          row["uf"] or "PI",
                "pais":        row["pais"] or "BR",
            }

    if endereco_texto and endereco_texto.strip():
        try:
            from gestor_enderecos import parsear_endereco
            e   = parsear_endereco(endereco_texto)
            cep = (e.cep or "").replace("-", "").replace(" ", "")
            return {
                "cep":         cep or "00000000",
                "logradouro":  f"{e.tipo_logradouro} {e.logradouro}".strip(),
                "numero":      e.numero or "S/N",
                "complemento": e.complemento or "",
                "bairro":      e.bairro or "",
                "cidade":      e.municipio or "",
                "uf":          e.uf or "PI",
                "pais":        e.pais or "BR",
            }
        except Exception:
            pass

    return {"cep": "00000000", "logradouro": "ENDERECO NAO INFORMADO",
            "numero": "SN",    "complemento": "", "bairro": "",
            "cidade": "",      "uf": "PI", "pais": "BR"}


def _bloco_polo_passivo(responsaveis: list, conn: sqlite3.Connection) -> str:
    """
    Gera um <int:parte> por réu.
    CPF → tipoPessoa="fisica" / tipoDocumento="CPF"
    CNPJ → tipoPessoa="juridica" / tipoDocumento="CMF"
    """
    blocos = []
    for resp in responsaveis:
        nome      = _esc(resp["nome"].replace("&", "E"))
        num_doc   = resp["numero_doc"] or ""
        tipo_doc  = (resp["tipo_doc"] or "CPF").upper()
        doc_limpo = re.sub(r"[.\-/]", "", num_doc)

        tipo_pessoa  = "juridica" if tipo_doc == "CNPJ" else "fisica"
        tipo_doc_pje = "CMF" #if tipo_doc == "CNPJ" else "CMF"

        end = _buscar_endereco(conn, num_doc, resp.get("endereco") or "")

        blocos.append(f"""
                <int:parte assistenciaJudiciaria="1" intimacaoPendente="1" relacionamentoProcessual="1">
                    <int:pessoa nome="{nome}" numeroDocumentoPrincipal="{doc_limpo}" tipoPessoa="{tipo_pessoa}">
                        <int:outroNome>?</int:outroNome>
                        <int:documento codigoDocumento="{doc_limpo}" emissorDocumento="MF" tipoDocumento="{tipo_doc_pje}" nome="{nome}"/>
                        <int:endereco cep="{end['cep']}">
                            <int:logradouro>{_esc(end['logradouro'])}</int:logradouro>
                            <int:numero>{_esc(end['numero'])}</int:numero>
                            <int:complemento>{_esc(end['complemento'])}</int:complemento>
                            <int:bairro>{_esc(end['bairro'])}</int:bairro>
                            <int:cidade>{_esc(end['cidade'])}</int:cidade>
                            <int:estado>{_esc(end['uf'])}</int:estado>
                            <int:pais>{_esc(end['pais'])}</int:pais>
                        </int:endereco>
                    </int:pessoa>
                </int:parte>""")

    return "\n".join(blocos)


def montar_body(certidao: dict, responsaveis: list,
                conn: sqlite3.Connection, cod_comarca: str,
                valor: float, conteudo_b64: str, pdf_hash: str,
                ID_MANIFESTANTE, SENHA_MANIFEST) -> bytes:
    """Monta o envelope SOAP exatamente como no sistema de referência."""
    polo_passivo = _bloco_polo_passivo(responsaveis, conn)

    body = f"""
            <soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/" xmlns:ser="http://www.cnj.jus.br/servico-intercomunicacao-2.2.2/" xmlns:tip="http://www.cnj.jus.br/tipos-servico-intercomunicacao-2.2.2" xmlns:int="http://www.cnj.jus.br/intercomunicacao-2.2.2">
                <soapenv:Header/>
                <soapenv:Body>
                    <ser:entregarManifestacaoProcessual>
                        <!--Optional:-->
                        <tip:idManifestante>{ID_MANIFESTANTE}</tip:idManifestante>
                        <!--Optional:-->
                        <tip:senhaManifestante>{SENHA_MANIFEST}</tip:senhaManifestante>
                        <!--Optional:-->
                        <tip:dadosBasicos competencia="9" classeProcessual="{CLASSE_PROCESSUAL}" codigoLocalidade="{cod_comarca}">
                            <!--1 or more repetitions:-->
                            <!-- POLO ATIVO -->
                            <int:polo polo="AT">
                                <!--1 or more repetitions:-->
                                <int:parte assistenciaJudiciaria="1" intimacaoPendente="1" relacionamentoProcessual="1">
                                    <!--Optional:-->
                                    <int:pessoa nome="{AT_NOME}" numeroDocumentoPrincipal="{AT_CNPJ}" tipoPessoa="juridica">
                                        <!--Zero or more repetitions:-->
                                        <int:outroNome>?</int:outroNome>
                                        <!--Zero or more repetitions:-->
                                        <int:documento codigoDocumento="{AT_CNPJ}" emissorDocumento="SRFB" tipoDocumento="CMF" nome="{AT_NOME}"/>
                                        <!--Zero or more repetitions:-->
                                        <int:endereco cep="{AT_CEP}">
                                            <!--Optional:-->
                                            <int:logradouro>{AT_LOGR}</int:logradouro>
                                            <!--Optional:-->
                                            <int:numero>{AT_NUM}</int:numero>
                                            <!--Optional:-->
                                            <int:bairro>{AT_BAIRRO}</int:bairro>
                                            <!--Optional:-->
                                            <int:cidade>{AT_CIDADE}</int:cidade>
                                            <!--Optional:-->
                                            <int:estado>{AT_UF}</int:estado>
                                            <!--Optional:-->
                                            <int:pais>{AT_PAIS}</int:pais>
                                        </int:endereco>
                                    </int:pessoa>
                                    <!--Zero or more repetitions:-->
                                    <int:advogado nome="{ADV_NOME}" numeroDocumentoPrincipal="{ADV_CNPJ}" intimacao="true" tipoRepresentante="P">
                                        <!--Zero or more repetitions:-->
                                        <int:endereco cep="{AT_CEP}">
                                            <!--Optional:-->
                                            <int:logradouro>Sen. Area Leão</int:logradouro>
                                            <!--Optional:-->
                                            <int:numero>{AT_NUM}</int:numero>
                                            <!--Optional:-->
                                            <int:bairro>{AT_BAIRRO}</int:bairro>
                                            <!--Optional:-->
                                            <int:cidade>{AT_CIDADE}</int:cidade>
                                            <!--Optional:-->
                                            <int:estado>{AT_UF}</int:estado>
                                            <!--Optional:-->
                                            <int:pais>{AT_PAIS}</int:pais>
                                        </int:endereco>
                                    </int:advogado>
                                </int:parte>
                            </int:polo>
                            <!-- POLO PASSIVO -->
                            <int:polo polo="PA">
                                <!--1 or more repetitions:-->
                                {polo_passivo}
                            </int:polo>
                            <!--1 or more repetitions:-->
                            <int:assunto principal="true">
                                <!--Optional:-->
                                <int:codigoNacional>{CODIGO_ASSUNTO}</int:codigoNacional>
                            </int:assunto>
                            <!--Optional:-->
                            <int:valorCausa>{valor}</int:valorCausa>
                            <!--Zero or more repetitions:-->
                            <int:outrosnumeros>{certidao['numero_processo']}</int:outrosnumeros>
                        </tip:dadosBasicos>
                        <!-- PETIÇÃO -->
                        <!--1 or more repetitions:-->
                        <tip:documento tipoDocumento="{TIPO_DOC_PJE}" mimetype="application/pdf" nivelSigilo="0" hash="{pdf_hash}" descricao="Petição Inicial">
                            <!--Optional:-->
                            <int:conteudo>{conteudo_b64}</int:conteudo>
                        </tip:documento>
                    </ser:entregarManifestacaoProcessual>
                </soapenv:Body>
            </soapenv:Envelope>"""

    return body.encode("utf-8")


# ============================================================
# GUI
# ============================================================

CORES = {
    "bg":     "#0f1923", "panel":  "#162032", "panel2": "#1c2d42",
    "accent": "#1e9fd4", "accent2":"#0d7aad", "success":"#22c55e",
    "warning":"#f59e0b", "danger": "#ef4444", "texto":  "#e8f4f8",
    "sub":    "#8badc4", "borda":  "#2a4a6b",
}
FONTE        = ("Segoe UI", 10)
FONTE_TITULO = ("Segoe UI Semibold", 11)
FONTE_MONO   = ("Consolas", 9)
FONTE_BTN    = ("Segoe UI Semibold", 11)


class EnviadorApp(tk.Toplevel):
    """
    Janela de envio de petição inicial ao PJe.
    Use abrir_enviador(master, db_path) para instanciar a partir do extrator.
    """

    def __init__(self, master=None, db_path: str = DB_PATH):
        super().__init__(master)
        self.db_path       = db_path
        self.conn          = abrir_banco(db_path)
        self._certidao     = None   # sqlite3.Row da certidão selecionada
        self._responsaveis = []     # sqlite3.Row dos réus

        self.title("Envio de Petição Inicial — PJe TJPI")
        self.geometry("1080x700")
        self.configure(bg=CORES["bg"])
        self.resizable(True, True)

        self._estilos_ttk()
        self._build_ui()
        self._carregar_certidoes()

    # ── Estilos ────────────────────────────────────────────────

    def _estilos_ttk(self):
        st = ttk.Style()
        try:
            st.theme_use("clam")
        except Exception:
            pass
        st.configure("Treeview", background=CORES["panel2"],
                     foreground=CORES["texto"], fieldbackground=CORES["panel2"],
                     rowheight=23, font=FONTE)
        st.configure("Treeview.Heading", background=CORES["borda"],
                     foreground=CORES["sub"], font=FONTE_TITULO)
        st.map("Treeview",
               background=[("selected", CORES["accent2"])],
               foreground=[("selected", "white")])
        st.configure("TScrollbar", background=CORES["panel2"],
                     troughcolor=CORES["bg"], arrowcolor=CORES["sub"])
        st.configure("TCombobox", fieldbackground=CORES["panel2"],
                     background=CORES["panel2"], foreground=CORES["texto"])
        st.map("TCombobox",
               fieldbackground=[("readonly", CORES["panel2"])],
               foreground=[("readonly", CORES["texto"])])

    # ── Construção ─────────────────────────────────────────────

    def _build_ui(self):
        # Cabeçalho
        cab = tk.Frame(self, bg=CORES["accent2"], pady=9)
        cab.pack(fill="x")
        tk.Label(cab, text="⚖  Envio de Petição Inicial — Execução Fiscal TCE/PI",
                 bg=CORES["accent2"], fg="white",
                 font=("Segoe UI Semibold", 13)).pack(side="left", padx=18)

        # Corpo
        corpo = tk.Frame(self, bg=CORES["bg"])
        corpo.pack(fill="both", expand=True, padx=10, pady=8)

        esq = tk.Frame(corpo, bg=CORES["panel"])
        esq.pack(side="left", fill="both", expand=True, padx=(0, 6))
        self._painel_certidoes(esq)

        dir_ = tk.Frame(corpo, bg=CORES["panel"], width=440)
        dir_.pack(side="right", fill="both", expand=False, padx=(6, 0))
        dir_.pack_propagate(False)
        self._painel_direita(dir_)

        # Barra de status
        self.var_status = tk.StringVar(value="Selecione uma certidão na lista.")
        barra = tk.Frame(self, bg=CORES["panel2"], pady=4)
        barra.pack(fill="x", side="bottom")
        self.lbl_status = tk.Label(barra, textvariable=self.var_status,
                                   bg=CORES["panel2"], fg=CORES["sub"],
                                   font=FONTE, anchor="w")
        self.lbl_status.pack(side="left", padx=12)

    def _painel_certidoes(self, p):
        self._sec(p, "CERTIDÕES DISPONÍVEIS")
        self._linha(p)

        fr = tk.Frame(p, bg=CORES["panel"])
        fr.pack(fill="x", padx=8, pady=(0, 4))
        tk.Label(fr, text="🔍", bg=CORES["panel"], fg=CORES["sub"]).pack(side="left")
        self.var_busca = tk.StringVar()
        self.var_busca.trace_add("write", lambda *a: self._filtrar())
        tk.Entry(fr, textvariable=self.var_busca, bg=CORES["panel2"],
                 fg=CORES["texto"], insertbackground=CORES["texto"],
                 relief="flat", font=FONTE, bd=4
                 ).pack(side="left", fill="x", expand=True, padx=4)

        cols = ("processo", "valor", "data", "status")
        self.tree_cert = ttk.Treeview(p, columns=cols, show="headings", height=20)
        for cid, lbl, w in [("processo","Processo TC",180),("valor","Valor Atual",120),
                             ("data","Data ref.",100),("status","Status PJe",135)]:
            self.tree_cert.heading(cid, text=lbl)
            self.tree_cert.column(cid, width=w, minwidth=60, anchor="w")
        self.tree_cert.tag_configure("pendente",
            background=CORES["panel2"], foreground=CORES["texto"])
        self.tree_cert.tag_configure("protocolada",
            background="#0a2210", foreground=CORES["success"])

        sb = ttk.Scrollbar(p, orient="vertical", command=self.tree_cert.yview)
        self.tree_cert.configure(yscrollcommand=sb.set)
        self.tree_cert.pack(side="left", fill="both", expand=True, padx=(8,0), pady=4)
        sb.pack(side="left", fill="y", pady=4)
        self.tree_cert.bind("<<TreeviewSelect>>", self._ao_selecionar)

        tk.Button(p, text="↺  Recarregar", command=self._carregar_certidoes,
                  bg=CORES["panel2"], fg=CORES["sub"], relief="flat",
                  font=FONTE, cursor="hand2"
                  ).pack(pady=6, padx=8, anchor="e")

    def _painel_direita(self, p):
        self._sec(p, "CERTIDÃO SELECIONADA")
        self._linha(p)

        inf = tk.Frame(p, bg=CORES["panel"])
        inf.pack(fill="x", padx=10)
        self.var_proc    = tk.StringVar(value="—")
        self.var_valor   = tk.StringVar(value="—")
        self.var_acordao = tk.StringVar(value="—")
        for lbl, var in [("Processo:", self.var_proc),
                         ("Valor:",    self.var_valor),
                         ("Acórdão:", self.var_acordao)]:
            r = tk.Frame(inf, bg=CORES["panel"])
            r.pack(fill="x", pady=1)
            tk.Label(r, text=lbl, width=10, anchor="w",
                     bg=CORES["panel"], fg=CORES["sub"], font=FONTE).pack(side="left")
            tk.Label(r, textvariable=var, anchor="w",
                     bg=CORES["panel"], fg=CORES["texto"], font=FONTE_MONO,
                     wraplength=310, justify="left").pack(side="left")

        self._linha(p)
        self._sec(p, "RÉUS REMANESCENTES")

        frr = tk.Frame(p, bg=CORES["panel"])
        frr.pack(fill="x", padx=8)
        cols_r = ("nome","doc","tipo","endereco")
        self.tree_reus = ttk.Treeview(frr, columns=cols_r, show="headings", height=7)
        for cid, lbl, w in [("nome","Nome",155),("doc","CPF/CNPJ",115),
                             ("tipo","Tipo",48),("endereco","Endereço",125)]:
            self.tree_reus.heading(cid, text=lbl)
            self.tree_reus.column(cid, width=w, minwidth=40, anchor="w")
        self.tree_reus.tag_configure("sem_end",
            background="#2a1a00", foreground=CORES["warning"])
        self.tree_reus.tag_configure("com_end",
            background=CORES["panel2"], foreground=CORES["texto"])
        sb_r = ttk.Scrollbar(frr, orient="vertical", command=self.tree_reus.yview)
        self.tree_reus.configure(yscrollcommand=sb_r.set)
        self.tree_reus.pack(side="left", fill="both", expand=True)
        sb_r.pack(side="left", fill="y")

        self._linha(p)
        self._sec(p, "COMARCA E ENVIO")

        env = tk.Frame(p, bg=CORES["panel"])
        env.pack(fill="x", padx=10, pady=4)
        env.columnconfigure(1, weight=1)

        tk.Label(env, text="Comarca:", bg=CORES["panel"],
                 fg=CORES["sub"], font=FONTE
                 ).grid(row=0, column=0, sticky="w", pady=3)
        self.var_comarca = tk.StringVar()
        self.combo_comarca = ttk.Combobox(env, textvariable=self.var_comarca,
                                          state="readonly", width=28, font=FONTE)
        self.combo_comarca.grid(row=0, column=1, sticky="ew", pady=3, padx=(6,0))
        self._carregar_comarcas()

        tk.Button(env, text="⚙ Gerenciar comarcas",
                  command=self._abrir_gerenciador,
                  bg=CORES["panel2"], fg=CORES["sub"],
                  relief="flat", font=("Segoe UI", 9), cursor="hand2"
                  ).grid(row=1, column=0, columnspan=2, sticky="e", pady=(0,4))

        self.btn_enviar = tk.Button(
            env, text="⚡  ENVIAR PETIÇÃO AO PJe",
            command=self.enviar_peticao,
            bg=CORES["accent2"], fg="white",
            activebackground=CORES["accent"],
            relief="flat", font=FONTE_BTN, cursor="hand2", pady=9)
        self.btn_enviar.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(6,0))

        self._linha(p)
        self._sec(p, "HISTÓRICO DE PROTOCOLOS")
        self.txt_log = tk.Text(p, height=5, bg=CORES["panel2"], fg=CORES["sub"],
                               font=FONTE_MONO, relief="flat", state="disabled")
        self.txt_log.pack(fill="x", padx=8, pady=(0,8))

    # ── Helpers ────────────────────────────────────────────────

    def _sec(self, p, txt):
        tk.Label(p, text=txt, bg=CORES["panel"], fg=CORES["accent"],
                 font=FONTE_TITULO).pack(anchor="w", padx=10, pady=(10,2))

    def _linha(self, p):
        tk.Frame(p, bg=CORES["borda"], height=1).pack(fill="x", padx=8, pady=3)

    def _set_status(self, msg: str, cor: str = None):
        self.var_status.set(msg)
        if cor:
            self.lbl_status.config(fg=cor)
        self.update_idletasks()

    # ── Dados ──────────────────────────────────────────────────

    def _carregar_comarcas(self):
        cur = self.conn.cursor()
        cur.execute("SELECT comarca FROM cod_comarcas ORDER BY comarca")
        comarcas = [r["comarca"] for r in cur.fetchall()]
        self.combo_comarca["values"] = comarcas
        if comarcas:
            self.combo_comarca.set(comarcas[0])

    def _carregar_certidoes(self):
        cur = self.conn.cursor()
        cur.execute("""
            SELECT c.id, c.numero_processo, c.valor_atualizado,
                   c.data_atualizacao, c.acordao_origem,
                   (SELECT pe.numero_processo_pje
                    FROM peticoes_enviadas pe
                    WHERE pe.certidao_id = c.id AND pe.status = 'enviada'
                    LIMIT 1) AS numero_pje
            FROM certidoes c
            ORDER BY c.data_insercao DESC
        """)
        self._rows_cert = cur.fetchall()
        self._popular_tree_cert(self._rows_cert)

    def _popular_tree_cert(self, rows):
        self.tree_cert.delete(*self.tree_cert.get_children())
        for row in rows:
            pje    = row["numero_pje"]
            status = f"✅ {pje}" if pje else "⏳ Pendente"
            tag    = "protocolada" if pje else "pendente"
            self.tree_cert.insert("", "end", iid=str(row["id"]),
                values=(row["numero_processo"], row["valor_atualizado"] or "—",
                        row["data_atualizacao"] or "—", status),
                tags=(tag,))

    def _filtrar(self):
        termo = self.var_busca.get().lower().strip()
        filtrado = (self._rows_cert if not termo else
                    [r for r in self._rows_cert
                     if termo in (r["numero_processo"] or "").lower()
                     or termo in (r["valor_atualizado"] or "").lower()])
        self._popular_tree_cert(filtrado)

    def _ao_selecionar(self, event=None):
        sel = self.tree_cert.selection()
        if not sel:
            return
        cert_id = int(sel[0])
        cur = self.conn.cursor()
        cur.execute("SELECT * FROM certidoes WHERE id=?", (cert_id,))
        self._certidao = cur.fetchone()
        if not self._certidao:
            return

        self.var_proc.set(self._certidao["numero_processo"])
        self.var_valor.set(self._certidao["valor_atualizado"] or "—")
        self.var_acordao.set(self._certidao["acordao_origem"] or "—")

        # Réus com endereço
        cur.execute("""
            SELECT r.id, r.nome, r.tipo_doc, r.numero_doc, r.endereco,
                   e.tipo_logradouro, e.logradouro, e.municipio, e.uf
            FROM responsaveis r
            LEFT JOIN enderecos_responsavel e ON e.id = (
                SELECT id FROM enderecos_responsavel
                WHERE numero_doc = r.numero_doc
                LIMIT 1
            )
            WHERE r.certidao_id = ? AND r.excluido = 0
        """, (cert_id,))
        self._responsaveis = cur.fetchall()

        self.tree_reus.delete(*self.tree_reus.get_children())
        for r in self._responsaveis:
            if r["logradouro"]:
                end_curto = (f"{r['tipo_logradouro'] or ''} {r['logradouro']}".strip()
                             + (f" — {r['municipio']}/{r['uf'] or 'PI'}"
                                if r["municipio"] else ""))
            elif r["endereco"]:
                end_curto = r["endereco"][:50]
            else:
                end_curto = None
            tag = "com_end" if end_curto else "sem_end"
            self.tree_reus.insert("", "end",
                values=(r["nome"][:32], r["numero_doc"] or "—",
                        r["tipo_doc"] or "—",
                        end_curto or "⚠ SEM ENDEREÇO"),
                tags=(tag,))

        # Histórico
        cur.execute("""
            SELECT numero_processo_pje, comarca, data_envio, status
            FROM peticoes_enviadas WHERE certidao_id=? ORDER BY data_envio DESC
        """, (cert_id,))
        logs = cur.fetchall()
        self.txt_log.config(state="normal")
        self.txt_log.delete("1.0", "end")
        if logs:
            for lg in logs:
                self.txt_log.insert("end",
                    f"[{(lg['data_envio'] or '')[:16]}] "
                    f"{(lg['status'] or '').upper()} → "
                    f"{lg['numero_processo_pje'] or 'erro'} | {lg['comarca'] or ''}\n")
        else:
            self.txt_log.insert("end", "(nenhum protocolo registrado para esta certidão)")
        self.txt_log.config(state="disabled")

        self._set_status(
            f"Certidão: {self._certidao['numero_processo']} — "
            f"{len(self._responsaveis)} réu(s) remanescente(s).")

    # ──────────────────────────────────────────────────────────
    # ENVIO  (lógica idêntica ao sistema de referência)
    # ──────────────────────────────────────────────────────────

    def enviar_peticao(self):
        # 1. Validações básicas
        if not self._certidao:
            messagebox.showinfo("Info", "Selecione uma certidão na lista antes de enviar.")
            return
        if not self._responsaveis:
            messagebox.showinfo("Info", "Nenhum réu remanescente encontrado para esta certidão.")
            return
        comarca = self.var_comarca.get().strip()
        if not comarca:
            messagebox.showinfo("Info", "Selecione uma Comarca antes de Gerar a Inicial.")
            return

        cert_id = self._certidao["id"]
        cur = self.conn.cursor()

        # 2. Verifica se já protocolada  (equivalente ao check na tabela "iniciais")
        cur.execute("""
            SELECT numero_processo_pje FROM peticoes_enviadas
            WHERE certidao_id=? AND status='enviada' LIMIT 1
        """, (cert_id,))
        row_exist = cur.fetchone()
        if row_exist:
            ok = messagebox.askyesno(
                "Já Protocolada",
                f"Esta certidão já foi objeto de protocolo em execução fiscal!\n"
                f"Processo PJe: {row_exist['numero_processo_pje']}\n\n"
                f"Deseja reenviar mesmo assim?"
            )
            if not ok:
                return

        # 3. Pergunta Núcleo 4.0  (idêntico ao ref)
        resposta_nucleo = messagebox.askyesno(
            "Encaminhar Petição",
            "Deseja encaminhar a petição ao NÚCLEO DE JUSTICA 4.0?"
        )
        comarca_envio = "NUCLEO DE JUSTICA 4.0" if resposta_nucleo else comarca

        # 4. Busca código da comarca  (idêntico ao ref)
        cur.execute(
            "SELECT cod_comarca FROM cod_comarcas WHERE comarca=? LIMIT 1",
            (comarca_envio,)
        )
        row_cc = cur.fetchone()
        if not row_cc:
            messagebox.showerror(
                "Erro",
                f'Não foi encontrado nenhum código para a comarca "{comarca_envio}"'
            )
            return
        cod_comarca = row_cc["cod_comarca"]

        # 5. Solicita arquivo PDF assinado  (idêntico ao ref: filedialog no fluxo)
        nomes_reus = ", ".join(r["nome"] for r in self._responsaveis)
        file_path1 = filedialog.askopenfilename(
            title=f"Selecione o arquivo PDF ASSINADO com a Inicial de {nomes_reus[:60]}"
        )
        if not file_path1:
            self._set_status("Envio cancelado: nenhum arquivo selecionado.")
            return

        # 6. Converte PDF  (idêntico ao ref)
        self._set_status("Aguarde. Estou realizando a conexão com a API do PJe.")
        try:
            conteudo_base64 = converter_para_base64(file_path1)
            pdf_hash        = hash_file(file_path1)
        except Exception as exc:
            messagebox.showerror("Erro", f"Não foi possível ler o arquivo PDF:\n{exc}")
            return

        valor = round(valor_float(self._certidao["valor_atualizado"]), 2)

        # 7. Monta XML SOAP
        resp_dicts = [dict(r) for r in self._responsaveis]
        
        
        # SOLICITAR SENHA PARA ENVIO
        import tkinter as tk
        from tkinter import simpledialog

        def solicitar_credenciais(master):
            dialog = tk.Toplevel(master)
            dialog.title("Credenciais do Manifestante")
            dialog.geometry("350x180")
            dialog.resizable(False, False)
            dialog.grab_set()  # modal

            tk.Label(dialog, text="CPF do Manifestante:").pack(pady=(10, 2))
            entry_cpf = tk.Entry(dialog, width=30)
            entry_cpf.pack()

            tk.Label(dialog, text="Senha do Manifestante:").pack(pady=(10, 2))
            entry_senha = tk.Entry(dialog, width=30, show="*")
            entry_senha.pack()

            resultado = {}

            def confirmar():
                resultado["cpf"] = entry_cpf.get().strip()
                resultado["senha"] = entry_senha.get().strip()
                dialog.destroy()

            tk.Button(dialog, text="Confirmar",
                    bg="#1A3A5C", fg="white",
                    command=confirmar).pack(pady=15)

            master.wait_window(dialog)

            return resultado.get("cpf"), resultado.get("senha")
        
        
        cpf, senha = solicitar_credenciais(self)

        if not cpf or not senha:
            messagebox.showerror("Erro", "CPF e senha são obrigatórios.")
            return
        body = montar_body(dict(self._certidao), resp_dicts, self.conn,
                           cod_comarca, valor, conteudo_base64, pdf_hash, cpf, senha)

        
        # 8. Confirmação final "Deseja, realmente, enviar?"  (idêntico ao ref)
        nomes_fmt = "\n".join(f"  • {r['nome']}" for r in self._responsaveis)
        resposta = messagebox.askyesno(
            "Tudo Pronto",
            f"Deseja, realmente, enviar a Inicial para protocolo?\n\n"
            f"Processo TCE:   {self._certidao['numero_processo']}\n"
            f"Comarca:        {comarca_envio}  (cód. {cod_comarca})\n"
            f"Valor da causa: R$ {valor:,.2f}\n"
            f"Réu(s):\n{nomes_fmt}\n\n"
            f"Arquivo: {Path(file_path1).name}"
        )
        if not resposta:
            self._set_status("Envio cancelado pelo usuário.")
            return

        # 9. Envia ao PJe  (idêntico ao ref)
        self.btn_enviar.config(state="disabled", text="Enviando…")
        self.update_idletasks()

        url     = PJE_URL
        headers = {"Content-Type": "text/xml; charset=utf-8"}
        response = None

        try:
            response = requests.post(url, data=body, headers=headers, timeout=120)
            print(response.text)
        except requests.exceptions.Timeout:
            self.btn_enviar.config(state="normal", text="⚡  ENVIAR PETIÇÃO AO PJe")
            messagebox.showerror("Erro",
                "Timeout: o servidor PJe não respondeu em 2 minutos.\nTente novamente.")
            self._set_status("❌ Timeout na conexão com o PJe.")
            return
        except Exception as exc:
            self.btn_enviar.config(state="normal", text="⚡  ENVIAR PETIÇÃO AO PJe")
            messagebox.showerror("Erro",
                f"Ocorreu um erro ao tentar enviar a Petição Inicial ao PJe:\n{exc}")
            self._set_status(f"❌ Erro: {exc}")
            return

        self.btn_enviar.config(state="normal", text="⚡  ENVIAR PETIÇÃO AO PJe")

        # 10. Trata resposta  (idêntico ao ref)
        if response.status_code == 200:
            conteudo      = response.text
            numero_processo = re.search(r"\d{20}", conteudo)

            if numero_processo:
                numero_processo = numero_processo.group()
                self._set_status("Estou finalizando. Salvando dados do protocolo.")
                self.update_idletasks()

                # Salva no banco
                self.conn.execute("""
                    INSERT INTO peticoes_enviadas
                        (certidao_id, numero_processo_pje, comarca, cod_comarca,
                         caminho_pdf_assinado, data_envio, status, resposta_pje)
                    VALUES (?,?,?,?,?,?,?,?)
                """, (cert_id, numero_processo, comarca_envio, cod_comarca,
                      file_path1, datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                      "enviada", conteudo[:2000]))
                self.conn.commit()

                messagebox.showinfo(
                    "Info",
                    f"Inicial protocolizada com sucesso.\n"
                    f"Processo Nº {numero_processo} criado com sucesso.\n"
                    f"Salvei no Banco de Dados."
                )
                self._set_status(
                    f"✅ Protocolada! Processo PJe: {numero_processo}",
                    cor=CORES["success"])
                self._carregar_certidoes()
                self.tree_cert.selection_set(str(cert_id))
                self._ao_selecionar()

            else:
                # Registra erro
                self.conn.execute("""
                    INSERT INTO peticoes_enviadas
                        (certidao_id, comarca, cod_comarca, caminho_pdf_assinado,
                         data_envio, status, resposta_pje)
                    VALUES (?,?,?,?,?,?,?)
                """, (cert_id, comarca_envio, cod_comarca, file_path1,
                      datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                      "erro", conteudo[:2000]))
                self.conn.commit()

                messagebox.showerror(
                    "Erro",
                    "Algo deu errado. Verifique no terminal do sistema a informação.\n"
                    "Em seguida, tente novamente."
                )
                self._set_status("❌ Número do protocolo não encontrado na resposta do PJe.")
                print(conteudo)   # debug: resposta completa no terminal

        else:
            messagebox.showerror(
                "Erro",
                "Que Pena 😞 Ocorreu um erro ao tentar enviar a Petição ao PJe.\n"
                f"HTTP {response.status_code}"
            )
            self._set_status(f"❌ Erro HTTP {response.status_code}.")

    # ── Gerenciador de comarcas ────────────────────────────────

    def _abrir_gerenciador(self):
        win = GerenciadorComarcas(master=self, db_path=self.db_path)
        # Ao fechar, atualiza o combo
        win.bind("<Destroy>", lambda e: self._carregar_comarcas())


# ============================================================
# GERENCIADOR DE COMARCAS
# ============================================================

class GerenciadorComarcas(tk.Toplevel):
    """Cadastro/edição de comarcas e códigos de localidade PJe."""

    def __init__(self, master=None, db_path: str = DB_PATH):
        super().__init__(master)
        self.db_path = db_path
        self.conn    = abrir_banco(db_path)
        self.title("Gerenciar Comarcas — PJe TJPI")
        self.geometry("520x460")
        self.configure(bg=CORES["bg"])
        self.grab_set()
        self._build()
        self._carregar()

    def _build(self):
        tk.Label(self, text="Comarcas e Códigos de Localidade PJe",
                 bg=CORES["accent2"], fg="white",
                 font=("Segoe UI Semibold", 12)
                 ).pack(fill="x")

        cols = ("comarca", "codigo")
        self.tree = ttk.Treeview(self, columns=cols, show="headings", height=14)
        self.tree.heading("comarca", text="Comarca")
        self.tree.heading("codigo",  text="Cód. Localidade PJe")
        self.tree.column("comarca", width=310)
        self.tree.column("codigo",  width=160)
        self.tree.pack(fill="both", expand=True, padx=8, pady=8)

        form = tk.Frame(self, bg=CORES["panel"], pady=6)
        form.pack(fill="x", padx=8)
        form.columnconfigure(1, weight=1)
        for i, (lbl, var_attr, w) in enumerate([
                ("Comarca:", "var_c",   28),
                ("Código:",  "var_cod", 12)]):
            tk.Label(form, text=lbl, bg=CORES["panel"],
                     fg=CORES["sub"], font=FONTE
                     ).grid(row=i, column=0, sticky="w", padx=6, pady=2)
            setattr(self, var_attr, tk.StringVar())
            tk.Entry(form, textvariable=getattr(self, var_attr),
                     bg=CORES["panel2"], fg=CORES["texto"],
                     insertbackground=CORES["texto"],
                     relief="flat", font=FONTE, width=w
                     ).grid(row=i, column=1, sticky="ew" if w > 14 else "w",
                            padx=6, pady=2)

        bts = tk.Frame(self, bg=CORES["bg"])
        bts.pack(fill="x", padx=8, pady=8)
        for txt, cmd, cor in [
            ("➕ Adicionar", self._adicionar, CORES["accent2"]),
            ("✏ Editar",    self._editar,    CORES["panel2"]),
            ("🗑 Excluir",   self._excluir,   "#6b1f1f"),
        ]:
            tk.Button(bts, text=txt, command=cmd, bg=cor, fg="white",
                      relief="flat", font=FONTE, cursor="hand2", padx=10
                      ).pack(side="left", padx=4)

        self.tree.bind("<<TreeviewSelect>>", self._ao_sel)

    def _carregar(self):
        self.tree.delete(*self.tree.get_children())
        cur = self.conn.cursor()
        cur.execute("SELECT id, comarca, cod_comarca FROM cod_comarcas ORDER BY comarca")
        for row in cur.fetchall():
            self.tree.insert("", "end", iid=str(row["id"]),
                             values=(row["comarca"], row["cod_comarca"]))

    def _ao_sel(self, event=None):
        sel = self.tree.selection()
        if sel:
            v = self.tree.item(sel[0])["values"]
            self.var_c.set(v[0])
            self.var_cod.set(v[1])

    def _adicionar(self):
        c   = self.var_c.get().strip().upper()
        cod = self.var_cod.get().strip()
        if not c or not cod:
            messagebox.showwarning("Atenção", "Preencha comarca e código.", parent=self)
            return
        try:
            self.conn.execute(
                "INSERT INTO cod_comarcas (comarca, cod_comarca) VALUES (?,?)", (c, cod))
            self.conn.commit()
            self._carregar()
        except sqlite3.IntegrityError:
            messagebox.showerror("Erro", f"Comarca '{c}' já existe.", parent=self)

    def _editar(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Atenção", "Selecione uma comarca.", parent=self)
            return
        c   = self.var_c.get().strip().upper()
        cod = self.var_cod.get().strip()
        if not c or not cod:
            messagebox.showwarning("Atenção", "Preencha comarca e código.", parent=self)
            return
        self.conn.execute(
            "UPDATE cod_comarcas SET comarca=?, cod_comarca=? WHERE id=?",
            (c, cod, int(sel[0])))
        self.conn.commit()
        self._carregar()

    def _excluir(self):
        sel = self.tree.selection()
        if not sel:
            return
        nome = self.tree.item(sel[0])["values"][0]
        if messagebox.askyesno("Confirmar", f"Excluir comarca '{nome}'?", parent=self):
            self.conn.execute("DELETE FROM cod_comarcas WHERE id=?", (int(sel[0]),))
            self.conn.commit()
            self._carregar()


# ============================================================
# PONTO DE ENTRADA
# ============================================================

def abrir_enviador(master=None, db_path: str = DB_PATH) -> "EnviadorApp":
    """
    Integração com extrator_certidao.py:

        from enviador_peticao import abrir_enviador
        # em algum botão/menu do extrator:
        abrir_enviador(master=self, db_path=self.db_path)
    """
    return EnviadorApp(master=master, db_path=db_path)


if __name__ == "__main__":
    import sys
    db   = sys.argv[1] if len(sys.argv) > 1 else DB_PATH
    root = tk.Tk()
    root.withdraw()
    app  = EnviadorApp(master=root, db_path=db)
    app.protocol("WM_DELETE_WINDOW", root.destroy)
    root.mainloop()
