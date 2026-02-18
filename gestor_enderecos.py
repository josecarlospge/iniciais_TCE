#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gestor_enderecos.py  --  v2.0
==============================
Extrai, parseia e persiste enderecos das partes interessadas a partir do
PDF de imputacao de debito do TCE/PI.

ESTRATEGIA DE EXTRACAO (dois estagios):
  1. pdfplumber detecta a estrutura da tabela (bounding boxes das celulas)
  2. pdf2image + pytesseract faz OCR em cada celula individualmente

Isso contorna o problema de encoding de fonte privada (PScript5/Acrobat
Distiller) que torna o texto extraido por pdfplumber ilegivel, mas preserva
a estrutura tabular necessaria para associar partes <-> enderecos.

Colunas do PDF de enderecos:
  Col 1 -- Tipo de certidao / cabecalho
  Col 2 -- Numero do processo TC/XXXXXX/XXXX
  Col 3 -- Objeto (imputacao de debito)
  Col 4 -- Partes interessadas (N. NOME CPF/CNPJ)
  Col 5 -- Enderecos (N. LOGRADOURO, No - BAIRRO - CEP - CIDADE/UF)
"""

import re
import sqlite3
import os
from dataclasses import dataclass
from typing import Optional


# ============================================================
# ESTRUTURAS DE DADOS
# ============================================================

@dataclass
class EnderecoEstruturado:
    tipo_logradouro: str = ""
    logradouro: str = ""
    numero: str = ""
    complemento: str = ""
    bairro: str = ""
    cep: str = ""
    municipio: str = ""
    uf: str = "PI"
    pais: str = "BR"
    endereco_original: str = ""

    def to_xml_fragment(self) -> str:
        logradouro_completo = f"{self.tipo_logradouro} {self.logradouro}".strip()
        linhas = [
            f'<int:endereco cep="{self.cep}">',
            f'    <int:logradouro>{logradouro_completo}</int:logradouro>',
            f'    <int:numero>{self.numero}</int:numero>',
        ]
        if self.complemento:
            linhas.append(f'    <int:complemento>{self.complemento}</int:complemento>')
        linhas += [
            f'    <int:bairro>{self.bairro}</int:bairro>',
            f'    <int:cidade>{self.municipio}</int:cidade>',
            f'    <int:estado>{self.uf}</int:estado>',
            f'    <int:pais>{self.pais}</int:pais>',
            '</int:endereco>',
        ]
        return '\n'.join(linhas)

    def formatado(self) -> str:
        partes = []
        if self.tipo_logradouro or self.logradouro:
            partes.append(f"{self.tipo_logradouro} {self.logradouro}".strip())
        if self.numero:
            partes.append(self.numero)
        if self.complemento:
            partes.append(self.complemento)
        if self.bairro:
            partes.append(self.bairro)
        if self.cep and len(self.cep) == 8:
            partes.append(f"CEP {self.cep[:5]}-{self.cep[5:]}")
        elif self.cep:
            partes.append(self.cep)
        if self.municipio:
            partes.append(f"{self.municipio}/{self.uf}")
        return " - ".join(partes)


@dataclass
class ParteComEndereco:
    numero_parte: int
    nome: str = ""
    cpf_cnpj: str = ""
    email: str = ""
    numero_processo: str = ""
    endereco: Optional[EnderecoEstruturado] = None


# ============================================================
# TIPOS DE LOGRADOURO
# ============================================================

_TIPOS_LOGRADOURO = [
    "AVENIDA", "AV", "RUA", "R", "TRAVESSA", "TV", "ALAMEDA", "AL",
    "RODOVIA", "ROD", "ESTRADA", "EST", "PRACA", "PC", "LARGO",
    "BECO", "VILA", "CONDOMINIO", "COND", "QUADRA", "QD", "CONJUNTO",
    "LOTEAMENTO", "SETOR", "AREA", "LOTE", "SERVIDAO", "COMUNIDADE",
    "MUCUNA", "LOCALIDADE",
]

_RE_TIPO = re.compile(
    r'^(' + '|'.join(sorted(_TIPOS_LOGRADOURO, key=len, reverse=True)) + r')\b[.]?\s*',
    re.IGNORECASE
)


# ============================================================
# PARSER DE ENDERECO
# ============================================================

def parsear_endereco(endereco_bruto: str) -> EnderecoEstruturado:
    """
    Converte string de endereco em EnderecoEstruturado.
    Formatos suportados:
      MUCUNA, S/N - CENTRO - CEP 64790-000 - DOM INOCENCIO/PI
      RUA Y, 303, APT. 1801 - CEP 64048-130 - JOQUEI - TERESINA/PI
      LOTEAMENTO - RUA 1001, QUADRA 1003, LOTE 28, No 154 - BAIRRO - CEP - CIDADE/UF
    """
    original = re.sub(r'\s+', ' ', endereco_bruto).strip()
    end = EnderecoEstruturado(endereco_original=original)

    # 1. Extrai CIDADE/UF
    m_cidade = re.search(r'([A-Z][A-Z\s\-]*)/([A-Z]{2})\s*$', original, re.IGNORECASE)
    if m_cidade:
        cidade_raw = m_cidade.group(1).strip()
        partes_cidade = [p.strip() for p in re.split(r'\s*-\s*', cidade_raw) if p.strip()]
        end.municipio = partes_cidade[-1].title() if partes_cidade else cidade_raw.title()
        bairros_cid = partes_cidade[:-1]
        end.uf = m_cidade.group(2).upper()
        original = original[:m_cidade.start()].rstrip(' -,')
        if bairros_cid and not end.bairro:
            end.bairro = " / ".join(b.title() for b in bairros_cid)

    # 2. Extrai CEP
    m_cep = re.search(r'CEP\s*(\d{5}[-.]?\d{3})', original, re.IGNORECASE)
    if m_cep:
        end.cep = re.sub(r'[-.]', '', m_cep.group(1))
        trecho_apos = original[m_cep.end():].strip(' -,')
        original = original[:m_cep.start()].rstrip(' -,')
        if trecho_apos:
            segs = [s.strip() for s in trecho_apos.split(' - ') if s.strip()]
            bairro_pos = [s for s in segs if '/' not in s]
            if bairro_pos and not end.bairro:
                end.bairro = " / ".join(bairro_pos).title()

    # 3. Processa logradouro restante
    partes_dash = [p.strip() for p in re.split(r'\s*-\s*', original) if p.strip()]
    logradouro_bruto = partes_dash[0] if partes_dash else ""
    bairro_candidatos = [
        re.sub(r'^BAIRRO\s+', '', p, flags=re.IGNORECASE).strip()
        for p in partes_dash[1:]
    ]
    if bairro_candidatos and not end.bairro:
        end.bairro = " / ".join(b.title() for b in bairro_candidatos)

    # 4. Tipo + nome + numero + complemento
    logradouro_bruto = re.sub(r'\bN[o]\s*', '', logradouro_bruto)
    segmentos = [s.strip() for s in logradouro_bruto.split(',')]
    seg0 = segmentos[0] if segmentos else ""

    m_tipo = _RE_TIPO.match(seg0)
    if m_tipo:
        end.tipo_logradouro = m_tipo.group(1).upper().rstrip('.')
        end.logradouro = seg0[m_tipo.end():].strip().title()
    else:
        end.tipo_logradouro = "RUA"
        end.logradouro = seg0.strip().title()

    if len(segmentos) >= 2:
        poss_num = segmentos[1].strip()
        if re.match(r'^(S/?N|SN|\d[\d\s./]*|[QLTB]\s*\d.*)$', poss_num, re.IGNORECASE):
            end.numero = poss_num.upper()
            if len(segmentos) >= 3:
                end.complemento = ", ".join(segmentos[2:]).strip().title()
        else:
            end.numero = "S/N"
            end.complemento = ", ".join(segmentos[1:]).strip().title()

    if not end.numero:
        end.numero = "S/N"

    return end


# ============================================================
# OCR POR CELULA
# ============================================================

def _configurar_ocr() -> tuple:
    """
    Detecta e configura Poppler e Tesseract.

    Prioridade Poppler:
      1. Pasta local  <script>/poppler-25.12.0/Library/bin  (Windows portavel)
      2. PATH do sistema (Linux/Mac)

    Prioridade Tesseract:
      1. Executavel local  <script>/poppler-25.12.0/Library/bin/tesseract.exe
      2. C:/Program Files/Tesseract-OCR/tesseract.exe
      3. PATH do sistema

    tessdata:
      1. <poppler_bin>/tessdata  (pacote local que ja inclui por.traineddata)
      2. Diretorio padrao do Tesseract instalado

    Retorna (poppler_path_str_ou_None, tesseract_cmd_str).
    """
    import shutil
    import pytesseract
    from pathlib import Path

    base_dir = Path(__file__).parent
    poppler_local = base_dir / "poppler-25.12.0" / "Library" / "bin"

    # --- Poppler ---
    if poppler_local.exists():
        poppler_path = str(poppler_local)
    else:
        poppler_path = None   # usa PATH do sistema (Linux/Mac)

    # --- Tesseract ---
    tess_local = poppler_local / "tesseract.exe"          # Windows portatil
    tess_win   = Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe")

    if tess_local.exists():
        tess_cmd = str(tess_local)
    elif tess_win.exists():
        tess_cmd = str(tess_win)
    else:
        tess_cmd = shutil.which("tesseract") or "tesseract"   # Linux/Mac

    pytesseract.pytesseract.tesseract_cmd = tess_cmd

    # --- tessdata (pacote de linguas) ---
    tessdata_local = poppler_local / "tessdata"
    if tessdata_local.exists():
        os.environ["TESSDATA_PREFIX"] = str(tessdata_local)
    elif tess_win.exists():
        os.environ["TESSDATA_PREFIX"] = str(tess_win.parent / "tessdata")
    # Se nenhum dos dois existir, mantem o TESSDATA_PREFIX do ambiente

    return poppler_path, tess_cmd


def _ocr_cell(img_page, cell_bbox_pts, dpi_scale: float, lang: str = "por+eng") -> str:
    from PIL import Image
    import pytesseract
    x0, top, x1, bottom = cell_bbox_pts
    m = 4
    px0, py0 = max(0, int(x0 * dpi_scale) + m), max(0, int(top * dpi_scale) + m)
    px1, py1 = int(x1 * dpi_scale) - m, int(bottom * dpi_scale) - m
    if px1 <= px0 or py1 <= py0:
        return ""
    crop = img_page.crop((px0, py0, px1, py1))
    w, h = crop.size
    crop = crop.resize((w * 2, h * 2), Image.LANCZOS)
    return pytesseract.image_to_string(crop, lang=lang, config="--psm 6").strip()


def extrair_tabela_ocr(caminho_pdf: str, dpi: int = 250,
                        verbose: bool = True) -> list:
    """
    Extrai linhas da tabela via OCR por celula.
    Retorna lista de dicts: {pagina, numero_processo, partes_texto, enderecos_texto}
    """
    import pdfplumber
    from pdf2image import convert_from_path
    import pytesseract
    from pathlib import Path

    poppler_path, tess_cmd = _configurar_ocr()
    if verbose:
        print(f"  Tesseract : {tess_cmd}")
        print(f"  Poppler   : {poppler_path or '(PATH do sistema)'}")
        tdata = os.environ.get("TESSDATA_PREFIX", "(padrao do sistema)")
        print(f"  tessdata  : {tdata}")

    scale = dpi / 72.0
    if verbose:
        print(f"  Convertendo PDF para imagens (DPI={dpi})...")
    kwargs = {"dpi": dpi}
    if poppler_path:
        kwargs["poppler_path"] = poppler_path
    imgs = convert_from_path(caminho_pdf, **kwargs)

    resultados = []
    with pdfplumber.open(caminho_pdf) as pdf:
        for p_idx, (pag, img) in enumerate(zip(pdf.pages, imgs)):
            tbl_objs = pag.find_tables()
            if not tbl_objs:
                continue
            rows = tbl_objs[0].rows
            if verbose:
                print(f"  Pag {p_idx+1}: {max(0, len(rows)-2)} linha(s) de dados")
            for r_idx, row in enumerate(rows):
                if r_idx <= 1:
                    continue
                cells = row.cells
                if len(cells) < 5:
                    continue
                if not all(cells[i] for i in [1, 3, 4]):
                    continue
                proc  = _ocr_cell(img, cells[1], scale)
                partes = _ocr_cell(img, cells[3], scale)
                endrs  = _ocr_cell(img, cells[4], scale)
                proc   = re.sub(r'\s+', ' ', proc).strip()
                partes = re.sub(r'[ \t]+', ' ', partes).strip()
                endrs  = re.sub(r'[ \t]+', ' ', endrs).strip()
                if proc or partes:
                    resultados.append({
                        "pagina": p_idx + 1,
                        "numero_processo": _normalizar_processo(proc),
                        "partes_texto": partes,
                        "enderecos_texto": endrs,
                    })
    return resultados


def _normalizar_processo(txt: str) -> str:
    txt = re.sub(r'\s', '', txt)
    m = re.search(r'TC[/,]?(\d{6})[/,]?(\d{4})', txt, re.IGNORECASE)
    if m:
        return f"TC/{m.group(1)}/{m.group(2)}"
    return txt.strip()


# ============================================================
# PARSER DE PARTES E ENDERECOS
# ============================================================

def parsear_partes(partes_texto: str) -> list:
    """
    Analisa o texto da coluna de partes e extrai lista de responsaveis.

    Tratamento especial de quebra de linha do OCR:
    - CPF/CNPJ numa linha sozinha (sem texto antes) NAO e tratado como inicio
      de nova parte — e tratado como continuacao do responsavel atual.
    - Exemplo: "152.308.643-20" sozinho na linha nao dispara nova parte "152".
    """
    partes = []
    linhas = partes_texto.splitlines()

    # Numero de parte: "2 - NOME" ou "2. NOME" ou "2) NOME"
    # Mas NAO deve casar com CPF ("152.308.643-20") nem CNPJ ("10.904.554/0001-77")
    re_inicio   = re.compile(r'^(\d{1,2})\s*[-.)|\s]\s*(.+)', re.IGNORECASE)

    re_cpf_fmt  = re.compile(r'CPF\s*[:.#]?\s*([\d]{3}[.\s][\d]{3}[.\s][\d]{3}[-\s][\d]{2})', re.IGNORECASE)
    re_cnpj_fmt = re.compile(r'CNPJ\s*[:.#]?\s*([\d]{2}[.\s][\d]{3}[.\s][\d]{3}[/\s][\d]{4}[-\s][\d]{2})', re.IGNORECASE)
    re_cpf_raw  = re.compile(r'([\d]{3}\.[\d]{3}\.[\d]{3}-[\d]{2})')
    re_cnpj_raw = re.compile(r'([\d]{2}\.[\d]{3}\.[\d]{3}/[\d]{4}-[\d]{2})')
    re_email    = re.compile(r'[\w.+\-]+@[\w.\-]+\.[a-z]{2,}', re.IGNORECASE)

    # Linha que e SO um CPF ou CNPJ (sem texto antes/depois relevante)
    re_linha_so_cpf  = re.compile(r'^\s*\d{3}\.\d{3}\.\d{3}-\d{2}\s*$')
    re_linha_so_cnpj = re.compile(r'^\s*\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}\s*$')

    num_atual   = None
    nome_linhas = []
    doc_atual   = ""
    tipo_doc    = ""
    email_atual = ""

    def _e_linha_doc(linha):
        """Verifica se a linha inteira e um CPF ou CNPJ (sem prefixo de parte)."""
        return bool(re_linha_so_cpf.match(linha) or re_linha_so_cnpj.match(linha))

    def _extrair_doc(linha):
        m = re_cpf_fmt.search(linha) or re_cpf_raw.search(linha)
        if m:
            return re.sub(r'[\s]', '', m.group(1)), "CPF"
        m = re_cnpj_fmt.search(linha) or re_cnpj_raw.search(linha)
        if m:
            return re.sub(r'[\s]', '', m.group(1)), "CNPJ"
        return "", ""

    def _limpar_nome(linhas_lista):
        nome = " ".join(linhas_lista)
        for pat in [re_cpf_fmt, re_cnpj_fmt, re_cpf_raw, re_cnpj_raw, re_email]:
            nome = pat.sub('', nome)
        nome = re.sub(r'\b(CPF|CNPJ)\b', '', nome, flags=re.IGNORECASE)
        nome = re.sub(r'\b\d{2}\s+\d{4,5}-\d{4}\b', '', nome)  # fone
        nome = re.sub(r'\s+', ' ', nome).strip(' -,.:#')
        return nome.title()

    def _salvar():
        if num_atual is not None:
            partes.append({
                "numero_parte": num_atual,
                "nome": _limpar_nome(nome_linhas),
                "cpf_cnpj": doc_atual,
                "tipo_doc": tipo_doc,
                "email": email_atual,
            })

    for linha in linhas:
        linha = linha.strip()
        if not linha:
            continue

        doc, tdoc = _extrair_doc(linha)
        m_email = re_email.search(linha)

        # Verifica se e linha de CPF/CNPJ puro (sem numerador de parte)
        # Nesse caso NUNCA trata como inicio de nova parte
        if _e_linha_doc(linha):
            if doc and not doc_atual:
                doc_atual = doc
                tipo_doc  = tdoc
            # CPF/CNPJ sozinho nao adiciona nada ao nome
            continue

        # Tenta casar como inicio de nova parte
        m_ini = re_inicio.match(linha)

        if m_ini:
            _salvar()
            num_atual   = int(m_ini.group(1))
            nome_linhas = [m_ini.group(2).strip()]
            doc_atual   = doc
            tipo_doc    = tdoc
            email_atual = m_email.group(0) if m_email else ""
        else:
            # Continuacao: nome extra, email, ou doc embutido no texto
            if doc and not doc_atual:
                doc_atual = doc
                tipo_doc  = tdoc
            if m_email and not email_atual:
                email_atual = m_email.group(0)
            if num_atual is not None:
                linha_limpa = re_cpf_fmt.sub('', re_cnpj_fmt.sub('', linha))
                linha_limpa = re_cpf_raw.sub('', re_cnpj_raw.sub('', linha_limpa))
                linha_limpa = re_email.sub('', linha_limpa)
                linha_limpa = re.sub(r'\b(CPF|CNPJ)\b', '', linha_limpa, flags=re.IGNORECASE)
                linha_limpa = re.sub(r'\b\d{2}\s+\d{4,5}-\d{4}\b', '', linha_limpa)
                linha_limpa = linha_limpa.strip(' -,.:')
                if linha_limpa:
                    nome_linhas.append(linha_limpa)
            else:
                # Parte unica sem numerador
                num_atual   = 1
                nome_linhas = []
                doc_atual   = doc
                tipo_doc    = tdoc
                email_atual = m_email.group(0) if m_email else ""
                linha_limpa = re_cpf_fmt.sub('', re_cnpj_fmt.sub('', linha))
                linha_limpa = re_cpf_raw.sub('', re_cnpj_raw.sub('', linha_limpa))
                linha_limpa = re_email.sub('', linha_limpa).strip(' -,.:')
                if linha_limpa:
                    nome_linhas.append(linha_limpa)

    _salvar()
    return partes


def parsear_enderecos_coluna(enderecos_texto: str) -> dict:
    mapa = {}
    linhas = enderecos_texto.splitlines()
    re_num = re.compile(r'^(\d+)\s*[-.)|\s]\s*(.+)', re.IGNORECASE)
    num_atual = None
    acumulo = []

    def _salvar():
        if num_atual is not None and acumulo:
            mapa[num_atual] = ' '.join(acumulo)

    for linha in linhas:
        linha = linha.strip()
        if not linha:
            continue
        m = re_num.match(linha)
        if m:
            _salvar()
            num_atual = int(m.group(1))
            acumulo = [m.group(2).strip()]
        else:
            if num_atual is not None:
                acumulo.append(linha)
            else:
                num_atual = 1
                acumulo = [linha]

    _salvar()

    if not mapa and enderecos_texto.strip():
        mapa[1] = ' '.join(enderecos_texto.splitlines())

    return mapa


# ============================================================
# PIPELINE PRINCIPAL
# ============================================================

def extrair_partes_e_enderecos_ocr(caminho_pdf: str,
                                    dpi: int = 250,
                                    verbose: bool = True) -> list:
    linhas = extrair_tabela_ocr(caminho_pdf, dpi=dpi, verbose=verbose)
    resultados = []

    for linha in linhas:
        proc   = linha["numero_processo"]
        partes = parsear_partes(linha["partes_texto"])
        endmap = parsear_enderecos_coluna(linha["enderecos_texto"])

        for p in partes:
            num = p["numero_parte"]
            end_bruto = endmap.get(num, "")
            end_struct = parsear_endereco(end_bruto) if end_bruto else None
            resultados.append(ParteComEndereco(
                numero_parte=num,
                nome=p["nome"],
                cpf_cnpj=p["cpf_cnpj"],
                email=p.get("email", ""),
                numero_processo=proc,
                endereco=end_struct,
            ))
            if verbose and end_struct:
                print(f"    OK Parte {num} ({p['cpf_cnpj'] or p['nome'][:30]}) "
                      f"-> {end_struct.formatado()[:60]}")

    return resultados


# ============================================================
# BANCO DE DADOS
# ============================================================

DDL_ENDERECOS = """
CREATE TABLE IF NOT EXISTS enderecos_responsavel (
    id                  INTEGER PRIMARY KEY AUTOINCREMENT,
    responsavel_id      INTEGER REFERENCES responsaveis(id),
    certidao_id         INTEGER REFERENCES certidoes(id),
    numero_doc          TEXT,
    tipo_logradouro     TEXT,
    logradouro          TEXT,
    numero              TEXT,
    complemento         TEXT,
    bairro              TEXT,
    cep                 TEXT,
    municipio           TEXT,
    uf                  TEXT DEFAULT 'PI',
    pais                TEXT DEFAULT 'BR',
    endereco_original   TEXT,
    data_insercao       TEXT DEFAULT CURRENT_TIMESTAMP
);
CREATE INDEX IF NOT EXISTS idx_end_numero_doc
    ON enderecos_responsavel(numero_doc);
CREATE INDEX IF NOT EXISTS idx_end_responsavel
    ON enderecos_responsavel(responsavel_id);
"""


def criar_tabela_enderecos(conn: sqlite3.Connection) -> None:
    conn.executescript(DDL_ENDERECOS)
    conn.commit()


def salvar_endereco(conn: sqlite3.Connection,
                    end: EnderecoEstruturado,
                    numero_doc: str = "",
                    responsavel_id: int = None,
                    certidao_id: int = None) -> int:
    cursor = conn.cursor()
    existe_id = None
    if numero_doc:
        cursor.execute("SELECT id FROM enderecos_responsavel WHERE numero_doc=?", (numero_doc,))
        row = cursor.fetchone()
        if row:
            existe_id = row[0]
    elif responsavel_id:
        cursor.execute("SELECT id FROM enderecos_responsavel WHERE responsavel_id=?", (responsavel_id,))
        row = cursor.fetchone()
        if row:
            existe_id = row[0]

    if existe_id:
        cursor.execute("""
            UPDATE enderecos_responsavel SET
                tipo_logradouro=?,logradouro=?,numero=?,complemento=?,
                bairro=?,cep=?,municipio=?,uf=?,pais=?,endereco_original=?
            WHERE id=?
        """, (end.tipo_logradouro, end.logradouro, end.numero, end.complemento,
              end.bairro, end.cep, end.municipio, end.uf, end.pais,
              end.endereco_original, existe_id))
        conn.commit()
        return existe_id
    else:
        cursor.execute("""
            INSERT INTO enderecos_responsavel
                (responsavel_id,certidao_id,numero_doc,
                 tipo_logradouro,logradouro,numero,complemento,
                 bairro,cep,municipio,uf,pais,endereco_original)
            VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)
        """, (responsavel_id, certidao_id, numero_doc,
              end.tipo_logradouro, end.logradouro, end.numero, end.complemento,
              end.bairro, end.cep, end.municipio, end.uf, end.pais,
              end.endereco_original))
        conn.commit()
        rid = cursor.lastrowid
        if responsavel_id:
            cursor.execute("UPDATE responsaveis SET endereco=? WHERE id=?",
                           (end.endereco_original, responsavel_id))
            conn.commit()
        return rid


def vincular_enderecos_por_doc(conn: sqlite3.Connection) -> int:
    cursor = conn.cursor()
    cursor.execute("""
        UPDATE enderecos_responsavel
        SET responsavel_id=(
            SELECT r.id FROM responsaveis r
            WHERE r.numero_doc=enderecos_responsavel.numero_doc LIMIT 1
        )
        WHERE responsavel_id IS NULL
          AND numero_doc IS NOT NULL AND numero_doc!=''
    """)
    conn.commit()
    return cursor.rowcount


def vincular_e_atualizar_endereco_texto(conn: sqlite3.Connection) -> int:
    cursor = conn.cursor()
    cursor.execute("""
        UPDATE responsaveis
        SET endereco=(
            SELECT e.endereco_original FROM enderecos_responsavel e
            WHERE e.numero_doc=responsaveis.numero_doc LIMIT 1
        )
        WHERE (endereco IS NULL OR endereco='')
          AND numero_doc IS NOT NULL AND numero_doc!=''
    """)
    conn.commit()
    return cursor.rowcount


def buscar_endereco_responsavel(conn: sqlite3.Connection,
                                numero_doc: str = None,
                                responsavel_id: int = None) -> Optional[EnderecoEstruturado]:
    cursor = conn.cursor()
    if numero_doc:
        cursor.execute("SELECT * FROM enderecos_responsavel WHERE numero_doc=? LIMIT 1", (numero_doc,))
    elif responsavel_id:
        cursor.execute("SELECT * FROM enderecos_responsavel WHERE responsavel_id=? LIMIT 1", (responsavel_id,))
    else:
        return None
    row = cursor.fetchone()
    if not row:
        return None
    cols = [d[0] for d in cursor.description]
    d = dict(zip(cols, row))
    return EnderecoEstruturado(
        tipo_logradouro=d.get("tipo_logradouro", ""),
        logradouro=d.get("logradouro", ""),
        numero=d.get("numero", ""),
        complemento=d.get("complemento", ""),
        bairro=d.get("bairro", ""),
        cep=d.get("cep", ""),
        municipio=d.get("municipio", ""),
        uf=d.get("uf", "PI"),
        pais=d.get("pais", "BR"),
        endereco_original=d.get("endereco_original", ""),
    )


# ============================================================
# PROCESSAMENTO COMPLETO DO PDF
# ============================================================

def _responsaveis_sem_endereco(conn: sqlite3.Connection,
                               certidao_id: int) -> list:
    """Retorna lista de (id, numero_doc) dos responsaveis ainda sem endereco."""
    cur = conn.cursor()
    cur.execute("""
        SELECT r.id, r.numero_doc
        FROM responsaveis r
        WHERE r.certidao_id = ?
          AND (r.endereco IS NULL OR r.endereco = '')
          AND r.numero_doc IS NOT NULL AND r.numero_doc != '' 
    """, (certidao_id,))
    return cur.fetchall()


def processar_pdf_enderecos(caminho_pdf: str,
                             conn: sqlite3.Connection,
                             certidao_id: int = None,
                             verbose: bool = True) -> list:
    """
    Pipeline publico:
      1. Cria tabela se necessario
      2. OCR + parse de todas as partes e enderecos
      3. Persiste no banco (vincula por CPF/CNPJ)
      4. Relatorio de responsaveis sem endereco encontrado
    """
    criar_tabela_enderecos(conn)

    if verbose:
        print(f"  Extraindo enderecos via OCR por celula...")

    partes = extrair_partes_e_enderecos_ocr(caminho_pdf, verbose=verbose)

    salvos = 0
    for parte in partes:
        if parte.endereco and parte.cpf_cnpj:
            salvar_endereco(conn, parte.endereco,
                            numero_doc=parte.cpf_cnpj,
                            certidao_id=certidao_id)
            salvos += 1

    vinculados = vincular_enderecos_por_doc(conn)
    atualizados = vincular_e_atualizar_endereco_texto(conn)

    if verbose:
        print(f"  {salvos} endereco(s) gravado(s) no banco")
        print(f"  {vinculados} vinculo(s), {atualizados} campo(s) texto atualizado(s)")

    if certidao_id:
        sem_end = _responsaveis_sem_endereco(conn, certidao_id)
        if sem_end and verbose:
            print(f"\n  ATENCAO: {len(sem_end)} responsavel(is) sem endereco:")
            for rid, doc in sem_end:
                print(f"    id={rid} doc={doc} -> nao encontrado no PDF de enderecos")
            print(f"  (Campo endereco ficara em branco na peticao)")

    return partes


# ============================================================
# DEMO / CLI
# ============================================================

def _demo_parser():
    exemplos = [
        "MUCUNA, S/N - CENTRO - CEP 64790-000 - DOM INOCENCIO/PI",
        "COMUNIDADE BAIXA VERDE, S/N - ZONA RURAL - CEP 64790-000 - DOM INOCENCIO/PI",
        "RUA BEZERRA DO MEL, 115, BAIRRO BOA SORTE - APT. 101 - CEP 64607-140 - PICOS/PI",
        "AVENIDA MARECHAL CASTELO BRANCO, S/N - ILHOTAS - CEP 64014-058 - TERESINA/PI",
        "RUA MIOSOTES, 303, APT. 1801 - CEP 64048-130 - JOQUEI - TERESINA/PI",
        "PRAÇA CENTRAL, S/N, CENTRO - CEP 64728-000 - PEDRO LAURENTINO/PI",
    ]
    print("=" * 72)
    print("DEMO -- parsear_endereco()")
    print("=" * 72)
    for ex in exemplos:
        e = parsear_endereco(ex)
        print(f"\nEntrada : {ex}")
        print(f"  tipo   : {e.tipo_logradouro}  |  logradouro : {e.logradouro}")
        print(f"  numero : {e.numero}  |  compl : {e.complemento}")
        print(f"  bairro : {e.bairro}  |  cep : {e.cep}  |  cidade : {e.municipio}/{e.uf}")
        print(f"  {e.formatado()}")
        print("-" * 72)


if __name__ == "__main__":
    import sys

    if len(sys.argv) == 1:
        _demo_parser()
    elif len(sys.argv) >= 2:
        pdf  = sys.argv[1]
        db   = sys.argv[2] if len(sys.argv) >= 3 else "certidoes_tce.db"
        cid  = int(sys.argv[3]) if len(sys.argv) >= 4 else None
        conn = sqlite3.connect(db)
        partes = processar_pdf_enderecos(pdf, conn, cid, verbose=True)
        conn.close()
        print(f"\nTotal: {len(partes)} parte(s) processada(s).")
