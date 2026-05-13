"""
Conversor de Portaria DOCX para HTML (Padrao Diario Oficial da Uniao)
Versao 2.6

Correcoes v2.6:
- CSS de impressao de tabelas corrigido:
  * break-inside: avoid nos <tr> (nunca corta linha no meio)
  * break-inside: auto na <table> dentro de @media print (permite continuar na proxima pagina)
  * display: table-header-group no <thead> (repete cabecalho em cada pagina)
"""

import mammoth
import zipfile
import hashlib
import re
import base64
import sys
from pathlib import Path


# ─────────────────────────────────────────────
#  LEITURA DO XML
# ─────────────────────────────────────────────

def extrair_dados_xml(caminho_docx):
    with zipfile.ZipFile(caminho_docx) as z:
        xml = z.read('word/document.xml').decode('utf-8', errors='replace')

    blocos = re.findall(r'<w:p[ >].*?</w:p>', xml, re.DOTALL)

    ementas = set()
    titulos_centro = set()
    cabecalho_inst = set()
    orgaos = set()
    alteracoes_linha = set()
    historico_versoes = set()

    ultimo_foi_titulo_portaria = False
    encontrou_titulo_principal = False

    for i, bp in enumerate(blocos):
        texto = re.sub(r'<[^>]+>', '', bp).strip()
        estilo = re.findall(r'<w:pStyle w:val="([^"]+)"', bp)
        ind = re.search(r'<w:ind[^>]*>', bp)
        jc = re.findall(r'<w:jc w:val="([^"]+)"', bp)

        left = 0
        if ind:
            lm = re.search(r'w:left="(\d+)"', ind.group())
            if lm:
                left = int(lm.group(1))

        sz_vals = re.findall(r'<w:sz w:val="(\d+)"', bp)
        sz_min = min(int(v) for v in sz_vals) if sz_vals else 9999

        # ── Historico de versoes ─────────────────────────────────────────────
        if ('Corpodetexto' in estilo
                and left == 252
                and sz_min <= 20
                and 'center' not in jc
                and texto):
            historico_versoes.add(re.sub(r'\s+', ' ', texto[:200]).strip())

        # ── Ementa ──────────────────────────────────────────────────────────
        if ultimo_foi_titulo_portaria and left >= 3000 and texto:
            ementas.add(re.sub(r'\s+', ' ', texto[:60]).strip())
            ultimo_foi_titulo_portaria = False

        if texto and not re.match(r'^PORTARIA\s+\w+\s+N', texto, re.IGNORECASE):
            if left < 3000:
                ultimo_foi_titulo_portaria = False

        # ── Titulo de portaria ───────────────────────────────────────────────
        if re.match(r'^PORTARIA\s+\w+\s+N', texto, re.IGNORECASE):
            ultimo_foi_titulo_portaria = True
            encontrou_titulo_principal = True

        # ── Cabecalho institucional ──────────────────────────────────────────
        elif not encontrou_titulo_principal and 'center' in jc and texto:
            cabecalho_inst.add(texto.strip())

        # ── Orgao ────────────────────────────────────────────────────────────
        if 'Ttulo1' in estilo and texto:
            orgaos.add(re.sub(r'\s+', ' ', texto.strip()))

        # ── Alteracoes centralizadas grandes (dou-retificacao) ───────────────
        if (texto
                and re.match(r'^(Alterada|Retificada|Alterado|Retificado)\s+pel', texto, re.IGNORECASE)
                and ('center' in jc or sz_min > 20)):
            alteracoes_linha.add(re.sub(r'\s+', ' ', texto[:120]))

        # ── Capitulo/secao centralizados ─────────────────────────────────────
        if 'center' in jc and texto:
            if re.match(r'^(CAP[IÍ]TULO|SE[CÇ][AÃ]O|SE[CÇ][AÃ]O|SE[cç][aã]o|T[IÍ]TULO)\b', texto, re.IGNORECASE):
                chave_c = re.sub(r'\s+', ' ', texto[:60]).strip()
                titulos_centro.add(chave_c)
                for j in range(i+1, min(i+4, len(blocos))):
                    tp = re.sub(r'<[^>]+>', '', blocos[j]).strip()
                    jp = re.findall(r'<w:jc w:val="([^"]+)"', blocos[j])
                    if tp and 'center' in jp:
                        titulos_centro.add(re.sub(r'\s+', ' ', tp[:80]).strip())
                        break

    # ── Dimensoes de imagens ─────────────────────────────────────────────────
    EPX = 914400 / 96
    try:
        with zipfile.ZipFile(caminho_docx) as zf:
            rels_raw = zf.read('word/_rels/document.xml.rels').decode('utf-8', errors='replace')
            rels = re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels_raw)
            rid_to_file = {rid: tgt.split('/')[-1] for rid, tgt in rels if 'media' in tgt}
            file_to_md5 = {}
            for rid, fname in rid_to_file.items():
                try:
                    data = zf.read(f'word/media/{fname}')
                    file_to_md5[fname] = hashlib.md5(data).hexdigest()
                except Exception:
                    pass

        file_dims = {}
        for d in re.findall(r'<w:drawing>(.*?)</w:drawing>', xml, re.DOTALL):
            rid = re.search(r'r:embed="([^"]+)"', d)
            ext = re.search(r'<wp:extent cx="(\d+)" cy="(\d+)"', d)
            if rid and ext:
                fname = rid_to_file.get(rid.group(1), '')
                if fname:
                    cx, cy = int(ext.group(1)), int(ext.group(2))
                    w, h = int(cx / EPX), int(cy / EPX)
                    prev = file_dims.get(fname, (0, 0))
                    if w * h > prev[0] * prev[1]:
                        file_dims[fname] = (w, h)

        md5_dims = {file_to_md5[f]: file_dims[f] for f in file_dims if f in file_to_md5}
    except Exception as e:
        print(f'  [AVISO] Nao foi possivel mapear dimensoes: {e}')
        md5_dims = {}

    return ementas, titulos_centro, md5_dims, cabecalho_inst, orgaos, alteracoes_linha, historico_versoes


# ─────────────────────────────────────────────
#  IMAGENS
# ─────────────────────────────────────────────

_md5_dims = {}

def encode_image(image):
    try:
        with image.open() as f:
            dados = f.read()
        enc = base64.b64encode(dados).decode('ascii')
        ct = getattr(image, 'content_type', 'image/jpeg')
        if callable(ct): ct = ct()
        if not ct: ct = 'image/jpeg'
        md5 = hashlib.md5(dados).hexdigest()
        w, h = _md5_dims.get(md5, (0, 0))
        if w > 0 and h > 0:
            if w <= 150 and h <= 150:
                classe = 'img-brasao'
                style = f'width:{w}px;height:{h}px;'
            else:
                classe = 'img-grande'
                style = f'max-width:min({w}px,100%);height:auto;'
        else:
            if len(dados) < 15000:
                classe = 'img-brasao'
                style = 'max-width:120px;'
            else:
                classe = 'img-grande'
                style = 'max-width:100%;height:auto;'
        return {'src': f'data:{ct};base64,{enc}', 'class': classe, 'style': style, 'alt': 'Imagem'}
    except Exception as e:
        print(f'   [AVISO] Imagem: {e}')
        return {'src': '', 'class': 'img-brasao', 'alt': ''}


def limpar_html(html):
    html = html.replace('&nbsp;', ' ').replace('\xa0', ' ')
    html = re.sub(r'<p>\s*\d+\s*</p>', '', html)
    html = re.sub(r'<(strong|em|u)>\s*</\1>', '', html)
    html = re.sub(r' {2,}', ' ', html)
    return html


# ─────────────────────────────────────────────
#  CLASSIFICACAO
# ─────────────────────────────────────────────

PADROES = [
    ('aviso_legal',     re.compile(r'Este texto n.o su', re.IGNORECASE)),
    ('asteriscos',      re.compile(r'\*{5,}')),
    ('cabecalho',       re.compile(r'^(Publicado em:|Edi|P.gina:|.rg.o:)', re.IGNORECASE)),
    ('retificacao',     re.compile(r'(Retificad[ao]|Alterad[ao]) pel', re.IGNORECASE)),
    ('titulo_portaria', re.compile(r'^PORTARIA\s+\w+\s+N', re.IGNORECASE)),
    ('preambulo',       re.compile(r'^O\s+(MINISTRO|SECRETAR|PRESIDENTE|DIRETOR)', re.IGNORECASE)),
    ('paragrafo_unico', re.compile(r'^Par.grafo .nico', re.IGNORECASE)),
    ('paragrafo_num',   re.compile(r'^.\s*\d+[o\u00ba]', re.IGNORECASE)),
    ('artigo',          re.compile(r'^Art\.?\s*\d+', re.IGNORECASE)),
    ('inciso',          re.compile(r'^[IVX]+\s*[-\u2013\u2014]')),
    ('alinea',          re.compile(r'^[a-z]\)\s')),
    ('assinatura',      re.compile(r'^[A-Z\u00c0-\u00dc ]{10,}$')),
    ('anexo_titulo',    re.compile(r'^ANEXO\b', re.IGNORECASE)),
    ('capitulo',        re.compile(r'^CAP[IÍ]TULO\b', re.IGNORECASE)),
    ('secao',           re.compile(r'^SE[CÇcç][AÃaã]O\b|^Se[çc][aã]o\b')),
    ('link_url',        re.compile(r'https?://')),
]

RE_SO_LINK = re.compile(r'^\s*<a\s[^>]*>.*?</a>\s*$', re.DOTALL)

def extrair_texto(frag):
    return re.sub(r'<[^>]+>', '', frag).strip()

def chave(texto):
    return re.sub(r'\s+', ' ', texto[:80]).strip()

def classificar(texto):
    for nome, p in PADROES:
        if p.search(texto):
            return nome
    return 'corpo'


# ─────────────────────────────────────────────
#  TRANSFORMACAO
# ─────────────────────────────────────────────

def transformar_html(html_bruto, ementas_xml, centros_xml, cabecalho_inst,
                     orgaos_xml, alteracoes_xml, historico_xml):

    def processar_p(match):
        inner = match.group(1).strip()

        if '<img' in inner:
            return f'<p class="dou-imagem">{inner}</p>'

        texto = extrair_texto(inner)
        if not texto:
            return ''

        # Cabecalho institucional
        if texto.strip() in cabecalho_inst:
            return f'<p class="dou-cabecalho-inst">{inner}</p>'

        # ── Historico ────────────────────────────────────────────────────────
        ch80_txt = re.sub(r'\s+', ' ', texto[:80]).strip()
        for hist in historico_xml:
            if ch80_txt == hist[:80] or hist[:80] == ch80_txt:
                return f'<p class="dou-historico">{inner}</p>'

        # Ementa
        ch = chave(texto)
        for e in ementas_xml:
            if ch.startswith(e[:50]) or e.startswith(ch[:50]):
                return f'<div class="dou-ementa">{inner}</div>'

        # Titulo centralizado
        for c in centros_xml:
            if ch == c or ch.startswith(c[:60]):
                if re.match(r'^(CAP[IÍ]TULO|SE[CÇcç][AÃaã]O|Se[çc][aã]o)\s+[IVX\d]', texto, re.IGNORECASE):
                    return f'<p class="dou-capitulo-num">{inner}</p>'
                else:
                    return f'<p class="dou-capitulo-titulo">{inner}</p>'

        # Link de titulo
        if RE_SO_LINK.match(inner) and re.search(r'portaria|n[o\u00ba]\s*[\d.,]+', texto, re.IGNORECASE):
            return f'<p class="dou-cabecalho dou-link-titulo">{inner}</p>'

        # Alteracoes centralizadas grandes
        ch120 = re.sub(r'\s+', ' ', texto[:120]).strip()
        for alt in alteracoes_xml:
            if ch120.startswith(alt[:60]) or alt.startswith(ch120[:60]):
                return f'<p class="dou-retificacao">{inner}</p>'

        classe = classificar(texto)

        if classe == 'titulo_portaria' and RE_SO_LINK.match(inner):
            return f'<p class="dou-cabecalho dou-link-titulo">{inner}</p>'

        mapa = {
            'cabecalho':       f'<p class="dou-cabecalho">{inner}</p>',
            'retificacao':     f'<p class="dou-retificacao">{inner}</p>',
            'aviso_legal':     f'<p class="dou-aviso-legal"><em>{texto}</em></p>',
            'asteriscos':      f'<p class="dou-asteriscos">{texto}</p>',
            'titulo_portaria': f'<h1 class="dou-titulo-portaria">{re.sub(r"</?strong>","",inner)}</h1>',
            'preambulo':       f'<p class="dou-preambulo">{inner}</p>',
            'artigo':          f'<p class="dou-artigo">{inner}</p>',
            'paragrafo_unico': f'<p class="dou-artigo">{inner}</p>',
            'paragrafo_num':   f'<p class="dou-artigo">{inner}</p>',
            'inciso':          f'<p class="dou-inciso">{inner}</p>',
            'alinea':          f'<p class="dou-alinea">{inner}</p>',
            'assinatura':      f'<p class="dou-assinatura">{inner}</p>',
            'anexo_titulo':    f'<h2 class="dou-anexo-titulo">{inner}</h2>',
            'capitulo':        f'<p class="dou-capitulo-num">{inner}</p>',
            'secao':           f'<p class="dou-capitulo-num">{inner}</p>',
            'link_url':        f'<p class="dou-link">{inner}</p>',
            'corpo':           f'<p class="dou-corpo">{inner}</p>',
        }
        return mapa.get(classe, f'<p class="dou-corpo">{inner}</p>')

    return re.sub(r'<p>(.*?)</p>', processar_p, html_bruto, flags=re.DOTALL)


def inserir_separador(html):
    padrao = re.compile(
        r'((?:<(?:p|h1)[^>]*class="dou-(?:cabecalho|retificacao|aviso.legal|imagem|link-titulo)[^"]*"[^>]*>.*?</(?:p|h1)>\s*)+)',
        re.DOTALL
    )
    return padrao.sub(r'\1<hr class="sep-cabecalho">', html, count=1)


# ─────────────────────────────────────────────
#  CSS
# ─────────────────────────────────────────────

CSS = """
*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

body {
    background: #d0d0d0;
    padding: 36px 16px 80px;
    font-family: 'Times New Roman', Times, serif;
    font-size: 14px;
    color: #000;
    line-height: 1.6;
}

.folha-dou {
    background: #fff;
    max-width: 870px;
    margin: 0 auto;
    padding: 50px 86px 70px;
    box-shadow: 0 3px 22px rgba(0,0,0,0.18);
    min-height: 1100px;
}

/* ─────────────────────────────────────────────
   IMPRESSAO / SALVAR PDF
   ───────────────────────────────────────────── */
@media print {
    body {
        background: white;
        padding: 0;
    }
    .folha-dou {
        box-shadow: none;
        padding: 18mm 24mm;
        max-width: 100%;
    }
    .btn-imprimir {
        display: none !important;
    }

    /* Tabelas grandes (maiores que uma pagina):
       - A tabela em si pode quebrar entre paginas (break-inside: auto)
       - Mas nunca corta uma linha no meio (break-inside: avoid no tr)
       - O cabecalho se repete em cada nova pagina (display: table-header-group) */
    table {
        break-inside: auto;
        page-break-inside: auto;
    }
    thead {
        display: table-header-group;
    }
    tr {
        break-inside: avoid;
        page-break-inside: avoid;
    }
}

.dou-imagem { text-align: center; margin: 10px 0 14px; }
.img-brasao  { display: block; margin: 0 auto; }
.img-grande  { display: block; margin: 12px auto; max-width: 100%; height: auto; }

.dou-orgao {
    text-align: center;
    font-family: Arial, sans-serif;
    font-size: 11.5px;
    margin: 4px 0;
    text-indent: 0 !important;
}

.dou-cabecalho-inst {
    text-align: center !important;
    font-family: Arial, sans-serif;
    font-size: 13px;
    font-weight: bold;
    margin: 3px 0;
    text-indent: 0 !important;
    letter-spacing: 0.03em;
}

.sep-cabecalho { border: none; border-top: 2px solid #000; margin: 14px 0 22px; }

.dou-cabecalho {
    text-align: center;
    font-family: Arial, sans-serif;
    font-size: 11.5px;
    margin: 3px 0;
    line-height: 1.4;
}
.dou-link-titulo { font-size: 12px; margin: 5px 0; }
.dou-link-titulo a, .dou-retificacao a { color: #003399; }

.dou-retificacao {
    text-align: center;
    font-family: Arial, sans-serif;
    font-size: 11.5px;
    margin: 3px 0;
}

.dou-historico {
    font-family: Arial, sans-serif;
    font-size: 10px;
    color: #111;
    margin: 1px 0 1px 1.8em;
    line-height: 1.5;
    text-align: left;
    text-indent: 0;
}
.dou-historico a {
    color: #003399;
    text-decoration: none;
}
.dou-historico a:hover { text-decoration: underline; }

.dou-aviso-legal {
    color: #cc0000;
    text-align: center;
    font-family: Arial, sans-serif;
    font-size: 11px;
    font-style: italic;
    margin: 8px 0 4px;
}

.dou-asteriscos {
    color: #cc0000;
    text-align: center;
    font-family: Arial, sans-serif;
    font-size: 12px;
    font-weight: bold;
    margin: 14px 0;
}

.dou-titulo-portaria {
    text-align: center;
    font-family: Arial, sans-serif;
    font-size: 15px;
    font-weight: bold;
    text-transform: uppercase;
    letter-spacing: 0.04em;
    margin: 28px 0 16px;
    line-height: 1.45;
}

.dou-ementa {
    margin-left: 46%;
    font-size: 13px;
    line-height: 1.5;
    margin-bottom: 22px;
    text-align: justify;
}

.dou-capitulo-num {
    text-align: center;
    font-family: Arial, sans-serif;
    font-size: 14px;
    font-weight: bold;
    text-transform: uppercase;
    margin: 28px 0 2px;
    letter-spacing: 0.03em;
}

.dou-capitulo-titulo {
    text-align: center;
    font-family: Arial, sans-serif;
    font-size: 14px;
    font-weight: bold;
    text-transform: uppercase;
    margin: 0 0 18px;
    letter-spacing: 0.03em;
}

.dou-preambulo {
    text-align: justify;
    font-size: 14px;
    margin-bottom: 12px;
    line-height: 1.7;
}

.dou-artigo {
    text-align: justify;
    font-size: 14px;
    text-indent: 2.5em;
    margin-bottom: 10px;
    line-height: 1.7;
}

.dou-inciso {
    text-align: justify;
    font-size: 14px;
    padding-left: 4em;
    margin-bottom: 8px;
    line-height: 1.7;
}

.dou-alinea {
    text-align: justify;
    font-size: 14px;
    padding-left: 6em;
    margin-bottom: 6px;
    line-height: 1.7;
}

.dou-corpo {
    text-align: justify;
    font-size: 14px;
    text-indent: 2.5em;
    margin-bottom: 10px;
    line-height: 1.7;
}

.dou-assinatura {
    text-align: center;
    font-family: Arial, sans-serif;
    font-size: 14px;
    font-weight: bold;
    text-transform: uppercase;
    letter-spacing: 0.03em;
    margin-top: 26px;
    margin-bottom: 4px;
}

.dou-anexo-titulo {
    text-align: center;
    font-family: Arial, sans-serif;
    font-size: 14px;
    font-weight: bold;
    text-transform: uppercase;
    letter-spacing: 0.04em;
    margin: 34px 0 16px;
    padding-top: 22px;
    border-top: 1px solid #999;
}

.dou-link { font-size: 12px; color: #222; margin: 3px 0 3px 2em; }
.dou-link a { color: #003399; }

ol, ul { margin: 6px 0 10px 4em; font-size: 14px; line-height: 1.7; }
li { margin-bottom: 4px; }

/* ─────────────────────────────────────────────
   TABELAS
   - Na tela: tenta manter a tabela inteira (break-inside: avoid)
   - Na impressao (ver @media print acima): permite quebrar entre paginas,
     mas nunca corta uma linha no meio
   ───────────────────────────────────────────── */
table {
    width: 100%;
    border-collapse: collapse;
    margin: 16px 0 20px;
    font-size: 12.5px;
    font-family: Arial, sans-serif;
    line-height: 1.4;
    table-layout: fixed;
    word-wrap: break-word;
    break-inside: avoid;
    page-break-inside: avoid;
}
thead {
    display: table-header-group;
}
table thead tr, table tr:first-child { background-color: #d9e1f2; }
table th, table td { border: 1px solid #666; padding: 5px 7px; vertical-align: top; text-align: left; }
table th { font-weight: bold; text-align: center; }
table tr:nth-child(even) { background-color: #f5f5f5; }
table tr:hover { background-color: #eef2fb; }
table p { text-indent: 0 !important; margin: 0; font-size: 12.5px; }

.btn-imprimir {
    display: block;
    margin: 28px auto 0;
    padding: 10px 30px;
    background: #1a3a6b;
    color: #fff;
    font-family: Arial, sans-serif;
    font-size: 14px;
    border: none;
    border-radius: 4px;
    cursor: pointer;
}
.btn-imprimir:hover { background: #0d2550; }
"""

TEMPLATE = """<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Portaria - Diario Oficial da Uniao</title>
  <style>{css}</style>
</head>
<body>
  <div class="folha-dou">
{corpo}
  </div>
  <button class="btn-imprimir" onclick="window.print()">Imprimir / Salvar PDF</button>
</body>
</html>
"""


# ─────────────────────────────────────────────
#  FUNCAO PRINCIPAL
# ─────────────────────────────────────────────

def converter(caminho_docx, caminho_html):
    docx = Path(caminho_docx)
    if not docx.exists():
        raise FileNotFoundError(f'Arquivo nao encontrado: {caminho_docx}')

    print(f'Lendo: {docx.name}')

    ementas_xml, centros_xml, dim_img, cabecalho_inst, orgaos_xml, alteracoes_xml, historico_xml = \
        extrair_dados_xml(caminho_docx)

    print(f'  Ementas: {len(ementas_xml)}')
    print(f'  Centros: {len(centros_xml)}')
    print(f'  Historico versoes: {len(historico_xml)}')
    for h in sorted(historico_xml):
        print(f'    -> {h[:100]}')
    print(f'  Imagens mapeadas: {len(dim_img)}')

    _md5_dims.clear()
    _md5_dims.update(dim_img)

    with open(docx, 'rb') as f:
        resultado = mammoth.convert_to_html(
            f, convert_image=mammoth.images.img_element(encode_image)
        )

    html = limpar_html(resultado.value)
    html = transformar_html(html, ementas_xml, centros_xml, cabecalho_inst,
                            orgaos_xml, alteracoes_xml, historico_xml)

    def rebaixar_h1_orgao(m):
        inner = m.group(1)
        texto_h1 = re.sub(r'<[^>]+>', '', inner).strip()
        ch = re.sub(r'\s+', ' ', texto_h1[:80])
        for org in orgaos_xml:
            if ch.startswith(org[:60]) or org.startswith(ch[:60]):
                return f'<p class="dou-orgao">{inner}</p>'
        return m.group(0)

    html = re.sub(r'<h1>(.*?)</h1>', rebaixar_h1_orgao, html, flags=re.DOTALL)
    html = inserir_separador(html)

    saida = TEMPLATE.format(css=CSS, corpo=html)
    with open(caminho_html, 'w', encoding='utf-8') as f:
        f.write(saida)

    kb = Path(caminho_html).stat().st_size / 1024
    print(f'Gerado: {caminho_html} ({kb:.1f} KB)')


if __name__ == '__main__':
    entrada = sys.argv[1] if len(sys.argv) > 1 else 'portaria.docx'
    saida   = sys.argv[2] if len(sys.argv) > 2 else 'portaria.html'
    converter(entrada, saida)
