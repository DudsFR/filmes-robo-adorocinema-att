#!/usr/bin/env python3
"""
Robô que mantém o arquivo filmes.csv atualizado.

- Filmes (título, título original, lançamento, direção, roteiro, sinopse): AdoroCinema
- Oscars (ganhou? / quais): Wikidata (base aberta da Wikipédia)

Regras:
- Filmes novos em cartaz entram com todas as colunas preenchidas.
- Em filmes que já estão na planilha, o robô só preenche células VAZIAS.
  Ele nunca apaga nem reescreve o que você já preencheu.
- O Oscar só é atualizado em linhas com "Aguardando" ou vazias.
"""
import csv
import html
import json
import re
import sys
import time
import unicodedata
from datetime import date
from pathlib import Path

import requests
from bs4 import BeautifulSoup

# ================= CONFIGURAÇÕES (pode mudar à vontade) =================
CONTATO = "seu-email@exemplo.com"   # o Wikidata pede um contato de quem consulta
QUANTOS_POR_VEZ = 10                # só os N filmes em mais salas de cinema entram
PAUSA = 2                           # segundos entre acessos, para não sobrecarregar o site
# =========================================================================

ARQUIVO = Path("filmes.csv")
CONTROLE = Path("controle.json")    # lembra quais filmes do AdoroCinema já foram vistos
LISTA_URL = "https://www.adorocinema.com/filmes/numero-cinemas/"
FILME_URL = "https://www.adorocinema.com/filmes/filme-{}/"
WIKIDATA = "https://query.wikidata.org/sparql"

COLUNAS = ["Ano", "Título no Brasil", "Título original", "Lançamento", "Direção",
           "Roteiro", "Ganhou Oscar?", "Oscars ganhos", "Sinopse"]
COLUNAS_FILME = ["Título no Brasil", "Título original", "Lançamento", "Direção", "Roteiro", "Sinopse"]

MESES = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho",
         "agosto", "setembro", "outubro", "novembro", "dezembro"]
DATA_PT = re.compile(r"(\d{1,2}) de (" + "|".join(MESES) + r") de (\d{4})", re.I)

# Categoria do Oscar (rótulo em inglês no Wikidata) -> nome usado na planilha.
# A ordem desta lista é a ordem em que as categorias aparecem na célula.
CATEGORIAS = [
    ("Best Picture", "Filme"), ("Best Director", "Direção"),
    ("Supporting Actor", "Ator coadjuvante"), ("Supporting Actress", "Atriz coadjuvante"),
    ("Best Actor", "Ator"), ("Best Actress", "Atriz"),
    ("Original Screenplay", "Roteiro original"), ("Adapted Screenplay", "Roteiro adaptado"),
    ("International Feature", "Filme internacional"), ("Foreign Language", "Filme estrangeiro"),
    ("Animated Feature", "Longa de animação"), ("Documentary Feature", "Documentário"),
    ("Casting", "Direção de elenco"), ("Cinematography", "Fotografia"),
    ("Film Editing", "Montagem"), ("Production Design", "Direção de arte"),
    ("Art Direction", "Direção de arte"), ("Costume", "Figurino"), ("Makeup", "Maquiagem"),
    ("Original Score", "Trilha sonora"), ("Original Song", "Canção original"),
    ("Sound Mixing", "Mixagem de som"), ("Sound Editing", "Edição de som"),
    ("Best Sound", "Som"), ("Visual Effects", "Efeitos visuais"),
    ("Documentary Short", "Curta documentário"), ("Animated Short", "Curta de animação"),
    ("Live Action Short", "Curta-metragem"),
]
ATUACAO = {"Ator", "Atriz", "Ator coadjuvante", "Atriz coadjuvante"}
ROTULOS_CREDITOS = {"Direção", "Roteiro", "Elenco", "Trilha Sonora", "Produção",
                    "Equipe técnica", "Distribuição", "Empresas envolvidas"}
FUNCOES_ROTEIRO = {"Roteirista", "Roteiro", "Autor da obra original", "Adaptado por",
                   "Autor do roteiro original", "Diálogos", "Colaborador do roteiro",
                   "Ideia original", "Autor", "Argumento"}

sessao = requests.Session()
sessao.headers["User-Agent"] = f"RoboPlanilhaFilmes/1.0 (uso pessoal; {CONTATO})"
erros = []


# ----------------------------- utilidades -----------------------------
def log(msg):
    print(msg, flush=True)


def normalizar(texto):
    texto = unicodedata.normalize("NFKD", str(texto or "")).encode("ascii", "ignore").decode()
    return re.sub(r"[^a-z0-9]", "", texto.lower())


def baixar(url, **kw):
    r = sessao.get(url, timeout=30, **kw)
    r.raise_for_status()
    time.sleep(PAUSA)
    return r


def resumir(texto, limite=220):
    texto = re.sub(r"\s+", " ", html.unescape(texto or "")).strip()
    frase = re.split(r"(?<=[.!?])\s", texto, maxsplit=1)[0]
    if len(frase) > limite:
        frase = frase[:limite].rsplit(" ", 1)[0].rstrip(",;:") + "…"
    return frase


def ano_da_linha(linha):
    m = re.search(r"\d{4}", linha.get("Ano", ""))
    return int(m.group()) if m else None


# ----------------------------- planilha -----------------------------
def ler_planilha():
    if not ARQUIVO.exists():
        log("ERRO: filmes.csv não encontrado. Envie a planilha exportada como CSV.")
        sys.exit(1)
    bruto_bytes = ARQUIVO.read_bytes()
    try:
        texto = bruto_bytes.decode("utf-8-sig")
    except UnicodeDecodeError:
        texto = bruto_bytes.decode("cp1252")   # CSV comum do Excel em português
    primeira = texto.split("\n", 1)[0]
    separador = ";" if primeira.count(";") > primeira.count(",") else ","
    leitor = csv.DictReader(texto.splitlines(), delimiter=separador)
    # aceita pequenas diferenças no nome das colunas
    mapa = {normalizar(c): c for c in COLUNAS}
    linhas = []
    for bruto in leitor:
        linha = {c: "" for c in COLUNAS}
        for chave, valor in bruto.items():
            if chave and normalizar(chave) in mapa:
                linha[mapa[normalizar(chave)]] = (valor or "").strip()
        if any(linha.values()):
            linhas.append(linha)
    return linhas


def salvar_planilha(linhas):
    linhas.sort(key=lambda l: ano_da_linha(l) or 9999)  # mantém a ordem dentro de cada ano
    with ARQUIVO.open("w", encoding="utf-8", newline="") as f:
        escritor = csv.DictWriter(f, fieldnames=COLUNAS, quoting=csv.QUOTE_MINIMAL)
        escritor.writeheader()
        escritor.writerows(linhas)


def achar_linha(linhas, titulo_original, ano):
    alvo = normalizar(titulo_original)
    for linha in linhas:
        a = ano_da_linha(linha)
        if normalizar(linha["Título original"]) == alvo and a and ano and abs(a - ano) <= 1:
            return linha
    return None


# ----------------------------- AdoroCinema -----------------------------
def listar_em_cartaz():
    soup = BeautifulSoup(baixar(LISTA_URL).text, "html.parser")
    links = soup.select("a.meta-title-link") or soup.find_all("a", href=True)
    ids = []
    for a in links:
        m = re.search(r"/filmes/filme-(\d+)/?$", a.get("href", ""))
        if m and m.group(1) not in ids:
            ids.append(m.group(1))
    return ids[:QUANTOS_POR_VEZ]


def ler_jsonld(soup):
    for tag in soup.find_all("script", type="application/ld+json"):
        try:
            dados = json.loads(tag.string or "", strict=False)
        except ValueError:
            continue
        itens = dados if isinstance(dados, list) else dados.get("@graph", [dados])
        for item in itens:
            if isinstance(item, dict) and item.get("@type") == "Movie":
                return item
    return {}


def nomes(valor):
    if isinstance(valor, dict):
        valor = [valor]
    if not isinstance(valor, list):
        return []
    return [html.unescape(v.get("name", "")).strip() for v in valor if isinstance(v, dict) and v.get("name")]


def linhas_de_texto(soup):
    return [l.strip() for l in soup.get_text("\n").split("\n") if l.strip()]


def valor_depois(linhas, rotulo):
    for i, l in enumerate(linhas):
        if l.startswith(rotulo):
            resto = l[len(rotulo):].strip(" :")
            if resto:
                return resto
            if i + 1 < len(linhas):
                return linhas[i + 1]
    return ""


def secao(linhas, titulo):
    """Linhas entre o título de uma seção da página de créditos e a próxima seção."""
    if titulo not in linhas:
        return []
    inicio = linhas.index(titulo) + 1
    saida = []
    for l in linhas[inicio:]:
        if l in ROTULOS_CREDITOS:
            break
        saida.append(l)
    return saida


def ler_creditos(id_filme):
    soup = BeautifulSoup(baixar(FILME_URL.format(id_filme) + "creditos/").text, "html.parser")
    linhas = linhas_de_texto(soup)
    direcao = [l for l in secao(linhas, "Direção") if not l.startswith("foto de")]

    itens = [l for l in secao(linhas, "Roteiro") if not l.startswith("foto de")]
    pares, funcao_primeiro = [], bool(itens) and itens[0] in FUNCOES_ROTEIRO
    pendente = None
    for item in itens:
        eh_funcao = item in FUNCOES_ROTEIRO
        if funcao_primeiro:
            if eh_funcao:
                pendente = item
            elif pendente:
                pares.append((pendente, item)); pendente = None
        else:
            if not eh_funcao:
                pendente = item
            elif pendente:
                pares.append((item, pendente)); pendente = None
    roteiristas = [n for f, n in pares if f in ("Roteirista", "Roteiro")] or [n for _, n in pares]
    return list(dict.fromkeys(direcao)), list(dict.fromkeys(roteiristas))


def ler_filme(id_filme):
    soup = BeautifulSoup(baixar(FILME_URL.format(id_filme)).text, "html.parser")
    ld = ler_jsonld(soup)
    linhas = linhas_de_texto(soup)

    titulo = html.unescape(ld.get("name", "")).strip()
    if not titulo:
        h = soup.select_one(".titlebar-title") or soup.find("h1")
        titulo = h.get_text(" ", strip=True) if h else ""
    original = valor_depois(linhas, "Título original") or titulo

    lancamento = ""
    if re.match(r"\d{4}-\d{2}-\d{2}", ld.get("datePublished", "")):
        a, m, d = ld["datePublished"][:10].split("-")
        lancamento = f"{d}/{m}/{a}"
    else:
        area = soup.select_one(".meta-body-info")
        achado = DATA_PT.search(area.get_text(" ") if area else " ".join(linhas))
        if achado:
            d, mes, a = achado.groups()
            lancamento = f"{int(d):02d}/{MESES.index(mes.lower()) + 1:02d}/{a}"

    sinopse = ld.get("description", "")
    if not sinopse:
        bloco = soup.select_one(".content-txt")
        sinopse = bloco.get_text(" ", strip=True) if bloco else ""

    direcao_cred, roteiro = ler_creditos(id_filme)
    direcao = nomes(ld.get("director")) or direcao_cred

    return {
        "Ano": lancamento[-4:] if lancamento else str(date.today().year),
        "Título no Brasil": titulo,
        "Título original": original,
        "Lançamento": lancamento,
        "Direção": ", ".join(direcao),
        "Roteiro": ", ".join(roteiro),
        "Sinopse": resumir(sinopse),
    }


# ----------------------------- Wikidata (Oscar) -----------------------------
def consultar(sparql):
    r = baixar(WIKIDATA, params={"query": sparql, "format": "json"})
    return r.json()["results"]["bindings"]


def achar_no_wikidata(titulos, ano):
    melhor = None
    for titulo in dict.fromkeys(t for t in titulos if t):
        busca = titulo.replace("\\", "").replace('"', '\\"')
        for idioma in ("en", "pt"):
            sparql = f"""
            SELECT DISTINCT ?f ?d WHERE {{
              SERVICE wikibase:mwapi {{
                bd:serviceParam wikibase:endpoint "www.wikidata.org";
                                wikibase:api "EntitySearch";
                                mwapi:search "{busca}"; mwapi:language "{idioma}".
                ?f wikibase:apiOutputItem mwapi:item.
              }}
              ?f wdt:P31/wdt:P279* wd:Q11424; wdt:P577 ?d.
            }}"""
            for r in consultar(sparql):
                distancia = abs(int(r["d"]["value"][:4]) - ano)
                if distancia <= 1 and (melhor is None or distancia < melhor[1]):
                    melhor = (r["f"]["value"].rsplit("/", 1)[1], distancia)
            if melhor and melhor[1] == 0:
                return melhor[0]
    return melhor[0] if melhor else None


def oscars_do_filme(qid):
    sparql = f"""
    SELECT ?premioLabel ?pessoaLabel WHERE {{
      {{ wd:{qid} p:P166 ?st. ?st ps:P166 ?premio. OPTIONAL {{ ?st pq:P1346 ?pessoa }} }}
      UNION
      {{ ?pessoa p:P166 ?st. ?st ps:P166 ?premio; pq:P1686 wd:{qid}. }}
      SERVICE wikibase:label {{ bd:serviceParam wikibase:language "en". }}
    }}"""
    ganhos = {}
    for r in consultar(sparql):
        premio = r["premioLabel"]["value"]
        if not premio.startswith("Academy Award for"):
            continue
        categoria = next((pt for chave, pt in CATEGORIAS if chave in premio),
                         premio.replace("Academy Award for ", ""))
        pessoas = ganhos.setdefault(categoria, [])
        pessoa = r.get("pessoaLabel", {}).get("value", "")
        if categoria in ATUACAO and pessoa and not re.match(r"Q\d+$", pessoa) and pessoa not in pessoas:
            pessoas.append(pessoa)
    ordem = [pt for _, pt in CATEGORIAS]
    textos = []
    for cat in sorted(ganhos, key=lambda c: ordem.index(c) if c in ordem else 99):
        textos.append(f"{cat} ({', '.join(ganhos[cat])})" if ganhos[cat] else cat)
    return textos


def atualizar_oscar(linha, hoje):
    ano = ano_da_linha(linha)
    if not ano:
        return False
    cerimonia = ano + 1
    if hoje < date(cerimonia, 3, 1):          # premiação ainda não aconteceu
        if not linha["Ganhou Oscar?"]:
            linha["Ganhou Oscar?"], linha["Oscars ganhos"] = "Aguardando", f"Premiação em {cerimonia}"
            return True
        return False

    qid = achar_no_wikidata([linha["Título original"], linha["Título no Brasil"]], ano)
    if not qid:
        log(f"  ! Não achei '{linha['Título no Brasil']}' no Wikidata; Oscar fica como está.")
        return False
    ganhos = oscars_do_filme(qid)
    if ganhos:
        linha["Ganhou Oscar?"] = "Sim"
        linha["Oscars ganhos"] = f"{len(ganhos)} — {', '.join(ganhos)}"
        return True
    if hoje >= date(cerimonia, 5, 15):        # dá tempo de o Wikidata registrar os vencedores
        linha["Ganhou Oscar?"], linha["Oscars ganhos"] = "Não", "—"
        return True
    return False


# ----------------------------- principal -----------------------------
def main():
    hoje = date.today()
    linhas = ler_planilha()
    controle = json.loads(CONTROLE.read_text(encoding="utf-8")) if CONTROLE.exists() else {}
    mudou = False
    log(f"Planilha com {len(linhas)} filmes.")

    # 1) Filmes em cartaz no AdoroCinema
    try:
        ids = listar_em_cartaz()
        log(f"{len(ids)} filmes em cartaz encontrados.")
    except Exception as e:
        ids = []
        erros.append(f"Não consegui ler a lista de filmes em cartaz: {e}")

    for id_filme in ids:
        try:
            visto = controle.get(id_filme)
            linha = achar_linha(linhas, visto["original"], visto["ano"]) if visto else None
            if linha and all(linha[c] for c in COLUNAS_FILME):
                continue                              # já está completo, nem visita o site
            dados = ler_filme(id_filme)
            ano = int(dados["Ano"])
            linha = linha or achar_linha(linhas, dados["Título original"], ano)
            if linha:
                for c in COLUNAS_FILME:
                    if not linha[c] and dados[c]:
                        linha[c] = dados[c]; mudou = True
                log(f"  = {linha['Título no Brasil']} (completado)")
            else:
                linha = {c: "" for c in COLUNAS}
                linha.update(dados)
                linhas.append(linha); mudou = True
                log(f"  + {dados['Título no Brasil']} (novo)")
            controle[id_filme] = {"original": linha["Título original"], "ano": ano_da_linha(linha)}
        except Exception as e:
            erros.append(f"Filme {FILME_URL.format(id_filme)}: {e}")

    # 2) Oscar das linhas pendentes
    for linha in linhas:
        if linha["Ganhou Oscar?"] in ("", "Aguardando"):
            try:
                if atualizar_oscar(linha, hoje):
                    mudou = True
                    log(f"  * Oscar: {linha['Título no Brasil']} -> {linha['Ganhou Oscar?']}")
            except Exception as e:
                erros.append(f"Oscar de '{linha['Título no Brasil']}': {e}")

    salvar_planilha(linhas)   # sempre regrava, para padronizar o formato do arquivo
    CONTROLE.write_text(json.dumps(controle, ensure_ascii=False, indent=1), encoding="utf-8")
    log("Planilha atualizada." if mudou else "Nada novo hoje.")

    if erros:
        log("\nPROBLEMAS ENCONTRADOS:")
        for e in erros:
            log(f"  - {e}")
        sys.exit(1)   # faz o GitHub avisar você por e-mail


if __name__ == "__main__":
    main()
