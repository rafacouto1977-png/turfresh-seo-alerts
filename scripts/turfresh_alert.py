"""
TurFresh - Alerta Semanal de GSC
==================================
Motor de comparacao contra media movel de 4 semanas, com piso de volume
em CADA semana (nao so na media) para evitar ruido de numeros pequenos.

Por que piso por semana e nao so na media: se 3 das 4 semanas anteriores
tiveram 20 impressoes e uma teve 1.200 (um pico isolado, ex. trafego de
bot ou evento sazonal), a MEDIA pode passar de 300 mesmo que a pagina seja
pequena no dia a dia. Exigir volume minimo em pelo menos 3 das 4 semanas
individualmente evita que um outlier valide um piso que nao existe de verdade.

Este primeiro corte tem so o Gatilho 1 (vazamento de CTR) funcionando.
Os outros 5 entram depois, um de cada vez, em cima desse mesmo motor.
"""

import os
import re
from collections import defaultdict
from datetime import date, timedelta

import pandas as pd
from google.oauth2 import service_account
from googleapiclient.discovery import build
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ===========================================================================
# CONFIG
# ===========================================================================
SITE_URL = os.environ.get("SITE_URL", "https://turfresh.com/")
GSC_CLIENT_EMAIL = os.environ.get("GSC_CLIENT_EMAIL")
GSC_PRIVATE_KEY = os.environ.get("GSC_PRIVATE_KEY")

WEEK_LAG_DAYS = 3          # GSC leva ~3 dias para fechar os dados de uma semana
N_TRAILING_WEEKS = 4

# --- Piso por semana individual (refinamento sobre o design original) ---
# Uma semana so entra na media se tiver pelo menos este volume. Precisa de
# pelo menos MIN_WEEKS_WITH_FLOOR das 4 semanas validas para a media contar.
WEEK_FLOOR_IMPRESSIONS = 30
MIN_WEEKS_WITH_FLOOR = 3

# --- Gatilho 1: vazamento de CTR ---
G1_MIN_IMPRESSIONS = 150   # PROVISORIO: o piso original de 500 nunca foi
# atingido por nenhuma query comercial na primeira execucao real (0 de 1642
# passaram, confirmado pelo diagnostico de funil). 150 e um chute mais
# conservador para nao ficar zerado de novo - a proxima execucao imprime a
# distribuicao real (mediana, p75, p90, p95, p99) para calibrar com dado
# em vez de outro chute.
G1_MAX_POSITION = 8
G1_CTR_ABSOLUTE_FLOOR = 0.01     # 1%
G1_CTR_DROP_RATIO = 0.40         # queda de 40% vs media

# --- Gatilho 2: queda de impressoes precoce ---
G2_TOP_N_PAGES = 20
G2_DROP_RATIO = 0.30
G2_MIN_AVG_IMPRESSIONS = 300
G2_MAX_POSITION_VARIATION = 2.0

# --- Gatilho 3: query nova emergindo ---
G3_MIN_IMPRESSIONS = 100
G3_MIN_IMPRESSIONS_QUESTION = 50
QUESTION_PATTERNS = [r"^how\b", r"^what\b", r"^why\b", r"^is\b", r"^can\b", r"^does\b",
                    r"^do\b", r"^are\b", r"^will\b", r"^should\b"]

# --- Gatilho 4: decaimento dos posts otimizados ---
G4_DROP_RATIO = 0.30
G4_MIN_AVG_CLICKS = 20
SEO_LOG_PATH = "data/seo_log_urls.csv"

# --- Gatilho 5: sinal de vida do conteudo novo ---
G5_POSITIVE_MIN_IMPRESSIONS = 50
G5_NEGATIVE_MIN_DAYS = 21
CITY_PAGES_PATH = "data/city_pages.csv"

# --- Gatilho 6: city pages radar ---
G6_NEW_PAGE_THRESHOLD = 20         # abaixo disso, pagina e tratada como nova/pequena
G6_RISE_FIRST_SIGNAL = 20         # primeira vez que bate isso = "comecou a rankear"
G6_RISE_RATIO = 0.25              # pagina estabelecida: sobe 25%+
G6_FALL_RATIO = 0.25
G6_FALL_MIN_AVG = 20
CITY_PAGE_URL_PATTERN = r"^/(?:arizona|california|nevada|florida|texas)/[a-z-]+/$"

# --- Gatilho 7: trafego de marca ---
BRAND_PATTERNS = [
    r"\bturfresh\b", r"\bturf\s*fresh\b", r"\btur\s*fresh\b", r"\bturffresh\b",
]
G7_MIN_AVG_IMPRESSIONS = 20
G7_DROP_RATIO = 0.30
G7_POSITION_DROP_ALERT = 4.0
G7_MAX_MEANINGFUL_POSITION = 20   # abaixo do top 20 a posicao e ruido, nao alerta

# ===========================================================================
# CLASSIFICADOR DE QUERY COMERCIAL/LOCAL
# RASCUNHO - Rafael precisa revisar e calibrar esta lista antes de confiar
# no Gatilho 1 de verdade. Baseado no que ja sei do ICP da TurFresh (servico
# local de limpeza de grama sintetica), nao em dado medido do site.
# ===========================================================================
COMMERCIAL_LOCAL_PATTERNS = [
    r"\bturf cleaning\b", r"\bartificial (?:turf|grass) cleaning\b",
    r"\bcleaning service\b", r"\bnear me\b", r"\bcost\b", r"\bprice\b",
    r"\bpricing\b", r"\bquote\b", r"\bhire\b", r"\bcompany\b",
    r"\bprofessional\b", r"\bbest.*service\b", r"\bhow much\b",
    # cidades prioritarias (mesma logica das 35 city pages)
    r"\bphoenix\b", r"\bpeoria\b", r"\bsanta ana\b", r"\bsan jose\b",
    r"\bthousand oaks\b", r"\bsan clemente\b", r"\bsacramento\b",
    r"\bwhittier\b", r"\bsanta clarita\b", r"\bfresno\b", r"\blos angeles\b",
    r"\briverside\b", r"\bmiami\b", r"\bjacksonville\b", r"\bfort lauderdale\b",
    r"\btampa\b", r"\bhouston\b", r"\baustin\b", r"\blas vegas\b",
    r"\birving\b", r"\bsan diego\b", r"\bmesa\b", r"\bgilbert\b",
    r"\bchandler\b", r"\btempe\b", r"\bglendale\b", r"\bnorth las vegas\b",
    r"\balhambra\b",
]
# Sinal negativo: quem pesquisa "como fazer sozinho" nao vai contratar.
DIY_PATTERNS = [
    r"\bhow to (?:clean|remove|diy)\b", r"\bmyself\b", r"\bdiy\b",
    r"\bhomemade\b", r"\bat home\b",
]


def is_commercial_local(query):
    q = query.lower()
    if any(re.search(p, q) for p in DIY_PATTERNS):
        return False
    return any(re.search(p, q) for p in COMMERCIAL_LOCAL_PATTERNS)


def is_brand(query):
    q = query.lower()
    return any(re.search(p, q) for p in BRAND_PATTERNS)


def is_question(query):
    q = query.lower().strip()
    return any(re.search(p, q) for p in QUESTION_PATTERNS)


def norm_path(url):
    p = re.sub(r"^https?://[^/]+", "", str(url)).split("?")[0].split("#")[0]
    if not p:
        p = "/"
    if len(p) > 1 and not p.endswith("/"):
        p += "/"
    return p.lower()


def load_seo_log_urls():
    """As 69 URLs vivas do SEO Log (as 44 redirecionadas ja foram excluidas
    na extracao). Retorna dict path -> data_otimizacao (string, pode ser vazia)."""
    if not os.path.exists(SEO_LOG_PATH):
        print(f"  Aviso: {SEO_LOG_PATH} nao encontrado. Gatilho 4 fica vazio.")
        return {}
    out = {}
    with open(SEO_LOG_PATH, newline="", encoding="utf-8") as f:
        import csv
        for row in csv.DictReader(f):
            path = norm_path(row["path"])
            out[path] = row.get("data_otimizacao", "")
    print(f"  SEO Log: {len(out)} URLs vivas carregadas")
    return out


def load_city_pages():
    """As 35 city pages (21 prioritarias + 14 novas). Retorna dict
    path -> {categoria, data_publicacao}."""
    if not os.path.exists(CITY_PAGES_PATH):
        print(f"  Aviso: {CITY_PAGES_PATH} nao encontrado. Gatilhos 5/6 ficam vazios.")
        return {}
    out = {}
    with open(CITY_PAGES_PATH, newline="", encoding="utf-8") as f:
        import csv
        for row in csv.DictReader(f):
            path = norm_path(row["path"])
            out[path] = {"categoria": row.get("categoria", ""),
                        "data_publicacao": row.get("data_publicacao", "")}
    print(f"  City pages: {len(out)} paginas carregadas")
    return out


# ===========================================================================
# GSC FETCH
# ===========================================================================
def get_gsc_service():
    if not GSC_CLIENT_EMAIL or not GSC_PRIVATE_KEY:
        raise RuntimeError("GSC_CLIENT_EMAIL ou GSC_PRIVATE_KEY nao definidos.")
    info = {"type": "service_account", "client_email": GSC_CLIENT_EMAIL,
            "private_key": GSC_PRIVATE_KEY,
            "token_uri": "https://oauth2.googleapis.com/token"}
    creds = service_account.Credentials.from_service_account_info(
        info, scopes=["https://www.googleapis.com/auth/webmasters.readonly"])
    return build("searchconsole", "v1", credentials=creds)


def is_garbage_query(query):
    """
    Queries que carregam uma URL dentro do texto nao representam busca
    humana real - sao ruido de bot, scraper, ou alguem colando conteudo
    errado na caixa de busca do Google. Isso importa porque um \\b (limite
    de palavra) no regex de marca trata '/' e '.' como limite, entao
    'https://turfresh.com/...' bate em '\\bturfresh\\b' mesmo nao sendo uma
    busca de marca de verdade - foi assim que uma query lixo contaminou o
    Gatilho 7. Filtrar aqui, uma vez, antes de qualquer classificador rodar,
    em vez de remendar cada regex separadamente.
    """
    q = str(query)
    if re.search(r"https?://|www\.", q, re.IGNORECASE):
        return True
    if len(q) > 100:   # buscas reais raramente passam disso
        return True
    return False


def filter_garbage_queries(df):
    if df.empty:
        return df
    mask = ~df["query"].apply(is_garbage_query)
    removed = (~mask).sum()
    if removed:
        print(f"  ({removed} queries-lixo removidas - continham URL ou eram anormalmente longas)")
    return df[mask].reset_index(drop=True)


def fetch_week(service, start_date, end_date):
    """Uma semana de dados query+page. Pagina se precisar."""
    rows_out, start_row = [], 0
    while True:
        req = {"startDate": start_date.isoformat(), "endDate": end_date.isoformat(),
               "dimensions": ["query", "page"], "rowLimit": 25000, "startRow": start_row}
        rows = service.searchanalytics().query(siteUrl=SITE_URL, body=req).execute().get("rows", [])
        if not rows:
            break
        for r in rows:
            rows_out.append({"query": r["keys"][0], "page": r["keys"][1],
                             "clicks": r["clicks"], "impressions": r["impressions"],
                             "position": r["position"]})
        if len(rows) < 25000:
            break
        start_row += 25000
    return filter_garbage_queries(pd.DataFrame(rows_out))


def fetch_trailing_weeks(service, n_weeks=N_TRAILING_WEEKS):
    """
    Retorna (semana_atual_df, [lista de N dataframes das semanas anteriores,
    mais antiga primeiro], janelas usadas).
    """
    today = date.today()
    cur_end = today - timedelta(days=WEEK_LAG_DAYS)
    cur_start = cur_end - timedelta(days=6)

    print(f"Semana atual: {cur_start} a {cur_end}")
    current = fetch_week(service, cur_start, cur_end)
    print(f"  {len(current)} linhas query+pagina\n")

    trailing = []
    windows = [(cur_start, cur_end)]
    week_start = cur_start
    for i in range(n_weeks):
        week_end = week_start - timedelta(days=1)
        week_start = week_end - timedelta(days=6)
        print(f"Semana -{i+1}: {week_start} a {week_end}")
        df = fetch_week(service, week_start, week_end)
        print(f"  {len(df)} linhas query+pagina\n")
        trailing.insert(0, df)   # mais antiga primeiro
        windows.insert(0, (week_start, week_end))

    return current, trailing, windows


# ===========================================================================
# MEDIA MOVEL COM PISO POR SEMANA
# ===========================================================================
def build_trailing_stats(trailing_weeks):
    """
    Para cada (query, page), calcula media de impressoes/clicks das semanas
    validas, exige piso MIN_WEEKS_WITH_FLOOR de MIN_WEEKS_WITH_FLOOR semanas
    com volume, e devolve tambem a posicao media (para o teste de estabilidade
    de posicao) e se a chave apareceu em QUALQUER semana anterior (Gatilho 3).

    Retorna dict: (query, page) -> {
        media_impr, media_clicks, media_position, semanas_com_piso,
        confiavel (bool), esteve_presente (bool)
    }
    """
    per_key = defaultdict(lambda: {"impr": [], "clicks": [], "pos": []})

    for week_df in trailing_weeks:
        if week_df.empty:
            continue
        for _, r in week_df.iterrows():
            k = (r["query"], r["page"])
            per_key[k]["impr"].append(float(r["impressions"]))
            per_key[k]["clicks"].append(float(r["clicks"]))
            per_key[k]["pos"].append(float(r["position"]))

    stats = {}
    for k, v in per_key.items():
        n_semanas_presentes = len(v["impr"])
        # semanas em que a chave nao apareceu contam como 0 impressoes/clicks
        impr_padded = v["impr"] + [0.0] * (N_TRAILING_WEEKS - n_semanas_presentes)
        clicks_padded = v["clicks"] + [0.0] * (N_TRAILING_WEEKS - n_semanas_presentes)

        semanas_com_piso = sum(1 for x in impr_padded if x >= WEEK_FLOOR_IMPRESSIONS)
        confiavel = semanas_com_piso >= MIN_WEEKS_WITH_FLOOR

        stats[k] = {
            "media_impr": sum(impr_padded) / N_TRAILING_WEEKS,
            "media_clicks": sum(clicks_padded) / N_TRAILING_WEEKS,
            "media_position": sum(v["pos"]) / len(v["pos"]) if v["pos"] else None,
            "semanas_com_piso": semanas_com_piso,
            "confiavel": confiavel,
            "esteve_presente": n_semanas_presentes > 0,
        }
    return stats


# ===========================================================================
# GATILHO 1 - VAZAMENTO DE CTR
# ===========================================================================
def gatilho_1_vazamento_ctr(current_df, trailing_stats):
    """
    Dispara quando: impressoes_semana >= 500 E posicao <= 8 E
    (CTR < 1% OU CTR caiu >= 40% vs media_4sem)

    Refinamento aplicado: a condicao de "CTR caiu vs media" so e avaliada se
    a posicao media das 4 semanas anteriores TAMBEM estava <= 8 (com folga de
    +2). Sem isso, uma pagina que acabou de subir para o top 8 essa semana
    teria uma "media" de CTR baixa so porque antes ela rankeava pior - isso
    pareceria vazamento sem ser.
    """
    alerts = []
    if current_df.empty:
        return alerts

    for _, r in current_df.iterrows():
        query, page = r["query"], r["page"]
        if not is_commercial_local(query):
            continue

        impr = float(r["impressions"])
        pos = float(r["position"])
        clicks = float(r["clicks"])

        if impr < G1_MIN_IMPRESSIONS or pos > G1_MAX_POSITION:
            continue

        ctr = clicks / impr if impr else 0
        k = (query, page)
        st = trailing_stats.get(k)

        by_absolute = ctr < G1_CTR_ABSOLUTE_FLOOR

        by_drop = False
        drop_pct = None
        if st and st["confiavel"] and st["media_clicks"] > 0 and st["media_impr"] > 0:
            media_ctr = st["media_clicks"] / st["media_impr"]
            # so avalia queda se a posicao ja estava boa nas semanas anteriores
            posicao_ja_estavel = (st["media_position"] is not None
                                  and st["media_position"] <= G1_MAX_POSITION + 2)
            if media_ctr > 0 and posicao_ja_estavel:
                drop_pct = (media_ctr - ctr) / media_ctr
                by_drop = drop_pct >= G1_CTR_DROP_RATIO

        if not (by_absolute or by_drop):
            continue

        if by_absolute and by_drop:
            motivo = f"CTR absoluto de {ctr*100:.2f}% (abaixo do piso de 1%) E caiu {drop_pct*100:.0f}% vs media"
        elif by_absolute:
            motivo = f"CTR de {ctr*100:.2f}%, abaixo do piso absoluto de 1% para posicao {pos:.1f}"
        else:
            motivo = f"CTR caiu {drop_pct*100:.0f}% vs a media das 4 semanas anteriores"

        confianca = "ALTA" if (st and st["confiavel"]) else "BAIXA (pouco historico)"

        alerts.append({
            "gatilho": "1. Vazamento de CTR",
            "query": query,
            "pagina": page,
            "posicao": round(pos, 1),
            "impressoes_semana": int(impr),
            "clicks_semana": int(clicks),
            "ctr_semana": f"{ctr*100:.2f}%",
            "media_ctr_4sem": (f"{(st['media_clicks']/st['media_impr'])*100:.2f}%"
                               if st and st["media_impr"] > 0 else "sem historico"),
            "motivo": motivo,
            "confianca": confianca,
        })

    alerts.sort(key=lambda x: -x["impressoes_semana"])
    return alerts


# ===========================================================================
# GATILHO 2 - QUEDA DE IMPRESSOES PRECOCE
# ===========================================================================
def gatilho_2_queda_precoce(current_df, trailing_stats):
    """
    Top 20 paginas + queries comerciais. Dispara quando impressoes cairam
    >=30% vs media_4sem (media confiavel, piso 300/sem) E posicao ficou
    estavel (variacao < 2). Posicao estavel + impressao caindo = Google
    mostrando menos, nao voce piorando (AI Overview, perda de feature).
    """
    alerts = []
    if current_df.empty:
        return alerts

    page_impr = current_df.groupby("page")["impressions"].sum().sort_values(ascending=False)
    top_pages = set(page_impr.head(G2_TOP_N_PAGES).index)
    money_queries = set(current_df[current_df["query"].apply(is_commercial_local)]["query"])

    for _, r in current_df.iterrows():
        query, page = r["query"], r["page"]
        if page not in top_pages and query not in money_queries:
            continue

        k = (query, page)
        st = trailing_stats.get(k)
        if not st or not st["confiavel"] or st["media_impr"] < G2_MIN_AVG_IMPRESSIONS:
            continue

        impr = float(r["impressions"])
        pos = float(r["position"])
        drop = (st["media_impr"] - impr) / st["media_impr"]
        if drop < G2_DROP_RATIO:
            continue

        pos_variation = abs(pos - st["media_position"]) if st["media_position"] else 999
        if pos_variation >= G2_MAX_POSITION_VARIATION:
            continue   # posicao tambem mudou - nao e o padrao "Google mostrando menos"

        alerts.append({
            "gatilho": "2. Queda de impressoes precoce",
            "query": query, "pagina": page, "posicao": round(pos, 1),
            "posicao_media_4sem": round(st["media_position"], 1),
            "impressoes_semana": int(impr),
            "media_impr_4sem": round(st["media_impr"], 0),
            "queda_pct": f"{drop*100:.0f}%",
            "motivo": (f"Impressoes cairam {drop*100:.0f}% vs media (posicao ficou "
                      f"estavel: {pos:.1f} vs media {st['media_position']:.1f}). "
                      f"Google esta mostrando menos essa pagina/query, nao um "
                      f"problema de ranking."),
            "confianca": "ALTA",
        })

    alerts.sort(key=lambda x: -x["impressoes_semana"])
    return alerts


# ===========================================================================
# GATILHO 3 - QUERY NOVA EMERGINDO
# ===========================================================================
def gatilho_3_query_nova(current_df, trailing_stats):
    """
    Query ausente nas 4 semanas anteriores E impressoes_semana >= 100.
    Piso cai para 50 se for pergunta (candidata a calendario de AI).
    """
    alerts = []
    if current_df.empty:
        return alerts

    seen_queries = current_df.groupby("query")["impressions"].sum()
    for query, impr in seen_queries.items():
        # ja apareceu em alguma semana anterior para QUALQUER pagina?
        already_seen = any(st["esteve_presente"] for k, st in trailing_stats.items()
                           if k[0] == query)
        if already_seen:
            continue

        pergunta = is_question(query)
        piso = G3_MIN_IMPRESSIONS_QUESTION if pergunta else G3_MIN_IMPRESSIONS
        if impr < piso:
            continue

        page = current_df[current_df["query"] == query].iloc[0]["page"]
        pos = current_df[current_df["query"] == query].iloc[0]["position"]

        alerts.append({
            "gatilho": "3. Query nova emergindo",
            "query": query, "pagina": page, "posicao": round(float(pos), 1),
            "impressoes_semana": int(impr),
            "tipo": "PERGUNTA - candidata a calendario AI" if pergunta else "comum",
            "motivo": (f"Query nao existia nas 4 semanas anteriores. "
                      f"{'E uma pergunta, boa candidata a conteudo de AI Overview.' if pergunta else ''}"),
            "confianca": "ALTA" if impr >= 200 else "MEDIA",
        })

    alerts.sort(key=lambda x: -x["impressoes_semana"])
    return alerts


# ===========================================================================
# GATILHO 4 - DECAIMENTO DOS POSTS OTIMIZADOS
# ===========================================================================
def gatilho_4_decaimento_posts(current_df, trailing_stats, seo_log_urls):
    """
    Escopo: as 69 URLs vivas do SEO Log (as 44 redirecionadas ja foram
    excluidas na extracao dos dados). Clicks cairam >=30% vs media_4sem,
    com piso de 20 clicks/sem para nao alertar ruido de posts pequenos.
    """
    alerts = []
    if current_df.empty or not seo_log_urls:
        return alerts

    current_df = current_df.copy()
    current_df["path"] = current_df["page"].apply(norm_path)
    scoped = current_df[current_df["path"].isin(seo_log_urls.keys())]

    page_clicks = scoped.groupby("path")["clicks"].sum()
    page_impr = scoped.groupby("path")["impressions"].sum()
    page_pos = scoped.groupby("path")["position"].mean()

    # agrega trailing_stats por pagina (nao por query+pagina)
    page_trailing = defaultdict(lambda: {"clicks": [0.0]*N_TRAILING_WEEKS})
    for (query, page), st in trailing_stats.items():
        path = norm_path(page)
        if path in seo_log_urls:
            # aproximacao: soma media_clicks de todas as queries dessa pagina
            page_trailing[path]["clicks"] = [
                page_trailing[path]["clicks"][0] + st["media_clicks"]]

    for path in page_clicks.index:
        media_clicks = page_trailing.get(path, {}).get("clicks", [0])[0]
        if media_clicks < G4_MIN_AVG_CLICKS:
            continue
        clicks_now = float(page_clicks[path])
        drop = (media_clicks - clicks_now) / media_clicks if media_clicks else 0
        if drop < G4_DROP_RATIO:
            continue

        alerts.append({
            "gatilho": "4. Decaimento posts otimizados",
            "pagina": path,
            "clicks_semana": int(clicks_now),
            "media_clicks_4sem": round(media_clicks, 1),
            "impressoes_semana": int(page_impr.get(path, 0)),
            "posicao": round(float(page_pos.get(path, 0)), 1),
            "queda_pct": f"{drop*100:.0f}%",
            "data_otimizacao": seo_log_urls.get(path, ""),
            "motivo": f"Clicks cairam {drop*100:.0f}% vs a media das 4 semanas anteriores.",
            "confianca": "ALTA" if media_clicks >= 40 else "MEDIA",
        })

    alerts.sort(key=lambda x: -x["media_clicks_4sem"])
    return alerts


# ===========================================================================
# GATILHO 5 - SINAL DE VIDA DO CONTEUDO NOVO
# ===========================================================================
def gatilho_5_sinal_vida(current_df, city_pages, run_date):
    """
    Escopo: city pages com data de publicacao conhecida, publicadas nos
    ultimos ~35 dias (folga sobre os 30 para nao perder o corte por causa
    do lag de dados do GSC).
    Positivo: primeira semana com impressoes >= 50.
    Negativo: publicada ha >= 21 dias E impressoes = 0 (problema de indexacao).
    """
    alerts = []
    if current_df.empty or not city_pages:
        return alerts

    current_df = current_df.copy()
    current_df["path"] = current_df["page"].apply(norm_path)
    page_impr = current_df.groupby("path")["impressions"].sum()

    for path, meta in city_pages.items():
        data_pub_str = meta.get("data_publicacao", "")
        if not data_pub_str or "verificar" in data_pub_str.lower():
            continue   # sem data confiavel, nao da para avaliar "sinal de vida"

        dt = _parse_date_flex(data_pub_str)
        if dt is None:
            continue

        dias_desde_publicacao = (run_date - dt).days
        if dias_desde_publicacao < 0 or dias_desde_publicacao > 35:
            continue

        impr = float(page_impr.get(path, 0))

        if dias_desde_publicacao <= 7 and impr >= G5_POSITIVE_MIN_IMPRESSIONS:
            alerts.append({
                "gatilho": "5. Sinal de vida (positivo)",
                "pagina": path, "dias_desde_publicacao": dias_desde_publicacao,
                "impressoes_semana": int(impr),
                "motivo": f"Primeira semana com {int(impr)} impressoes - Google comecou a testar a pagina.",
                "confianca": "ALTA",
            })
        elif dias_desde_publicacao >= G5_NEGATIVE_MIN_DAYS and impr == 0:
            alerts.append({
                "gatilho": "5. Sinal de vida (negativo)",
                "pagina": path, "dias_desde_publicacao": dias_desde_publicacao,
                "impressoes_semana": 0,
                "motivo": (f"Publicada ha {dias_desde_publicacao} dias, ZERO impressoes. "
                          f"Possivel problema de indexacao - checar no GSC (Inspecao de URL)."),
                "confianca": "ALTA",
            })

    return alerts


def _parse_date_flex(s):
    """Datas no SEO Log vem em formatos inconsistentes (07 Mai, 10 Jul 2026,
    Verificar). Tenta alguns formatos comuns em portugues; devolve None se
    nao conseguir, para o chamador decidir pular em vez de quebrar."""
    s = str(s).strip()
    meses = {"jan":1,"fev":2,"mar":3,"abr":4,"mai":5,"jun":6,"jul":7,"ago":8,
              "set":9,"out":10,"nov":11,"dez":12}
    m = re.match(r"(\d{1,2})\s+([a-zA-Z]{3})\s*(\d{4})?", s)
    if not m:
        return None
    dia, mes_str, ano = m.groups()
    mes = meses.get(mes_str.lower()[:3])
    if not mes:
        return None
    ano = int(ano) if ano else date.today().year
    try:
        return date(ano, mes, int(dia))
    except ValueError:
        return None


# ===========================================================================
# GATILHO 6 - CITY PAGES (SUBINDO / CAINDO)
# ===========================================================================
def gatilho_6_city_pages(current_df, trailing_stats, city_pages):
    """
    Escopo: as 35 city pages cadastradas + qualquer URL que bata no padrao
    /{estado}/{cidade}/ (deteccao automatica, cresce sozinha com paginas novas).
    """
    alerts = []
    if current_df.empty:
        return alerts

    current_df = current_df.copy()
    current_df["path"] = current_df["page"].apply(norm_path)

    is_city = (current_df["path"].isin(city_pages.keys())
              | current_df["path"].str.match(CITY_PAGE_URL_PATTERN))
    scoped = current_df[is_city]

    page_impr_now = scoped.groupby("path")["impressions"].sum()
    page_clicks_now = scoped.groupby("path")["clicks"].sum()

    page_trailing = defaultdict(lambda: [0.0] * N_TRAILING_WEEKS)
    for (query, page), st in trailing_stats.items():
        path = norm_path(page)
        if path in page_impr_now.index:
            # aproxima somando as medias de impressao de todas as queries da pagina
            pass  # tratado abaixo de forma agregada

    # media por pagina: soma media_impr de todas as (query,page) daquela pagina
    page_media = defaultdict(float)
    for (query, page), st in trailing_stats.items():
        path = norm_path(page)
        if path in page_impr_now.index:
            page_media[path] += st["media_impr"]

    for path in page_impr_now.index:
        impr_now = float(page_impr_now[path])
        media = page_media.get(path, 0.0)

        if media < G6_NEW_PAGE_THRESHOLD:
            if impr_now >= G6_RISE_FIRST_SIGNAL:
                alerts.append({
                    "gatilho": "6. City page subindo",
                    "pagina": path, "impressoes_semana": int(impr_now),
                    "media_impr_4sem": round(media, 0),
                    "motivo": f"Pagina nova/pequena comecou a rankear: {int(impr_now)} impressoes essa semana.",
                    "confianca": "MEDIA",
                })
        else:
            variacao = (impr_now - media) / media
            if variacao >= G6_RISE_RATIO:
                alerts.append({
                    "gatilho": "6. City page subindo",
                    "pagina": path, "impressoes_semana": int(impr_now),
                    "media_impr_4sem": round(media, 0),
                    "motivo": f"Impressoes subiram {variacao*100:.0f}% vs media de 4 semanas.",
                    "confianca": "ALTA",
                })
            elif -variacao >= G6_FALL_RATIO and media >= G6_FALL_MIN_AVG:
                alerts.append({
                    "gatilho": "6. City page caindo",
                    "pagina": path, "impressoes_semana": int(impr_now),
                    "media_impr_4sem": round(media, 0),
                    "motivo": f"Impressoes cairam {-variacao*100:.0f}% vs media de 4 semanas.",
                    "confianca": "ALTA",
                })

    alerts.sort(key=lambda x: -x["impressoes_semana"])
    return alerts


# ===========================================================================
# GATILHO 7 - TRAFEGO DE MARCA
# ===========================================================================
def gatilho_7_marca(current_df, trailing_stats):
    """
    Duas formas de disparar, porque marca tem um risco que query comum nao
    tem - alguem pode superar voce no seu proprio nome:
      A) impressoes/clicks caindo com posicao estavel (Google mostrando
         menos a marca - AI Overview, Knowledge Panel, etc.)
      B) POSICAO da marca caindo (concorrente/terceiro ultrapassando voce
         na busca do seu proprio nome - prioridade maxima, mesmo com volume
         baixo, porque isso e estrutural, nao ruido).
    """
    alerts = []
    if current_df.empty:
        return alerts

    brand_current = current_df[current_df["query"].apply(is_brand)]
    if brand_current.empty:
        return alerts

    for _, r in brand_current.iterrows():
        query, page = r["query"], r["page"]
        k = (query, page)
        st = trailing_stats.get(k)
        if not st:
            continue

        impr = float(r["impressions"])
        pos = float(r["position"])

        # B) posicao caindo - grave, mas so se a confianca for real. Sem
        # isso, qualquer variante de marca de cauda longa com 2-3 impressoes
        # (posicao naturalmente ruidosa) disparava "urgente" por oscilacao
        # normal, e query em posicao 90+ (invisivel) alertava mesmo sem
        # significado operacional nenhum.
        if (st["confiavel"] and st["media_position"] is not None
                and st["media_position"] <= G7_MAX_MEANINGFUL_POSITION):
            pos_drop = pos - st["media_position"]
            if pos_drop >= G7_POSITION_DROP_ALERT:
                alerts.append({
                    "gatilho": "7. Marca - POSICAO CAINDO",
                    "query": query, "pagina": page, "posicao": round(pos, 1),
                    "posicao_media_4sem": round(st["media_position"], 1),
                    "impressoes_semana": int(impr),
                    "motivo": (f"Posicao da marca caiu de {st['media_position']:.1f} para "
                              f"{pos:.1f}. Alguem pode estar superando voce na busca do "
                              f"seu proprio nome - verificar manualmente com urgencia."),
                    "confianca": "ALTA",
                    "urgencia": "MAXIMA",
                })
                continue   # nao precisa checar o outro motivo tambem

        # A) volume caindo, posicao estavel
        if not st["confiavel"] or st["media_impr"] < G7_MIN_AVG_IMPRESSIONS:
            continue
        drop = (st["media_impr"] - impr) / st["media_impr"]
        if drop < G7_DROP_RATIO:
            continue
        pos_variation = abs(pos - st["media_position"]) if st["media_position"] else 999
        if pos_variation >= G2_MAX_POSITION_VARIATION:
            continue

        alerts.append({
            "gatilho": "7. Marca - trafego caindo",
            "query": query, "pagina": page, "posicao": round(pos, 1),
            "impressoes_semana": int(impr), "media_impr_4sem": round(st["media_impr"], 0),
            "motivo": (f"Impressoes de marca cairam {drop*100:.0f}% vs media, posicao "
                      f"estavel. Google mostrando menos a marca - checar Knowledge Panel, "
                      f"GBP, ou AI Overview no lugar do seu site."),
            "confianca": "ALTA",
            "urgencia": "ALTA",
        })

    alerts.sort(key=lambda x: (x.get("urgencia") != "MAXIMA", -x["impressoes_semana"]))
    return alerts


# ===========================================================================
# RADAR - VISIBILIDADE COMPLETA (nao dispara alerta, so mostra o estado)
# ===========================================================================
def build_radar(current_df, trailing_stats, all_alerts):
    """
    Uma linha por pagina com volume real: impressoes/clicks atuais, media de
    4 semanas, tendencia, e se algum gatilho pegou ela essa semana. Isso
    complementa os 7 gatilhos - eles avisam quando algo precisa de acao, o
    Radar deixa ver o estado de tudo, mesmo o que nao disparou nada.
    """
    if current_df.empty:
        return []

    flagged_pages = defaultdict(list)
    for a in all_alerts:
        if "pagina" in a:
            flagged_pages[norm_path(a["pagina"])].append(a["gatilho"])

    page_agg = current_df.groupby("page", as_index=False).agg(
        impressoes=("impressions", "sum"), clicks=("clicks", "sum"),
        posicao=("position", "mean"))
    page_agg = page_agg[page_agg["impressoes"] >= 30]   # corta ruido de cauda longa

    page_media = defaultdict(float)
    page_media_clicks = defaultdict(float)
    for (query, page), st in trailing_stats.items():
        path = norm_path(page)
        page_media[path] += st["media_impr"]
        page_media_clicks[path] += st["media_clicks"]

    rows = []
    for _, r in page_agg.iterrows():
        path = norm_path(r["page"])
        media = page_media.get(path, 0)
        tendencia = ""
        if media > 0:
            var = (r["impressoes"] - media) / media
            tendencia = f"{var*100:+.0f}%"

        rows.append({
            "pagina": path,
            "impressoes_semana": int(r["impressoes"]),
            "clicks_semana": int(r["clicks"]),
            "posicao": round(float(r["posicao"]), 1),
            "media_impr_4sem": round(media, 0),
            "tendencia": tendencia,
            "gatilhos_ativos": ", ".join(sorted(set(flagged_pages.get(path, [])))) or "-",
        })

    rows.sort(key=lambda x: -x["impressoes_semana"])
    return rows


# ===========================================================================
# DIAGNOSTICO
# ===========================================================================
def print_funnel_diagnostic(current_df):
    """
    0 alertas pode ser 'esta tudo bem' ou pode ser 'o filtro nunca chega no
    final'. Sem isso os dois casos ficam indistinguiveis - e nao vou
    reportar saude do site sem checar qual dos dois esta acontecendo.
    Roda toda semana, nao so uma vez - se o classificador ficar ruim com o
    tempo (site muda), isso aparece aqui antes de virar zero alertas silencioso.
    """
    print("=" * 70)
    print("DIAGNOSTICO DO FUNIL")
    print("=" * 70)
    total_queries = current_df["query"].nunique()
    comm = current_df[current_df["query"].apply(is_commercial_local)]
    n_comm = comm["query"].nunique()
    n_comm_impr = comm[comm["impressions"] >= G1_MIN_IMPRESSIONS]["query"].nunique()
    n_comm_impr_pos = comm[(comm["impressions"] >= G1_MIN_IMPRESSIONS)
                           & (comm["position"] <= G1_MAX_POSITION)]["query"].nunique()
    brand = current_df[current_df["query"].apply(is_brand)]
    print(f"  Queries unicas na semana: {total_queries}")
    print(f"  Classificadas comercial/local: {n_comm}")
    print(f"    ...com {G1_MIN_IMPRESSIONS}+ impr numa semana: {n_comm_impr}")
    print(f"    ...E posicao <= {G1_MAX_POSITION}: {n_comm_impr_pos}")
    print(f"  Classificadas como marca: {brand['query'].nunique()}")
    if n_comm == 0:
        print("\n  ALERTA: regex comercial/local nao capturou nenhuma query.")
    if brand.empty:
        print("\n  ALERTA: regex de marca nao capturou nenhuma query - confirmar variacoes do nome.")

    # Distribuicao real de impressoes entre queries comerciais - para
    # calibrar G1_MIN_IMPRESSIONS com evidencia, nao com outro chute.
    if n_comm > 0:
        impr_por_query = comm.groupby("query")["impressions"].sum()
        pcts = impr_por_query.quantile([0.5, 0.75, 0.90, 0.95, 0.99])
        print(f"\n  Distribuicao de impressoes/semana entre as {n_comm} queries comerciais:")
        print(f"    mediana: {pcts[0.5]:.0f}  |  p75: {pcts[0.75]:.0f}  |  "
              f"p90: {pcts[0.90]:.0f}  |  p95: {pcts[0.95]:.0f}  |  p99: {pcts[0.99]:.0f}")
        if n_comm_impr == 0 and pcts[0.99] < G1_MIN_IMPRESSIONS:
            print(f"    O piso de {G1_MIN_IMPRESSIONS} esta acima do p99 - nenhuma query")
            print(f"    individual vai bater isso numa semana so. Considerar baixar para")
            print(f"    perto do p90 ({pcts[0.90]:.0f}) na proxima calibracao.")
    print()


# ===========================================================================
# EXCEL
# ===========================================================================
HEADER = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
RED = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
ORANGE = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
YELLOW = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
GREEN = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")


def _sheet(wb, title, rows, columns):
    ws = wb.create_sheet(title)
    for i, (key, h, w) in enumerate(columns, 1):
        c = ws.cell(1, i, h)
        c.font = Font(bold=True, color="FFFFFF", size=10)
        c.fill = HEADER
        c.alignment = Alignment(vertical="center", horizontal="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.row_dimensions[1].height = 26
    for ri, row in enumerate(rows, 2):
        for ci, (key, _, _) in enumerate(columns, 1):
            cell = ws.cell(ri, ci, row.get(key, ""))
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            if key == "urgencia" and row.get("urgencia") == "MAXIMA":
                cell.fill = RED
                cell.font = Font(bold=True)
            elif key == "confianca":
                fill = {"ALTA": GREEN, "MEDIA": YELLOW, "BAIXA": ORANGE}.get(row.get("confianca"))
                if fill:
                    cell.fill = fill
    ws.freeze_panes = "A2"
    if rows:
        ws.auto_filter.ref = f"A1:{get_column_letter(len(columns))}{len(rows)+1}"
    return ws


def write_report(g1, g2, g3, g4, g5, g6, g7, radar, run_date, path="turfresh_alertas.xlsx"):
    wb = Workbook()
    ws = wb.active
    ws.title = "Resumo"
    ws["A1"] = f"TurFresh - Alertas Semanais - {run_date.isoformat()}"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A3"] = f"Gatilho 1 (vazamento CTR): {len(g1)}"
    ws["A4"] = f"Gatilho 2 (queda precoce): {len(g2)}"
    ws["A5"] = f"Gatilho 3 (query nova): {len(g3)}"
    ws["A6"] = f"Gatilho 4 (decaimento posts): {len(g4)}"
    ws["A7"] = f"Gatilho 5 (sinal de vida): {len(g5)}"
    ws["A8"] = f"Gatilho 6 (city pages): {len(g6)}"
    ws["A9"] = f"Gatilho 7 (marca): {len(g7)}"
    ws.column_dimensions["A"].width = 45

    G1_COLS = [("query","Query",40),("pagina","Pagina",38),("posicao","Pos",8),
               ("impressoes_semana","Impr/sem",10),("ctr_semana","CTR",9),
               ("media_ctr_4sem","CTR medio 4sem",13),("confianca","Confianca",11),
               ("motivo","Motivo",60)]
    _sheet(wb, "G1 Vazamento CTR", g1, G1_COLS)

    G2_COLS = [("query","Query",40),("pagina","Pagina",38),("posicao","Pos",8),
               ("posicao_media_4sem","Pos media 4sem",13),
               ("impressoes_semana","Impr/sem",10),("media_impr_4sem","Media 4sem",12),
               ("queda_pct","Queda",9),("confianca","Confianca",11),("motivo","Motivo",60)]
    _sheet(wb, "G2 Queda Precoce", g2, G2_COLS)

    G3_COLS = [("query","Query",40),("pagina","Pagina",38),("posicao","Pos",8),
               ("impressoes_semana","Impr/sem",10),("tipo","Tipo",28),
               ("confianca","Confianca",11),("motivo","Motivo",55)]
    _sheet(wb, "G3 Query Nova", g3, G3_COLS)

    G4_COLS = [("pagina","Pagina",42),("clicks_semana","Clicks/sem",11),
               ("media_clicks_4sem","Media 4sem",12),("impressoes_semana","Impr/sem",10),
               ("posicao","Pos",8),("queda_pct","Queda",9),
               ("data_otimizacao","Data otim.",11),("confianca","Confianca",11),
               ("motivo","Motivo",55)]
    _sheet(wb, "G4 Posts Decaindo", g4, G4_COLS)

    G5_COLS = [("pagina","Pagina",42),("dias_desde_publicacao","Dias",8),
               ("impressoes_semana","Impr/sem",10),("confianca","Confianca",11),
               ("motivo","Motivo",65)]
    _sheet(wb, "G5 Sinal de Vida", g5, G5_COLS)

    G6_COLS = [("gatilho","Direcao",22),("pagina","Pagina",42),
               ("impressoes_semana","Impr/sem",10),("media_impr_4sem","Media 4sem",12),
               ("confianca","Confianca",11),("motivo","Motivo",60)]
    _sheet(wb, "G6 City Pages", g6, G6_COLS)

    G7_COLS = [("urgencia","Urgencia",10),("gatilho","Tipo",25),("query","Query",30),
               ("pagina","Pagina",38),("posicao","Pos",8),
               ("posicao_media_4sem","Pos media 4sem",13),
               ("impressoes_semana","Impr/sem",10),("confianca","Confianca",11),
               ("motivo","Motivo",60)]
    _sheet(wb, "G7 Marca", g7, G7_COLS)

    RADAR_COLS = [("pagina","Pagina",42),("impressoes_semana","Impr/sem",11),
                  ("clicks_semana","Clicks/sem",11),("posicao","Pos",8),
                  ("media_impr_4sem","Media 4sem",12),("tendencia","Tendencia",11),
                  ("gatilhos_ativos","Gatilhos ativos",35)]
    _sheet(wb, "Radar", radar, RADAR_COLS)

    wb.save(path)
    return path


# ===========================================================================
# EMAIL
# ===========================================================================
GMAIL_USER = os.environ.get("GMAIL_USER")
GMAIL_APP_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD")
ALERT_EMAIL_TO = os.environ.get("ALERT_EMAIL_TO")

MAX_ITEMS_PER_SECTION = 8   # o e-mail mostra o topo; a planilha tem a lista inteira

GATILHO_INFO = [
    ("g1", "Vazamento de CTR", "#FCE4D6"),
    ("g2", "Queda de impressoes precoce", "#FFF2CC"),
    ("g3", "Query nova emergindo", "#DDEEFF"),
    ("g4", "Decaimento de posts otimizados", "#FCE4D6"),
    ("g5", "Sinal de vida de conteudo novo", "#DDEEFF"),
    ("g6", "City pages (subindo/caindo)", "#FFF2CC"),
    ("g7", "Trafego de marca", "#FFF2CC"),
]


def _html_escape(s):
    return (str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def _build_html_email(data, run_date, total, marca_maxima):
    """
    Estrutura visual: caixa vermelha de urgencia no topo (se houver), depois
    um resumo em grade com contagem por gatilho, depois cada gatilho que tem
    alerta vira uma secao com no maximo MAX_ITEMS_PER_SECTION linhas - o
    resto fica so na planilha anexa, para o e-mail nao virar parede de texto.
    """
    css_box = ("border-radius:8px;padding:16px 20px;margin-bottom:16px;"
               "font-family:Arial,Helvetica,sans-serif;")
    css_table = ("width:100%;border-collapse:collapse;margin-bottom:20px;"
                "font-family:Arial,Helvetica,sans-serif;font-size:13px;")

    html = [f'<div style="font-family:Arial,Helvetica,sans-serif;max-width:720px;">']
    html.append(f'<h2 style="color:#1F3864;margin-bottom:4px;">TurFresh - Alertas Semanais</h2>')
    html.append(f'<p style="color:#666;margin-top:0;">{run_date.isoformat()}</p>')

    if marca_maxima:
        html.append(f'<div style="{css_box}background:#FFEBEE;border:2px solid #D32F2F;">')
        html.append('<strong style="color:#D32F2F;font-size:15px;">⚠ URGENTE - Posicao de marca caindo</strong>')
        html.append('<ul style="margin:8px 0 0 0;padding-left:20px;">')
        for a in marca_maxima[:MAX_ITEMS_PER_SECTION]:
            html.append(f'<li><b>{_html_escape(a["query"])}</b>: posicao '
                        f'{a["posicao_media_4sem"]} → {a["posicao"]} '
                        f'<span style="color:#888;">({_html_escape(a["pagina"])})</span></li>')
        if len(marca_maxima) > MAX_ITEMS_PER_SECTION:
            html.append(f'<li><i>+{len(marca_maxima)-MAX_ITEMS_PER_SECTION} outras na planilha</i></li>')
        html.append('</ul></div>')

    # resumo em grade
    html.append(f'<table style="{css_table}"><tr>')
    html.append('<td colspan="2" style="background:#1F3864;color:white;padding:8px 12px;'
                'font-weight:bold;">Resumo</td></tr>')
    for key, label, color in GATILHO_INFO:
        n = len(data[key])
        bg = color if n > 0 else "#F5F5F5"
        html.append(f'<tr><td style="padding:6px 12px;border-bottom:1px solid #eee;'
                    f'background:{bg};">{label}</td>'
                    f'<td style="padding:6px 12px;border-bottom:1px solid #eee;'
                    f'background:{bg};text-align:right;font-weight:bold;">{n}</td></tr>')
    html.append(f'<tr><td style="padding:8px 12px;font-weight:bold;">Total</td>'
               f'<td style="padding:8px 12px;text-align:right;font-weight:bold;">{total}</td></tr>')
    html.append('</table>')

    # secoes detalhadas (so as que tem alerta, so o topo)
    section_renderers = {
        "g1": lambda a: f'<b>{_html_escape(a["query"])}</b> - CTR {a["ctr_semana"]} '
                        f'(pos {a["posicao"]}, {a["impressoes_semana"]:,} impr)',
        "g2": lambda a: f'<b>{_html_escape(a["query"])}</b> - impressoes cairam {a["queda_pct"]}',
        "g3": lambda a: f'<b>{_html_escape(a["query"])}</b> - {a["impressoes_semana"]} impr '
                        f'({a.get("tipo","")})',
        "g4": lambda a: f'<b>{_html_escape(a["pagina"])}</b> - clicks cairam {a["queda_pct"]}',
        "g5": lambda a: f'<b>{_html_escape(a["pagina"])}</b> - {a["motivo"]}',
        "g6": lambda a: f'<b>{_html_escape(a["pagina"])}</b> - {a["gatilho"].replace("6. ","")}',
        "g7": lambda a: f'<b>{_html_escape(a["query"])}</b> - {a["gatilho"].replace("7. Marca - ","")}',
    }
    for key, label, color in GATILHO_INFO:
        rows = data[key]
        if not rows:
            continue
        html.append(f'<div style="{css_box}background:{color};">')
        html.append(f'<strong>{label} ({len(rows)})</strong>')
        html.append('<ul style="margin:8px 0 0 0;padding-left:20px;">')
        for a in rows[:MAX_ITEMS_PER_SECTION]:
            html.append(f'<li>{section_renderers[key](a)}</li>')
        if len(rows) > MAX_ITEMS_PER_SECTION:
            html.append(f'<li><i>+{len(rows)-MAX_ITEMS_PER_SECTION} outras na planilha anexa</i></li>')
        html.append('</ul></div>')

    html.append('<p style="color:#666;font-size:12px;">Detalhe completo, com evidencia e '
               'motivo de cada alerta, na planilha anexa.</p>')
    html.append('</div>')
    return "\n".join(html)


def send_email(g1, g2, g3, g4, g5, g6, g7, report_path, run_date):
    import smtplib
    from email.message import EmailMessage

    data = {"g1": g1, "g2": g2, "g3": g3, "g4": g4, "g5": g5, "g6": g6, "g7": g7}
    total = sum(len(v) for v in data.values())
    marca_maxima = [a for a in g7 if a.get("urgencia") == "MAXIMA"]

    # fallback em texto puro, para clientes de email sem suporte a HTML
    text_lines = [f"TurFresh - Alertas Semanais - {run_date.isoformat()}", "=" * 55, ""]
    if marca_maxima:
        text_lines.append("URGENTE - POSICAO DE MARCA CAINDO:")
        for a in marca_maxima[:MAX_ITEMS_PER_SECTION]:
            text_lines.append(f"  {a['query']}: {a['posicao_media_4sem']} -> {a['posicao']}")
        text_lines.append("")
    text_lines.append(f"Total: {total} alertas")
    for key, label, _ in GATILHO_INFO:
        text_lines.append(f"  {label}: {len(data[key])}")
    text_lines.append("\nDetalhe completo na planilha anexa.")
    text = "\n".join(text_lines)

    html = _build_html_email(data, run_date, total, marca_maxima)

    if not (GMAIL_USER and GMAIL_APP_PASSWORD and ALERT_EMAIL_TO):
        print(text)
        return

    msg = EmailMessage()
    prefix = "[URGENTE] " if marca_maxima else ""
    msg["Subject"] = f"{prefix}[TurFresh] {total} alertas - {run_date.isoformat()}"
    msg["From"] = GMAIL_USER
    msg["To"] = ALERT_EMAIL_TO
    msg.set_content(text)
    msg.add_alternative(html, subtype="html")
    with open(report_path, "rb") as f:
        msg.add_attachment(f.read(), maintype="application",
                           subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           filename=f"turfresh-alertas-{run_date.isoformat()}.xlsx")
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as s:
        s.login(GMAIL_USER, GMAIL_APP_PASSWORD)
        s.send_message(msg)
    print("Email enviado (HTML).")
    print(text)


# ===========================================================================
# MAIN

# ===========================================================================
def main():
    run_date = date.today()
    service = get_gsc_service()
    current, trailing, windows = fetch_trailing_weeks(service)

    print("Calculando medias moveis de 4 semanas...")
    stats = build_trailing_stats(trailing)
    print(f"  {len(stats)} combinacoes query+pagina no historico\n")

    print("Carregando dados de referencia...")
    seo_log_urls = load_seo_log_urls()
    city_pages = load_city_pages()
    print()

    print_funnel_diagnostic(current)

    print("Rodando os 7 gatilhos...")
    g1 = gatilho_1_vazamento_ctr(current, stats)
    g2 = gatilho_2_queda_precoce(current, stats)
    g3 = gatilho_3_query_nova(current, stats)
    g4 = gatilho_4_decaimento_posts(current, stats, seo_log_urls)
    g5 = gatilho_5_sinal_vida(current, city_pages, run_date)
    g6 = gatilho_6_city_pages(current, stats, city_pages)
    g7 = gatilho_7_marca(current, stats)

    all_alerts = g1 + g2 + g3 + g4 + g5 + g6 + g7
    print(f"  G1 vazamento CTR: {len(g1)}")
    print(f"  G2 queda precoce: {len(g2)}")
    print(f"  G3 query nova: {len(g3)}")
    print(f"  G4 decaimento posts: {len(g4)}")
    print(f"  G5 sinal de vida: {len(g5)}")
    print(f"  G6 city pages: {len(g6)}")
    print(f"  G7 marca: {len(g7)}")
    print(f"  TOTAL: {len(all_alerts)} alertas\n")

    radar = build_radar(current, stats, all_alerts)
    print(f"Radar: {len(radar)} paginas com volume real\n")

    report_path = write_report(g1, g2, g3, g4, g5, g6, g7, radar, run_date)
    print(f"Relatorio: {report_path}\n")

    send_email(g1, g2, g3, g4, g5, g6, g7, report_path, run_date)


if __name__ == "__main__":
    main()
