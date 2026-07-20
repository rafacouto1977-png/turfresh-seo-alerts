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
G1_MIN_IMPRESSIONS = 500
G1_MAX_POSITION = 8
G1_CTR_ABSOLUTE_FLOOR = 0.01     # 1%
G1_CTR_DROP_RATIO = 0.40         # queda de 40% vs media

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
    return pd.DataFrame(rows_out)


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
# MAIN
# ===========================================================================
def main():
    service = get_gsc_service()
    current, trailing, windows = fetch_trailing_weeks(service)

    print("Calculando medias moveis de 4 semanas...")
    stats = build_trailing_stats(trailing)
    print(f"  {len(stats)} combinacoes query+pagina no historico\n")

    print("Rodando Gatilho 1 (vazamento de CTR)...")
    g1 = gatilho_1_vazamento_ctr(current, stats)
    print(f"  {len(g1)} alertas\n")

    print("=" * 70)
    print(f"RESULTADOS - {len(g1)} alertas de vazamento de CTR")
    print("=" * 70)
    for a in g1[:20]:
        print(f"\n[{a['confianca']}] {a['query']}")
        print(f"  {a['pagina']}")
        print(f"  pos {a['posicao']} | {a['impressoes_semana']:,} impr | "
              f"CTR {a['ctr_semana']} (media 4sem: {a['media_ctr_4sem']})")
        print(f"  {a['motivo']}")


if __name__ == "__main__":
    main()
