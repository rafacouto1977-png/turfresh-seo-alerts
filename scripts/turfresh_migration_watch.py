"""
TurFresh - Migration Watch (checagem diaria pos-migracao)
============================================================
Separado do turfresh_alert.py de proposito: aquele responde "que padrao
esta mudando aos poucos" (comparacao semanal, por query). Este responde
"algo quebrou agora" (comparacao diaria, agregada no nivel do site). Sao
perguntas diferentes, com cadencias diferentes - rodar a logica pesada dos
9 gatilhos todo dia nao ajudaria em nada e arriscaria mexer no script que
ja esta estavel em producao.

Pensado para o watch intensivo dos 30 dias pos-migracao (23/set a
23/out/2026). Depois disso, considere desligar o cron deste workflow (ou
trocar para semanal, acompanhando a comparacao de 30-90 dias que ja esta
no ClickUp).

Dois sinais, nada mais:

1. INDEXACAO (via API de Sitemaps, mesmo client GSC de sempre). Estrutural:
   uma queda no numero de paginas indexadas nao precisa de "confirmacao em
   dias seguidos" como trafego - indexacao nao oscila por ruido do jeito
   que cliques oscilam. Isto e um PROXY, nao o Relatorio de Indexacao de
   Paginas em si (esse relatorio nao tem endpoint na API) - ajuda a pegar
   quedas estruturais sem esperar a revisao manual diaria, mas nao
   substitui ela.

2. TRAFEGO AGREGADO DO SITE (cliques/impressoes por dia, sem quebrar por
   query/pagina), comparado contra a media do MESMO DIA DA SEMANA no
   periodo pre-migracao (terca contra a media das tercas, nao contra uma
   janela corrida de 7 dias - isso misturaria fim de semana com dia util
   e geraria alarme falso). A baseline fica CONGELADA nas semanas antes do
   cutover, recalculada do zero a cada execucao a partir do historico do
   GSC (nao precisa persistir) - assim ela nunca comeca a incorporar dias
   pos-migracao como se fossem "normais".

Regra de decisao (numeros do brief de migracao do Rafael + um gatilho
imediato para quedas obvias):
  - queda < 30% num dia: normal ou dentro do esperado (o brief fala em
    10-15% de oscilacao normal - a faixa entre isso e 30% fica no log da
    Action, sem e-mail, porque um dia isolado nessa faixa ainda pode ser
    ruido).
  - queda >= 50% NUM DIA SO: e-mail na hora. Um tombo desse tamanho e
    extremamente improvavel de ser ruido - nao vale esperar confirmacao.
  - queda >= 30% por 2 DIAS SEGUIDOS: e-mail. Esse e o numero exato que o
    Rafael definiu como "sinal real" no brief da migracao - so exige dois
    dias seguidos antes de mandar, porque um dia so nessa faixa ainda pode
    ser oscilacao normal.
  - indexacao caiu (qualquer valor, 1 dia basta): sempre e-mail.

Silencio proposital fora desses casos: rodando todo dia, mandar e-mail
mesmo sem nada para reportar viraria ruido em poucas semanas. Sem e-mail
nenhum dia = tudo dentro do esperado.
"""

import json
import os
from collections import defaultdict
from datetime import date, timedelta

from google.oauth2 import service_account
from googleapiclient.discovery import build

# ===========================================================================
# CONFIG
# ===========================================================================
SITE_URL = os.environ.get("SITE_URL", "https://turfresh.com/")
GSC_CLIENT_EMAIL = os.environ.get("GSC_CLIENT_EMAIL")
GSC_PRIVATE_KEY = os.environ.get("GSC_PRIVATE_KEY")
GMAIL_USER = os.environ.get("GMAIL_USER")
GMAIL_APP_PASSWORD = os.environ.get("GMAIL_APP_PASSWORD")
ALERT_EMAIL_TO = os.environ.get("ALERT_EMAIL_TO")

GSC_DAILY_LAG_DAYS = 3   # mesmo lag do script semanal - dado do dia so amadurece depois de ~3 dias

MIGRATION_CUTOVER_DATE = date(2026, 9, 22)   # ultimo dia confirmado pre-migracao
BASELINE_WEEKS = 8
BASELINE_END = MIGRATION_CUTOVER_DATE - timedelta(days=1)          # 21/09/2026
BASELINE_START = BASELINE_END - timedelta(weeks=BASELINE_WEEKS) + timedelta(days=1)

MIN_BASELINE_IMPRESSIONS = 200   # piso: dia da semana com baseline fraca demais nao entra na comparacao

DROP_IMMEDIATE_RATIO = 0.50      # queda brusca de 1 dia so - alerta na hora, sem esperar confirmacao
DROP_SUSTAINED_RATIO = 0.30      # numero do brief do Rafael para "sinal real"
DROP_SUSTAINED_STREAK_DAYS = 2   # so alerta nessa faixa se persistir por 2 dias seguidos

STATE_PATH = "data/migration_watch_state.json"
WEEKDAY_NAMES = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]


# ===========================================================================
# GSC
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


def fetch_daily_series(service, start_date, end_date):
    """Uma linha por dia (nao por query+pagina) - agregado do site inteiro."""
    req = {"startDate": start_date.isoformat(), "endDate": end_date.isoformat(),
           "dimensions": ["date"], "rowLimit": 5000}
    resp = service.searchanalytics().query(siteUrl=SITE_URL, body=req).execute()
    out = {}
    for r in resp.get("rows", []):
        out[r["keys"][0]] = {"clicks": r["clicks"], "impressions": r["impressions"],
                             "position": r["position"]}
    return out


def get_latest_available_day():
    return date.today() - timedelta(days=GSC_DAILY_LAG_DAYS)


def fetch_single_day(service, day):
    return fetch_daily_series(service, day, day).get(day.isoformat())


def compute_weekday_baseline(service):
    """
    Media por dia da semana no periodo pre-migracao, recalculada do zero a
    cada execucao (nao precisa persistir - o GSC guarda o historico). Assim
    a baseline nunca "desliza" incorporando dias pos-migracao como normais.
    """
    series = fetch_daily_series(service, BASELINE_START, BASELINE_END)
    by_weekday = defaultdict(lambda: {"impr": [], "clicks": []})
    for d_str, vals in series.items():
        d = date.fromisoformat(d_str)
        by_weekday[d.weekday()]["impr"].append(vals["impressions"])
        by_weekday[d.weekday()]["clicks"].append(vals["clicks"])

    baseline = {}
    for wd, vals in by_weekday.items():
        n = len(vals["impr"])
        if n == 0:
            continue
        baseline[wd] = {"avg_impressions": sum(vals["impr"]) / n,
                        "avg_clicks": sum(vals["clicks"]) / n, "n_days": n}
    return baseline


def evaluate_traffic_drop(day, day_data, baseline):
    """
    Compara o dia contra a media do MESMO dia da semana, nao contra uma
    janela corrida - domingo tem trafego menor que terca por natureza, e
    misturar isso geraria alarme falso toda semana.
    """
    wd = day.weekday()
    base = baseline.get(wd)
    if not day_data or not base or base["avg_impressions"] < MIN_BASELINE_IMPRESSIONS:
        return None

    impr_drop = (base["avg_impressions"] - day_data["impressions"]) / base["avg_impressions"]
    clicks_drop = ((base["avg_clicks"] - day_data["clicks"]) / base["avg_clicks"]
                  if base["avg_clicks"] > 0 else 0.0)

    return {
        "weekday": WEEKDAY_NAMES[wd],
        "impressions_today": day_data["impressions"],
        "impressions_baseline": round(base["avg_impressions"]),
        "impr_drop_pct": impr_drop,
        "clicks_today": day_data["clicks"],
        "clicks_baseline": round(base["avg_clicks"]),
        "clicks_drop_pct": clicks_drop,
        "worst_drop_pct": max(impr_drop, clicks_drop),
        # posicao agregada do site inteiro nao vira gatilho - e dominada por
        # QUAIS queries tiveram impressao naquele dia, ruido demais em
        # nivel de agregado. Fica so no log, para contexto visual.
        "position_today": round(day_data.get("position", 0), 1),
    }


def check_indexed_pages(service, state):
    """
    Usa a API de Sitemaps (mesmo service da Search Console) para somar
    paginas indexadas em todos os sitemaps submetidos, incluindo os filhos
    de um sitemap index (padrao do Rank Math: sitemap_index.xml apontando
    para page-sitemap.xml, post-sitemap.xml etc). Isto e um PROXY do
    Relatorio de Indexacao de Paginas, nao o relatorio em si - atualiza mais
    devagar, porque o Google so reprocessa o sitemap de tempos em tempos.
    Por isso so alerta em QUEDA; nunca conclui "tudo ok" so por estabilidade.
    """
    try:
        resp = service.sitemaps().list(siteUrl=SITE_URL).execute()
    except Exception as e:
        print(f"  Aviso: nao foi possivel ler sitemaps ({e}) - pulando checagem de indexacao.")
        return None

    total_indexed = 0
    total_submitted = 0
    found_any_contents = False

    for entry in resp.get("sitemap", []):
        sub_entries = [entry]
        if entry.get("isSitemapsIndex"):
            try:
                child_resp = service.sitemaps().list(
                    siteUrl=SITE_URL, sitemapIndex=entry["path"]).execute()
                sub_entries = child_resp.get("sitemap", [])
            except Exception as e:
                print(f"  Aviso: falha ao ler sub-sitemaps de {entry.get('path')}: {e}")
                continue
        for sub in sub_entries:
            for c in sub.get("contents", []):
                found_any_contents = True
                total_indexed += int(c.get("indexed", 0))
                total_submitted += int(c.get("submitted", 0))

    if not found_any_contents:
        print("  Aviso: sitemaps sem dado de 'contents' ainda (comum logo apos migracao) - pulando comparacao.")
        return None

    previous = state.get("last_indexed_count")
    result = {"indexed_now": total_indexed, "submitted_now": total_submitted,
              "indexed_previous": previous,
              "dropped": previous is not None and total_indexed < previous}
    state["last_indexed_count"] = total_indexed
    return result


# ===========================================================================
# ESTADO (persistido no repo entre execucoes - runner do GH Actions e stateless)
# ===========================================================================
def load_state():
    if os.path.exists(STATE_PATH):
        with open(STATE_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_state(state):
    os.makedirs(os.path.dirname(STATE_PATH), exist_ok=True)
    with open(STATE_PATH, "w", encoding="utf-8") as f:
        json.dump(state, f, indent=2, sort_keys=True)


# ===========================================================================
# E-MAIL
# ===========================================================================
def send_alert_email(alerts, day):
    import smtplib
    from email.message import EmailMessage

    subject = f"[URGENT] [TurFresh Migration Watch] {day.isoformat()}"

    text_lines = [f"TurFresh Migration Watch - {day.isoformat()}", "=" * 50, ""]
    html = ['<div style="font-family:Arial,Helvetica,sans-serif;max-width:640px;">',
            '<h2 style="color:#1F3864;margin-bottom:4px;">TurFresh Migration Watch</h2>',
            f'<p style="color:#666;margin-top:0;">{day.isoformat()}</p>']

    for a in alerts:
        text_lines.append(f"[{a['subject']}]")
        text_lines.append(a["body"])
        text_lines.append("")
        html.append('<div style="border-left:4px solid #D32F2F;padding:10px 14px;'
                    'margin-bottom:10px;background:#FAFAFA;'
                    'font-family:Arial,Helvetica,sans-serif;">'
                    f'<div style="font-weight:bold;color:#D32F2F;">{a["subject"]}</div>'
                    f'<div style="font-size:13px;color:#333;margin-top:4px;">{a["body"]}</div>'
                    '</div>')
    html.append('</div>')

    text = "\n".join(text_lines)
    html_str = "\n".join(html)

    if not (GMAIL_USER and GMAIL_APP_PASSWORD and ALERT_EMAIL_TO):
        print(text)
        return

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = GMAIL_USER
    msg["To"] = ALERT_EMAIL_TO
    msg.set_content(text)
    msg.add_alternative(html_str, subtype="html")
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as s:
        s.login(GMAIL_USER, GMAIL_APP_PASSWORD)
        s.send_message(msg)
    print("Email enviado.")
    print(text)


# ===========================================================================
# MAIN
# ===========================================================================
def main():
    state = load_state()
    service = get_gsc_service()

    day = get_latest_available_day()
    print(f"Checando dia: {day.isoformat()} ({WEEKDAY_NAMES[day.weekday()]}, "
          f"lag de {GSC_DAILY_LAG_DAYS} dias)\n")

    print(f"Calculando baseline pre-migracao ({BASELINE_START} a {BASELINE_END})...")
    baseline = compute_weekday_baseline(service)
    for wd in sorted(baseline):
        b = baseline[wd]
        print(f"  {WEEKDAY_NAMES[wd]}: media {b['avg_impressions']:.0f} impr / "
              f"{b['avg_clicks']:.0f} clicks ({b['n_days']} semanas de historico)")
    print()

    day_data = fetch_single_day(service, day)
    traffic = evaluate_traffic_drop(day, day_data, baseline)

    if traffic:
        print(f"Trafego do dia ({traffic['weekday']}): {traffic['impressions_today']} impr "
              f"(baseline {traffic['impressions_baseline']}, "
              f"{traffic['impr_drop_pct']*100:+.0f}%) | "
              f"{traffic['clicks_today']} clicks (baseline {traffic['clicks_baseline']}, "
              f"{traffic['clicks_drop_pct']*100:+.0f}%) | posicao media {traffic['position_today']}\n")
    else:
        print("Trafego do dia: sem dado suficiente para comparar (dia sem GSC ou baseline fraca para esse dia da semana).\n")

    streak = state.get("consecutive_severe_days", 0)
    if traffic and traffic["worst_drop_pct"] >= DROP_SUSTAINED_RATIO:
        streak += 1
    else:
        streak = 0
    state["consecutive_severe_days"] = streak
    state["last_check_date"] = day.isoformat()

    print("Checando indexacao via sitemaps...")
    index_check = check_indexed_pages(service, state)
    if index_check:
        print(f"  Indexadas agora: {index_check['indexed_now']} "
              f"(execucao anterior: {index_check['indexed_previous']})\n")

    alerts = []

    if index_check and index_check["dropped"]:
        alerts.append({
            "subject": "Indexed pages dropped",
            "body": (f"Sitemap-reported indexed pages dropped from "
                    f"{index_check['indexed_previous']} to {index_check['indexed_now']}. "
                    f"This is a structural signal, not daily traffic noise - check the "
                    f"Page Indexing report in GSC today."),
        })

    if traffic:
        is_immediate = traffic["worst_drop_pct"] >= DROP_IMMEDIATE_RATIO
        is_sustained = (traffic["worst_drop_pct"] >= DROP_SUSTAINED_RATIO
                        and streak >= DROP_SUSTAINED_STREAK_DAYS)
        if is_immediate or is_sustained:
            reason = ("single-day drop this severe" if is_immediate else
                      f"day {streak} in a row over the 30% threshold")
            alerts.append({
                "subject": f"Traffic drop vs pre-migration baseline ({reason})",
                "body": (f"{traffic['weekday']}: impressions {traffic['impressions_today']} vs "
                        f"baseline {traffic['impressions_baseline']} "
                        f"({traffic['impr_drop_pct']*100:.0f}% drop), clicks "
                        f"{traffic['clicks_today']} vs baseline {traffic['clicks_baseline']} "
                        f"({traffic['clicks_drop_pct']*100:.0f}% drop). "
                        f"Baseline is the pre-migration daily average for this weekday "
                        f"({BASELINE_START.isoformat()} to {BASELINE_END.isoformat()})."),
            })

    save_state(state)

    if not alerts:
        print("Nada acima do limite hoje. Sem e-mail (silencio = tudo dentro do esperado).")
        return

    send_alert_email(alerts, day)


if __name__ == "__main__":
    main()
