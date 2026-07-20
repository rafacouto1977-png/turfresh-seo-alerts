"""
Teste de Conexao - TurFresh GSC Alerts
========================================
Nao faz nenhuma analise. So confirma que a canalizacao inteira funciona:
secrets certos, service account com permissao, formato de URL da
propriedade correto. Roda uma vez, antes de construir os 6 gatilhos
em cima de uma conexao que ainda nao provou que funciona.
"""

import os
from datetime import date, timedelta

from google.oauth2 import service_account
from googleapiclient.discovery import build

SITE_URL = os.environ.get("SITE_URL")
GSC_CLIENT_EMAIL = os.environ.get("GSC_CLIENT_EMAIL")
GSC_PRIVATE_KEY = os.environ.get("GSC_PRIVATE_KEY")


def main():
    print(f"SITE_URL: {SITE_URL}")
    print(f"GSC_CLIENT_EMAIL: {GSC_CLIENT_EMAIL}")
    print(f"GSC_PRIVATE_KEY presente: {'sim' if GSC_PRIVATE_KEY else 'NAO - secret faltando'}")
    print()

    if not (SITE_URL and GSC_CLIENT_EMAIL and GSC_PRIVATE_KEY):
        print("ERRO: um ou mais secrets nao estao configurados.")
        return

    info = {
        "type": "service_account",
        "client_email": GSC_CLIENT_EMAIL,
        "private_key": GSC_PRIVATE_KEY,
        "token_uri": "https://oauth2.googleapis.com/token",
    }
    creds = service_account.Credentials.from_service_account_info(
        info, scopes=["https://www.googleapis.com/auth/webmasters.readonly"]
    )
    service = build("searchconsole", "v1", credentials=creds)
    print("Autenticacao OK - credenciais aceitas pelo Google.\n")

    end = date.today() - timedelta(days=3)
    start = end - timedelta(days=6)
    print(f"Buscando dados de {start} a {end}...\n")

    request = {
        "startDate": start.isoformat(),
        "endDate": end.isoformat(),
        "dimensions": ["page"],
        "rowLimit": 10,
    }
    response = service.searchanalytics().query(siteUrl=SITE_URL, body=request).execute()
    rows = response.get("rows", [])

    if not rows:
        print("CONECTOU, mas voltou 0 linhas. Possiveis causas:")
        print("  - Site sem trafego nesse periodo (improvavel para a TurFresh)")
        print("  - SITE_URL nao bate exatamente com a propriedade no GSC")
        print("  - Service account nao foi adicionada como usuario no GSC")
        return

    total_clicks = sum(r["clicks"] for r in rows)
    total_impressions = sum(r["impressions"] for r in rows)

    print(f"SUCESSO. {len(rows)} paginas retornadas (top 10 por clicks).")
    print(f"Total nessa amostra: {total_clicks} clicks, {total_impressions} impressoes.\n")
    print("Top paginas:")
    for r in rows:
        print(f"  {r['clicks']:>5} clicks | {r['impressions']:>6} impr | {r['keys'][0]}")

    print("\nConexao validada. Pronto para construir os 6 gatilhos em cima disso.")


if __name__ == "__main__":
    main()
