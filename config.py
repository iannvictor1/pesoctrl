from pathlib import Path

BASE = Path(r"C:\Users\Supervidor Externo\OneDrive\Desktop\Controle_Pesagem")

ARQUIVO_RECEBIMENTO_GLOBAL = BASE / "recebimento_global.json"

CONFIG_IMPRESSORAS = {
    "impressora_1": {
        "nome_amigavel": "Impressora 1",
        "base": BASE / "impressora_1",
    },
    "impressora_2": {
        "nome_amigavel": "Impressora 2",
        "base": BASE / "impressora_2",
    },
}


def montar_config(nome_impressora: str) -> dict:
    if nome_impressora not in CONFIG_IMPRESSORAS:
        raise ValueError(f"Impressora inválida: {nome_impressora}")

    base_impressora = CONFIG_IMPRESSORAS[nome_impressora]["base"]

    cfg = {
        "id": nome_impressora,
        "nome_amigavel": CONFIG_IMPRESSORAS[nome_impressora]["nome_amigavel"],
        "base": base_impressora,
        "pasta_origem": base_impressora / "entrada_etiquetas",
        "pasta_fila": base_impressora / "fila_processamento",
        "pasta_historico_etiquetas": base_impressora / "historico_etiquetas",
        "pasta_historico_pesagens": base_impressora / "historico_pesagens",
        "pasta_descarte_fila": base_impressora / "fila_descartada",
        "arquivo_excel_atual": base_impressora / "pesagens_em_andamento.xlsx",
        "arquivo_log": base_impressora / "monitor_log.txt",
        "arquivo_status": base_impressora / "sessao_status.json",
        "arquivo_pid": base_impressora / "monitor_pid.json",
    }

    return cfg