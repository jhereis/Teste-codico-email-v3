# Teste-codico-email-v3

# -*- coding: utf-8 -*-
"""
Envio de e-mails via Outlook a partir de Excel (pandas + win32com)
- Ignora linhas com VALOR vazio/NaN/zero
- CC opcional
- Log de sucesso/erro (CSV/XLSX)
- Popup ao finalizar
"""

import re
import time
import locale
import ctypes
from pathlib import Path

import pandas as pd
import win32com.client as win32

# ===================== CONFIGURAÇÕES =====================

# Caminho da planilha Excel
CAMINHO_XLSX = r"C:\Users\DJHENIF\Downloads\tesour_antecip_prospecc.xlsx"

# (Opcional) nome da aba. Deixe None para a primeira aba
SHEET_NAME = None  # ex: "Base"

# Nomes das colunas na planilha (ajuste se necessário)
COL_EMAIL_DEST   = "EMAIL"
COL_EMAIL_CC     = "EMAIL_RESPONSAVEL"   # opcional
COL_VALOR        = "VALOR"
COL_CNPJ         = "CNPJ"                # opcional no corpo
COL_FORNECEDOR   = "FORNECEDOR"
COL_CLIENTE      = "NOME"
COL_REFERENCIA2  = "REFERENCIA 2"        # se não existir, o script segue sem usar

# Enviar de verdade (True) ou abrir para revisão (False)
ENVIAR_DE_VERDADE = True

# Pausa entre envios (segundos) para evitar throttling
SLEEP_SEGUNDOS = 0.8

# ===================== LOCALE / MOEDA =====================

def configurar_locale():
    for loc in ("pt_BR.UTF-8", "pt_BR.utf8", "Portuguese_Brazil.1252"):
        try:
            locale.setlocale(locale.LC_ALL, loc)
            return
        except locale.Error:
            pass
configurar_locale()

def moeda_brasil(v) -> str:
    try:
        return locale.currency(float(v), grouping=True)
    except Exception:
        return f"R$ {float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

# ===================== UTILIDADES =====================

EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

def email_valido(s: str) -> bool:
    return isinstance(s, str) and EMAIL_RE.match(s) is not None

def limpar_email(x):
    if pd.isna(x):
        return None
    s = str(x).strip()
    return s if email_valido(s) else None

def normalizar_numero(x):
    """Aceita 1234.56, 1.234,56, '1.234,56', '1234,56', etc."""
    if pd.isna(x):
        return None
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).strip()
    if not s:
        return None
    # Se vier no formato brasileiro 1.234,56
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return None

def montar_assunto(nome_cliente: str) -> str:
    return f"FREDE {nome_cliente} – Antecipação de Recebíveis"

def montar_corpo(nome_cliente: str, nome_fornecedor: str, valor_fmt: str) -> str:
    return (
f"""Olá, tudo bem?

Referência: FREDE {nome_fornecedor} – Antecipação de recebíveis e fluxo de caixa.

Gostaríamos de apresentar uma solução para você ter mais fluxo de caixa: a antecipação de recebíveis.

Com taxas mais atrativas do que as praticadas no mercado, sua empresa recebe os valores antes e sem a incidência de inadimplência.

Caso tenha interesse, estamos à disposição para enviar uma cotação personalizada para você.

Volume total aproximado disponível: {valor_fmt}

Caso essa demanda não esteja sob sua responsabilidade, poderia nos indicar o contato adequado?

Atenciosamente,
Tesouraria
""")

# ===================== OUTLOOK =====================

def enviar_email(outlook, destinatario: str, assunto: str, corpo: str, cc: str = None):
    mail = outlook.CreateItem(0)
    mail.To = destinatario
    mail.Subject = assunto
    mail.Body = corpo
    if cc:
        mail.CC = cc
    if ENVIAR_DE_VERDADE:
        mail.Send()
    else:
        mail.Display()

# ===================== MAIN =====================

def main():
    # Carrega planilha
    df = pd.read_excel(CAMINHO_XLSX, sheet_name=SHEET_NAME)

    colunas_necessarias = [COL_EMAIL_DEST, COL_VALOR, COL_FORNECEDOR, COL_CLIENTE]
    faltantes = [c for c in colunas_necessarias if c not in df.columns]
    if faltantes:
        raise ValueError(f"Colunas ausentes na planilha: {faltantes}\nDisponíveis: {list(df.columns)}")

    # Normaliza coluna VALOR e filtra não nulos e > 0
    df[COL_VALOR] = df[COL_VALOR].apply(normalizar_numero)
    df = df[df[COL_VALOR].notna() & (df[COL_VALOR] > 0)].copy()

    # Normaliza e-mails
    df["__email_dest"] = df[COL_EMAIL_DEST].apply(limpar_email)
    if COL_EMAIL_CC in df.columns:
        df["__email_cc"] = df[COL_EMAIL_CC].apply(limpar_email)
    else:
        df["__email_cc"] = None

    total_previsto = len(df)
    print(f"Serão enviados {total_previsto} e-mails (VALOR > 0).")

    # Outlook
    outlook = win32.Dispatch("outlook.application")

    # Logs
    logs = []

    for idx, row in df.iterrows():
        dest = row["__email_dest"]
        cc   = row["__email_cc"]

        # pula se e-mail inválido/ausente
        if not dest:
            logs.append({"linha": int(idx), "destinatario": row.get(COL_EMAIL_DEST, ""),
                         "status": "ERRO", "motivo": "E-mail ausente/ inválido"})
            continue

        nome_cliente    = str(row.get(COL_CLIENTE, "")).strip()
        nome_fornecedor = str(row.get(COL_FORNECEDOR, "")).strip()
        valor_fmt       = moeda_brasil(row.get(COL_VALOR, 0.0))

        assunto = montar_assunto(nome_cliente)
        corpo   = montar_corpo(nome_cliente, nome_fornecedor, valor_fmt)

        try:
            enviar_email(outlook, dest, assunto, corpo, cc=cc)
            logs.append({"linha": int(idx), "destinatario": dest, "cc": cc,
                         "assunto": assunto, "status": "ENVIADO"})
            time.sleep(SLEEP_SEGUNDOS)
        except Exception as e:
            logs.append({"linha": int(idx), "destinatario": dest, "cc": cc,
                         "assunto": assunto, "status": "ERRO", "motivo": str(e)})

    # Relatórios
    relatorio = pd.DataFrame(logs)
    saida_dir = Path(CAMINHO_XLSX).parent
    rel_csv   = saida_dir / "rel_envio_outlook.csv"
    rel_xlsx  = saida_dir / "rel_envio_outlook.xlsx"
    relatorio.to_csv(rel_csv, index=False, encoding="utf-8-sig")
    with pd.ExcelWriter(rel_xlsx, engine="xlsxwriter") as w:
        relatorio.to_excel(w, index=False)

    enviados = (relatorio["status"] == "ENVIADO").sum()
    erros    = (relatorio["status"] == "ERRO").sum()
    msg = (f"Processo concluído.\n"
           f"Enviados: {enviados}\nErros: {erros}\nTotal previsto: {total_previsto}\n\n"
           f"Relatórios:\n{rel_csv}\n{rel_xlsx}")

    print(msg)
    try:
        ctypes.windll.user32.MessageBoxW(0, msg, "Envio de E-mails", 0x40)  # ícone informação
    except Exception:
        pass

if __name__ == "__main__":
    main()
    
