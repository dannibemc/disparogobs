import streamlit as st
import pandas as pd
import os
import re
from datetime import datetime, timedelta
import jinja2
import logging

# Configuração do logger
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def obter_caminho_base():
    caminhos_base = [
        os.path.expandvars(r"%USERPROFILE%\\OneDrive - Leverage Companhia Securitizadora\\Documentos - Leverage Portal"),
        os.path.expandvars(r"%USERPROFILE%\\Documentos - Leverage Portal"),
        os.path.expandvars(r"%USERPROFILE%\\Leverage Companhia Securitizadora\\Leverage Portal - Documentos")
    ]
    for caminho in caminhos_base:
        if os.path.exists(caminho):
            return caminho
    st.error("Caminho base não encontrado.")
    return None

def campo_valido(campo):
    if campo is None or (isinstance(campo, str) and not campo.strip()) or str(campo).strip().lower() == "nan":
        return False
    if isinstance(campo, (float, int)):
        return not pd.isna(campo)
    if isinstance(campo, pd.Timestamp):
        return not pd.isna(campo)
    return True

def extrair_destinatarios(email, dados_fixos):
    email_primario = email if pd.notna(email) and isinstance(email, str) else ""
    email_secundario = dados_fixos.get("E-MAILS", "")
    email_secundario = email_secundario if pd.notna(email_secundario) and isinstance(email_secundario, str) else ""
    raw = f"{email_primario},{email_secundario}"
    return [e.strip() for e in raw.split(",") if e.strip() and e.strip().lower() != "nan"]

def carregar_template_operacao(nome_aba, caminho_html):
    caminho_template = os.path.join(caminho_html, f"{nome_aba}.html")
    if not os.path.exists(caminho_template):
        st.warning(f"Template HTML não encontrado para {nome_aba}: {caminho_template}")
        return None, None, None, None
    try:
        with open(caminho_template, "r", encoding="utf-8") as f:
            template_str = f.read()
        exige_iteracao = template_exige_iteracao(template_str)
        vars_globais, vars_loop = detectar_variaveis_completas_template_jinja(template_str)
        return jinja2.Template(template_str), vars_globais, vars_loop, exige_iteracao
    except Exception as e:
        st.error(f"Erro ao carregar template {caminho_template}: {e}")
        return None, None, None, None

def aplicar_formatacao(valor, formato):
    try:
        if pd.isna(valor):
            return ""
        if formato == 'M':
            return formatar_valor_monetario(valor)
        elif formato == 'D':
            return formatar_data_ddmmaaaa(valor)
        elif isinstance(valor, (int, float)):
            return str(valor)
        return str(valor)
    except Exception as e:
        logging.error(f"Erro ao formatar valor: {e}")
        return ""

def formatar_data_ddmmaaaa(val):
    try:
        data = pd.to_datetime(val, errors='coerce', dayfirst=True)
        if pd.notna(data):
            return data.strftime('%d/%m/%Y')
        return val
    except ValueError:
        logging.error(f"Erro ao converter para data: {val}")
        return val

def formatar_valor_monetario(val):
    try:
        valor_formatado = f"{float(val):,.2f}".replace(",", "v").replace(".", ",").replace("v", ".")
        return f"R$ {valor_formatado}"
    except ValueError:
        logging.error(f"Erro ao formatar valor monetário: {val}")
        return val

def template_exige_iteracao(template_str):
    return "{{ series" in template_str or "{% for" in template_str

def detectar_variaveis_completas_template_jinja(template_str):
    env = jinja2.Environment()
    parsed_content = env.parse(template_str)
    vars_globais = jinja2.meta.find_undeclared_variables(parsed_content)
    blocos_for = re.findall(r"{%\s*for\s+(\w+)\s+in\s+(\w+)\s*%}(.*?){%\s*endfor\s*%}", template_str, flags=re.DOTALL)
    variaveis_loop = set()
    for var_loop, lista, bloco in blocos_for:
        matches = re.findall(r"{{\s*" + re.escape(var_loop) + r"\.([\w_]+)\s*}}", bloco)
        variaveis_loop.update(matches)
    return vars_globais, variaveis_loop

def enviar_email_outlook(assunto, destinatarios, corpo_html, anexos=None, remetente="gestao@leveragesec.com.br"):
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        email = outlook.CreateItem(0)
        email.To = ""
        email.BCC = "; ".join(destinatarios)
        email.Subject = assunto
        email.SentOnBehalfOfName = remetente
        email.HTMLBody = corpo_html
        if anexos:
            for anexo in anexos:
                try:
                    if os.path.exists(anexo):
                        email.Attachments.Add(anexo)
                        logging.info(f"Anexo adicionado: {anexo}")
                    else:
                        logging.warning(f"Anexo especificado, mas não encontrado: {anexo}")
                except Exception as e:
                    logging.error(f"Erro ao adicionar anexo {anexo}: {e}")
        else:
            logging.info("Nenhum anexo especificado.")
        try:
            email.Send()
            logging.info(f"E-mail enviado para: {destinatarios}")
        except Exception as e:
            logging.error(f"Erro ao enviar e-mail para {destinatarios}: {e}")
    except Exception as e:
        logging.error(f"Erro ao enviar email via Outlook: {e}")
        st.error("Erro ao enviar email via Outlook. Verifique o log para detalhes.")

def deve_enviar_email(data_vencimento, data_hoje, dias_notificacao, recorrencia=None, dias_semana_validos=None):
    datas_notificacao = [data_vencimento - timedelta(days=d) for d in dias_notificacao if d >= 0]

    # Ajuste para segunda-feira
    if data_vencimento.weekday() == 0 and 2 in dias_notificacao:
        datas_notificacao.append(data_vencimento - timedelta(days=4))

    # Lógica de recorrência
    if recorrencia and data_hoje > data_vencimento and (data_hoje - data_vencimento).days % recorrencia == 0:
        datas_notificacao.append(data_hoje)

    # Verifica o dia da semana
    if dias_semana_validos and data_hoje.weekday() not in dias_semana_validos:
        return False

    return data_hoje in datas_notificacao

def preparar_series(grupo, campos_necessarios):
    grupo = grupo.sort_values(by="SERIE")
    colunas_disponiveis = {col.upper(): col for col in grupo.columns}
    series = []
    for _, linha in grupo.iterrows():
        numero = str(linha.get("SERIE", "")).strip()
        serie_formatada = f"{numero}ª" if numero.isdigit() else numero
        dados_serie = {"serie": serie_formatada}
        for campo in campos_necessarios:
            nome_base, _, sufixo = campo.rpartition('_')
            formato = sufixo.upper() if sufixo.upper() in {'M', 'D', 'S'} else None
            nome_busca = nome_base if formato else campo
            coluna_compat = next((colunas_disponiveis[c] for c in colunas_disponiveis if c.lower().startswith(nome_busca.lower())), None)
            valor = linha.get(coluna_compat, "") if coluna_compat else ""
            valor_formatado = aplicar_formatacao(valor, formato)
            dados_serie[campo] = valor_formatado if campo_valido(valor_formatado) else ""
        series.append(dados_serie)
    return series

def processar_emails(caminho_excel, caminho_html, caminho_anexos):
    try:
        planilhas = pd.read_excel(caminho_excel, sheet_name=None)
    except FileNotFoundError:
        st.error(f"Arquivo Excel não encontrado: {caminho_excel}")
        return 0, 0

    dados_operacoes = planilhas.get("dados das operações", pd.DataFrame())
    if not dados_operacoes.empty:
        dados_operacoes.columns = pd.Index([str(col).strip().upper() for col in dados_operacoes.columns])

    data_hoje = datetime.today().date()
    total_enviados = 0
    total_ignorados = 0

    log_area = st.empty()
    log_text = ""

    def append_log(message):
        nonlocal log_text
        log_text += f"{message}\n"
        log_area.text_area("Log:", log_text, height=300)

    for nome_aba, df in planilhas.items():
        if nome_aba.strip().lower() in ["dados das operações", "readme", "read-me"]:
            continue
        append_log(f"\n🔄 Processando operação: {nome_aba}")
        if df.empty:
            append_log(f"Aba {nome_aba} vazia. Ignorada.")
            total_ignorados += 1
            continue

        df.columns = pd.Index([str(col).strip().upper() for col in df.columns])

        template, vars_globais, vars_loop, exige_iteracao = carregar_template_operacao(nome_aba, caminho_html)
        if template is None:
            total_ignorados += 1
            continue
        if exige_iteracao and ("ID" not in df.columns or "SERIE" not in df.columns):
            append_log(f"Template exige iteração, mas colunas ID e SERIE ausentes em {nome_aba}. Ignorado.")
            total_ignorados += 1
            continue

        dados_fixos_df = dados_operacoes[dados_operacoes["OPERACAO"].str.upper() == nome_aba.upper()]
        if dados_fixos_df.empty:
            append_log(f"Operação {nome_aba} não encontrada na aba 'dados das operações'. Ignorado.")
            total_ignorados += 1
            continue
        dados_fixos = dados_fixos_df.iloc[0].to_dict()

        email_remetente = dados_fixos.get("E-MAILS_LEVERAGE", "").strip()
        if not campo_valido(email_remetente):
            append_log(f"{nome_aba}: remetente ('E-MAILS_LEVERAGE') ausente ou inválido. Ignorado.")
            total_ignorados += 1
            continue

        dias_raw = dados_fixos.get("DIAS_COBRANÇA", "")
        try:
            dias_notificacao = [int(d.strip()) for d in str(dias_raw).split(";") if d.strip().lstrip("-").isdigit()]
        except ValueError:
            append_log(f"{nome_aba}: valor inválido na coluna 'DIAS_COBRANÇA'. Ignorado.")
            total_ignorados += 1
            continue

        recorrencia = None
        try:
            valor = dados_fixos.get("RECORRENCIA", "")
            if campo_valido(valor):
                recorrencia = int(float(valor))
        except ValueError:
            append_log(f"{nome_aba}: valor inválido na coluna 'RECORRENCIA'.")

        dias_semana_validos = []
        try:
            dia_semana_raw = dados_fixos.get("DIA_DA_SEMANA", "")
            dias_semana_validos = [int(float(parte.strip())) for parte in str(dia_semana_raw).split(";") if parte.strip()]
        except ValueError:
            append_log(f"{nome_aba}: valor inválido na coluna 'DIA_DA_SEMANA'.")

        agrupador = df.groupby("ID") if exige_iteracao else df.iterrows()
        for k, grupo in agrupador:
            grupo = grupo if exige_iteracao else pd.DataFrame([grupo])
            linha_exemplo = grupo.iloc[0]
            id_referencia = f"ID {k}" if exige_iteracao else f"linha {grupo.index[0] + 2}"

            assunto = linha_exemplo.get("ASSUNTO_EMAIL", "")
            if not campo_valido(assunto):
                append_log(f"{id_referencia}: assunto do e-mail não informado. Ignorado.")
                total_ignorados += 1
                continue

            todos_emails = sum([extrair_destinatarios(linha.get("EMAIL", ""), dados_fixos) for _, linha in grupo.iterrows()], [])
            emails_unicos = list({e.lower(): e for e in todos_emails if campo_valido(e)}.values())
            if not emails_unicos:
                append_log(f"{id_referencia}: sem destinatários válidos. Ignorado.")
                total_ignorados += 1
                continue

            if all(str(linha.get("PAGAMENTO_REALIZADO", "")).strip().lower() == "sim" for _, linha in grupo.iterrows()):
                append_log(f"{id_referencia}: pagamento já realizado. Ignorado.")
                total_ignorados += 1
                continue

            if not exige_iteracao:
                coluna_venc = next((col for col in grupo.columns if col.upper().startswith("DATA_VENCIMENTO")), None)
                if coluna_venc:
                    try:
                        data_venc = pd.to_datetime(grupo.iloc[0][coluna_venc], dayfirst=True).date()
                        if not deve_enviar_email(data_venc, data_hoje, dias_notificacao, recorrencia, dias_semana_validos):
                            append_log(f"{id_referencia}: fora do calendário. Ignorado.")
                            total_ignorados += 1
                            continue
                    except (ValueError, TypeError) as e:
                        append_log(f"{id_referencia}: erro ao processar data de vencimento: {e}. Ignorado.")
                        total_ignorados += 1
                        continue

            coluna_venc = next((col for col in grupo.columns if col.upper().startswith("DATA_VENCIMENTO")), None)
            if not coluna_venc:
                append_log(f"{id_referencia}: coluna de vencimento não encontrada. Ignorado.")
                total_ignorados += 1
                continue

            contexto = {}
            for var in vars_globais:
                if var.startswith("s."): continue
                nome_base, _, sufixo = var.rpartition('_')
                formato = sufixo.upper() if sufixo.upper() in {'M', 'D', 'S'} else None
                nome_busca = nome_base if formato else var
                col_match = next((col for col in linha_exemplo.index if col.strip().lower() == var.lower()), None)
                if not col_match:
                    col_match = next((col for col in linha_exemplo.index if col.strip().lower().startswith(nome_busca.lower())), None)
                if not
