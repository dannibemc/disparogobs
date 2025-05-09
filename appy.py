import streamlit as st
import pandas as pd
import os
import re
from datetime import datetime, timedelta
import jinja2
import logging
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders

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

def enviar_email_smtp(assunto, destinatarios, corpo_html, remetente, senha, servidor_smtp, porta_smtp, anexos=None):
    msg = MIMEMultipart()
    msg['From'] = remetente
    msg['To'] = ", ".join(destinatarios)
    msg['Subject'] = assunto

    msg.attach(MIMEText(corpo_html, 'html'))  # Use 'html' for HTML content

    if anexos:
        for anexo in anexos:
            try:
                with open(anexo, "rb") as arquivo_anexo:
                    parte_anexo = MIMEBase('application', 'octet-stream')
                    parte_anexo.set_payload(arquivo_anexo.read())
                    encoders.encode_base64(parte_anexo)
                    parte_anexo.add_header('Content-Disposition', f'attachment; filename="{os.path.basename(anexo)}"')
                    msg.attach(parte_anexo)
            except Exception as e:
                logging.error(f"Erro ao anexar {anexo}: {e}")

    try:
        servidor = smtplib.SMTP(servidor_smtp, porta_smtp)
        servidor.starttls()  # Secure the connection
        servidor.login(remetente, senha)
        servidor.sendmail(remetente, destinatarios, msg.as_string())
        servidor.quit()
        logging.info(f"E-mail enviado para: {destinatarios}")
    except Exception as e:
        logging.error(f"Erro ao enviar e-mail via SMTP: {e}")
        st.error(f"Erro ao enviar e-mail: {e}")

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

def processar_emails(caminho_excel, caminho_html, caminho_anexos, remetente, senha, servidor_smtp, porta_smtp):
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
        assunto = dados_fixos.get("ASSUNTO", f"Notificação de Operação - {nome_aba}")
        dias_notificacao = [int(d) for d in str(dados_fixos.get("DIAS_NOTIFICACAO", "0")).split(",") if str(d).strip().isdigit()]
        recorrencia = dados_fixos.get("RECORRENCIA")
        recorrencia = int(recorrencia) if pd.notna(recorrencia) and str(recorrencia).strip().isdigit() else None
        dias_semana_validos_str = dados_fixos.get("DIAS_SEMANA_VALIDOS")
        dias_semana_validos = [int(d) for d in str(dias_semana_validos_str).split(",") if str(d).strip().isdigit()] if pd.notna(dias_semana_validos_str) else None

        grupos = df.groupby("ID") if exige_iteracao else [("", df)]

        for id_operacao, grupo in grupos:
            linha_principal = grupo.iloc[0]
            data_vencimento = linha_principal.get("DATA_VENCIMENTO")
            email = linha_principal.get("E-MAIL")

            if not campo_valido(data_vencimento) or not campo_valido(email):
                append_log(f"Dados insuficientes para ID {id_operacao} em {nome_aba}. Ignorado.")
                total_ignorados += 1
                continue

            data_vencimento = pd.to_datetime(data_vencimento).date()
            if not deve_enviar_email(data_vencimento, data_hoje, dias_notificacao, recorrencia, dias_semana_validos):
                append_log(f"Notificação para ID {id_operacao} em {nome_aba} não necessária hoje. Ignorado.")
                total_ignorados += 1
                continue

            destinatarios = extrair_destinatarios(email, dados_fixos)
            if not destinatarios:
                append_log(f"Sem destinatários válidos para ID {id_operacao} em {nome_aba}. Ignorado.")
                total_ignorados += 1
                continue

            contexto = dados_fixos.copy()
            for var in vars_globais:
                contexto[var] = aplicar_formatacao(dados_fixos.get(var), None)

            if exige_iteracao:
                contexto["series"] = preparar_series(grupo, vars_loop)
            else:
                for col in df.columns:
                    valor = df.iloc[0].get(col)
                    formato = None
                    contexto[col] = aplicar_formatacao(valor, formato)

            try:
                corpo_html = template.render(contexto)
            except Exception as e:
                append_log(f"Erro ao renderizar template para ID {id_operacao} em {nome_aba}: {e}")
                total_ignorados += 1
                continue

            try:
                enviar_email_smtp(assunto, destinatarios, corpo_html, remetente, senha, servidor_smtp, porta_smtp, caminho_anexos)
                total_enviados += 1
                append_log(f"E-mail enviado para ID {id_operacao} em {nome_aba} para: {', '.join(destinatarios)}")
            except Exception as e:
                append_log(f"Erro ao enviar e-mail para ID {id_operacao} em {nome_aba}: {e}")

    st.success(f"Processamento concluído. Emails enviados: {total_enviados} | Emails ignorados: {total_ignorados}")

def main():
    st.title("Processamento de E-mails Automatizado")

    caminho_excel = st.text_input("Caminho do Arquivo Excel:", "")
    caminho_html = st.text_input("Caminho da Pasta HTML:", "")
    caminho_anexos = st.text_input("Caminho da Pasta de Anexos (opcional):", "")

    # SMTP Settings (Get these from your email provider)
    remetente = st.text_input("Seu E-mail:", "")
    senha = st.text_input("Sua Senha/App Password:", "", type="password")  # Use type="password" for security
    servidor_smtp = st.text_input("Servidor SMTP:", "")
    porta_smtp = st.number_input("Porta SMTP:", value=587)  # Common port for TLS

    if st.button("Iniciar Processamento"):
        if caminho_excel and caminho_html and remetente and senha and servidor_smtp and porta_smtp:
            processar_emails(caminho_excel, caminho_html,
                             caminho_excel, caminho_html, caminho_anexos, remetente, senha, servidor_smtp, porta_smtp)
           else:
               for id_operacao, grupo in grupos:
               # ... outras linhas de código ...

               if exige_iteracao:
                   contexto["series"] = preparar_series(grupo, vars_loop)
               else:
                   # A linha abaixo (originalmente linha 300) deve estar indentada aqui
                   for col in df.columns:
                       valor = df.iloc[0].get(col)
                       formato = None
                       contexto[col] = aplicar_formatacao(valor, formato)

               try:
                   corpo_html = template.render(contexto)
               except Exception as e:
                   append_log(f"Erro ao renderizar template para ID {id_operacao} em {nome_aba}: {e}")
                   total_ignorados += 1
                   continue
   if __name__ == "__main__":
       main()
