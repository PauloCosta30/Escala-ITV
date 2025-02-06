import streamlit as st
import pandas as pd
import os
from datetime import datetime, timedelta
import io  # Para salvar o Excel na memória
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

FILE_NAME = "escala_lab.xlsx"
SALAS = ["Geral 1", "Geral 2", "Geral 3", "Geral 4", "Citometria - Bancada", "Geologia 2", "Geologia Micrótomo", "Cultivo A1", "Cultivo A2", "Cultivo A3/Fluxo", "Cultivo Subsolo 2"]

# Gerar as próximas 7 datas
def gerar_datas_proximos_7_dias():
    hoje = datetime.now()
    return [(hoje + timedelta(days=i)).strftime('%Y-%m-%d') for i in range(7)]

# Função para carregar ou criar o arquivo Excel
def load_data():
    if os.path.exists(FILE_NAME):
        return pd.read_excel(FILE_NAME)
    else:
        return pd.DataFrame(columns=["Data e Período"] + SALAS)

# Função para salvar os dados no Excel
def save_data(df):
    df.to_excel(FILE_NAME, index=False, engine="openpyxl")

# Função para enviar e-mail de confirmação
def enviar_email(nome, email, reservas):
    remetente = "paulo.henriquee30@gmail.com"  # Substitua pelo seu e-mail
    senha = "spkq lnax gulm vlsy"  # Substitua pela senha de app do Gmail
    destinatario = email
    
    mensagem = MIMEMultipart()
    mensagem['From'] = remetente
    mensagem['To'] = destinatario
    mensagem['Subject'] = "Confirmação de Preenchimento da Escala"

    corpo = f"Olá {nome},\n\nVocê preencheu a escala conforme os seguintes dados:\n"
    for reserva in reservas:
        corpo += f"- {reserva['Data e Período']} | Sala: {reserva['Sala']}\n"
    corpo += "\nObrigado!"

    mensagem.attach(MIMEText(corpo, 'plain'))
    
    try:
        servidor = smtplib.SMTP('smtp.gmail.com', 587)
        servidor.starttls()
        servidor.login(remetente, senha)
        servidor.send_message(mensagem)
        servidor.quit()
        st.success("E-mail de confirmação enviado com sucesso!")
    except Exception as e:
        st.error(f"Erro ao enviar e-mail: {e}")

# Função para verificar conflitos de reservas
def verificar_conflito(data_periodo, sala, escala):
    if data_periodo in escala["Data e Período"].values:
        valor_existente = escala.loc[escala["Data e Período"] == data_periodo, sala].values[0]
        if pd.notna(valor_existente) and valor_existente.strip() != "":
            return False  # Conflito detectado
    return True

# Função principal para preenchimento da escala
def preencher_escala(data_periodo, sala, nome, escala):
    if verificar_conflito(data_periodo, sala, escala):
        if data_periodo in escala["Data e Período"].values:
            escala.loc[escala["Data e Período"] == data_periodo, sala] = nome
        else:
            nova_linha = pd.Series([data_periodo] + [nome if col == sala else "" for col in SALAS], index=escala.columns)
            escala = pd.concat([escala, pd.DataFrame([nova_linha])], ignore_index=True)
        save_data(escala)
        return escala, True
    else:
        return escala, False

# Carregar ou criar a escala
escala = load_data()

st.title("Gerenciamento de Escala de Laboratório")

# Inputs do usuário
nome = st.text_input("Nome:")
email = st.text_input("E-mail:")
dias_selecionados = st.multiselect("Selecione os dias:", gerar_datas_proximos_7_dias())

reservas = []
if len(dias_selecionados) > 0:
    st.subheader("Configuração para cada dia e período selecionado")
    for dia in dias_selecionados:
        st.markdown(f"### Configuração para {dia}")
        for periodo in ["Manhã", "Tarde"]:
            st.markdown(f"#### Período: {periodo}")
            sala = st.selectbox(f"Sala para {dia} - {periodo}:", SALAS, key=f"sala_{dia}_{periodo}")
            if sala:
                reservas.append({"Data e Período": f"{dia} - {periodo}", "Sala": sala})

if st.button("Reservar"):
    if nome.strip() == "" or email.strip() == "":
        st.warning("Por favor, insira seu nome e e-mail.")
    elif len(reservas) == 0:
        st.warning("Por favor, selecione pelo menos um dia e configure o período e a sala.")
    else:
        conflito = False
        for reserva in reservas:
            escala, sucesso = preencher_escala(reserva["Data e Período"], reserva["Sala"], nome, escala)
            if not sucesso:
                st.error(f"Conflito detectado: {reserva['Data e Período']} na sala {reserva['Sala']} já foi reservada.")
                conflito = True
        if not conflito:
            enviar_email(nome, email, reservas)
            st.success("Reservas realizadas com sucesso!")

st.subheader("Escala Atual")
st.dataframe(escala)

# Criar um buffer de memória para salvar o Excel antes de baixar
output = io.BytesIO()
with pd.ExcelWriter(output, engine="openpyxl") as writer:
    escala.to_excel(writer, index=False)
output.seek(0)

st.download_button(
    label="Baixar Planilha",
    data=output,
    file_name="escala_lab.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

