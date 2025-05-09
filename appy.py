import streamlit as st
import numpy as np
import matplotlib.pyplot as plt

st.title("Simulações de Física com Streamlit")

st.sidebar.header("Controles")

velocidade_inicial = st.sidebar.slider("Velocidade Inicial (m/s)", 0.1, 100.0, 50.0)
angulo_graus = st.sidebar.slider("Ângulo (graus)", 0, 90, 45)
angulo_radianos = np.deg2rad(angulo_graus)
gravidade = 9.81

# Cálculos do movimento de projétil
tempo_total = (2 * velocidade_inicial * np.sin(angulo_radianos)) / gravidade
tempos = np.linspace(0, tempo_total, 100)
posicao_x = velocidade_inicial * np.cos(angulo_radianos) * tempos
posicao_y = velocidade_inicial * np.sin(angulo_radianos) * tempos - 0.5 * gravidade * tempos**2

# Plotagem da trajetória
fig, ax = plt.subplots()
ax.plot(posicao_x, posicao_y)
ax.set_xlabel("Distância Horizontal (m)")
ax.set_ylabel("Altura (m)")
ax.set_title("Trajetória do Projétil")
ax.grid(True)
st.pyplot(fig)

# Exibição de informações adicionais
st.subheader("Resultados:")
st.write(f"Tempo total de voo: {tempo_total:.2f} s")
st.write(f"Alcance máximo horizontal: {np.max(posicao_x):.2f} m")
st.write(f"Altura máxima atingida: {np.max(posicao_y):.2f} m")