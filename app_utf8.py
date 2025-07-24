import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import io

def planificar(procesos, personal, restricciones):
    restricciones_dict = dict(zip(restricciones['Restricción'], restricciones['Valor']))
    max_procesos = int(restricciones_dict.get('Máximo procesos por persona', 999))
    max_horas = int(restricciones_dict.get('Máximo horas por semana', 999))
    turnos_compatibles = restricciones_dict.get('Requiere turnos compatibles', 'No') == 'Sí'
    evitar_solapamiento = restricciones_dict.get('Evitar solapamiento de recursos', 'No') == 'Sí'

    personal['Horas restantes'] = personal['Horas disponibles/semana'].clip(upper=max_horas)
    personal['Procesos asignados'] = 0
    personal['Turno'] = personal['Turno'].fillna('Indistinto')
    procesos = procesos.sort_values(by=['Prioridad', 'Deadline'])

    asignaciones = []
    recursos_en_uso = set()

    for _, proceso in procesos.iterrows():
        candidatos = personal.copy()
        candidatos = candidatos[candidatos['Habilidades'].str.contains(proceso['Tipo Recurso'], na=False)]
        candidatos = candidatos[candidatos['Horas restantes'] >= proceso['Duración Estimada (hs)']]
        candidatos = candidatos[candidatos['Procesos asignados'] < max_procesos]

        if turnos_compatibles:
            turno_req = 'Mañana' if proceso['Prioridad'] <= 2 else 'Tarde'
            candidatos = candidatos[candidatos['Turno'] == turno_req]

        if evitar_solapamiento:
            candidatos = candidatos[~candidatos['Recursos disponibles'].apply(
                lambda x: any(herr in recursos_en_uso for herr in str(x).split(',')))]

        if not candidatos.empty:
            candidato = candidatos.sort_values(by='Procesos asignados').iloc[0]
            asignaciones.append({
                'Proceso': proceso['Nombre Proceso'],
                'Asignado a': candidato['Nombre'],
                'Horas del proceso': proceso['Duración Estimada (hs)'],
                'Prioridad': proceso['Prioridad'],
                'Deadline': proceso['Deadline']
            })

            idx = personal.index[personal['ID'] == candidato['ID']][0]
            personal.at[idx, 'Horas restantes'] -= proceso['Duración Estimada (hs)']
            personal.at[idx, 'Procesos asignados'] += 1

            if evitar_solapamiento:
                recursos_en_uso.update(str(candidato['Recursos disponibles']).split(','))

    df_asignaciones = pd.DataFrame(asignaciones)
    resumen = df_asignaciones.groupby('Asignado a').agg({
        'Proceso': 'count',
        'Horas del proceso': 'sum'
    }).reset_index()
    resumen.columns = ['Nombre', 'Cantidad de procesos asignados', 'Total de horas asignadas']

    reporte = pd.merge(
        personal[['Nombre', 'Horas disponibles/semana']],
        resumen,
        on='Nombre', how='left'
    ).fillna({'Cantidad de procesos asignados': 0, 'Total de horas asignadas': 0})

    return df_asignaciones, reporte

def generar_graficos(reporte, df_asignaciones):
    fig1, ax1 = plt.subplots(figsize=(10, 6))
    ax1.bar(reporte['Nombre'], reporte['Horas disponibles/semana'], label='Horas disponibles', alpha=0.5)
    ax1.bar(reporte['Nombre'], reporte['Total de horas asignadas'], label='Horas asignadas', alpha=0.8)
    ax1.set_title('Horas asignadas vs disponibles')
    ax1.legend()
    plt.xticks(rotation=45)
    st.pyplot(fig1)

    fig2, ax2 = plt.subplots(figsize=(6, 6))
    conteo = df_asignaciones['Asignado a'].value_counts()
    ax2.pie(conteo, labels=conteo.index, autopct='%1.1f%%', startangle=140)
    ax2.set_title('Distribución de procesos')
    st.pyplot(fig2)

st.title("Planificador de Procesos Inteligente")

archivo_excel = st.file_uploader("Subí el archivo Excel con las hojas: Procesos, Personal, Restricciones", type=["xlsx"])

if archivo_excel:
    xls = pd.ExcelFile(archivo_excel)
    procesos = pd.read_excel(xls, 'Procesos')
    personal = pd.read_excel(xls, 'Personal')
    restricciones = pd.read_excel(xls, 'Restricciones')

    if st.button("Generar asignaciones"):
        df_asignaciones, reporte = planificar(procesos, personal, restricciones)

        st.subheader("Asignaciones Generadas")
        st.dataframe(df_asignaciones)

        st.subheader("Resumen por Persona")
        st.dataframe(reporte)

        st.subheader("Visualización")
        generar_graficos(reporte, df_asignaciones)

        with io.BytesIO() as buffer:
            with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                df_asignaciones.to_excel(writer, sheet_name='Asignaciones', index=False)
                reporte.to_excel(writer, sheet_name='Resumen por Persona', index=False)
            st.download_button(
                label="Descargar resultados en Excel",
                data=buffer.getvalue(),
                file_name="resultado_planificacion.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
