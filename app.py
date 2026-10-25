import streamlit as st
import os
import subprocess
import tempfile
from datetime import datetime
from docxtpl import DocxTemplate, InlineImage
from docx.shared import Cm

st.set_page_config(page_title="Gerador de Laudos - NBR 14653", page_icon="🏢", layout="wide")

TABELA_ROSS = {
    2:   {'A': 1.02,  'B': 1.05,  'C': 3.51,  'D': 9.03,  'E': 18.90, 'F': 39.30, 'G': 53.10, 'H': 75.40},
    4:   {'A': 2.08,  'B': 2.11,  'C': 4.55,  'D': 10.00, 'E': 19.80, 'F': 34.60, 'G': 53.60, 'H': 75.70},
    6:   {'A': 3.18,  'B': 3.21,  'C': 5.62,  'D': 11.00, 'E': 20.70, 'F': 35.30, 'G': 54.10, 'H': 76.00},
    8:   {'A': 4.32,  'B': 4.35,  'C': 6.73,  'D': 12.10, 'E': 21.60, 'F': 36.10, 'G': 54.60, 'H': 76.30},
    10:  {'A': 5.50,  'B': 5.53,  'C': 7.88,  'D': 13.20, 'E': 22.60, 'F': 36.90, 'G': 55.20, 'H': 76.60},
    12:  {'A': 6.72,  'B': 6.75,  'C': 9.07,  'D': 14.30, 'E': 23.60, 'F': 37.70, 'G': 55.80, 'H': 76.90},
    14:  {'A': 7.98,  'B': 8.01,  'C': 10.30, 'D': 15.40, 'E': 24.60, 'F': 38.50, 'G': 56.40, 'H': 77.20},
    16:  {'A': 9.28,  'B': 9.31,  'C': 11.60, 'D': 16.60, 'E': 25.70, 'F': 39.40, 'G': 57.00, 'H': 77.50},
    18:  {'A': 10.60, 'B': 10.60, 'C': 12.90, 'D': 17.80, 'E': 26.80, 'F': 40.30, 'G': 57.60, 'H': 77.80},
    20:  {'A': 12.00, 'B': 12.00, 'C': 14.20, 'D': 19.10, 'E': 27.90, 'F': 41.80, 'G': 58.30, 'H': 78.20},
    22:  {'A': 13.40, 'B': 13.40, 'C': 15.60, 'D': 20.40, 'E': 29.10, 'F': 42.20, 'G': 59.00, 'H': 78.50},
    24:  {'A': 14.90, 'B': 14.90, 'C': 17.00, 'D': 21.80, 'E': 30.30, 'F': 43.10, 'G': 59.60, 'H': 78.90},
    26:  {'A': 16.40, 'B': 16.40, 'C': 18.50, 'D': 23.10, 'E': 31.50, 'F': 44.10, 'G': 60.40, 'H': 79.30},
    28:  {'A': 17.90, 'B': 17.90, 'C': 20.00, 'D': 24.60, 'E': 32.80, 'F': 45.20, 'G': 61.10, 'H': 79.60},
    30:  {'A': 19.50, 'B': 19.50, 'C': 21.50, 'D': 26.00, 'E': 34.10, 'F': 46.20, 'G': 61.80, 'H': 80.00},
    32:  {'A': 21.10, 'B': 21.10, 'C': 23.10, 'D': 27.50, 'E': 35.40, 'F': 47.30, 'G': 62.60, 'H': 80.40},
    34:  {'A': 22.80, 'B': 22.80, 'C': 24.70, 'D': 29.00, 'E': 36.80, 'F': 48.40, 'G': 63.40, 'H': 80.80},
    36:  {'A': 24.50, 'B': 24.50, 'C': 26.40, 'D': 30.50, 'E': 38.10, 'F': 49.50, 'G': 64.20, 'H': 81.30},
    38:  {'A': 26.20, 'B': 26.20, 'C': 28.10, 'D': 32.20, 'E': 39.60, 'F': 50.70, 'G': 65.00, 'H': 81.70},
    40:  {'A': 28.00, 'B': 28.00, 'C': 29.90, 'D': 33.80, 'E': 41.00, 'F': 51.90, 'G': 65.90, 'H': 82.10},
    42:  {'A': 29.80, 'B': 29.80, 'C': 31.60, 'D': 35.50, 'E': 42.50, 'F': 53.10, 'G': 66.70, 'H': 82.60},
    44:  {'A': 31.70, 'B': 31.70, 'C': 33.40, 'D': 37.20, 'E': 44.00, 'F': 54.40, 'G': 67.60, 'H': 83.10},
    46:  {'A': 33.60, 'B': 33.60, 'C': 35.20, 'D': 38.90, 'E': 45.60, 'F': 55.60, 'G': 68.50, 'H': 83.50},
    48:  {'A': 35.50, 'B': 35.50, 'C': 37.10, 'D': 40.70, 'E': 47.20, 'F': 56.90, 'G': 69.40, 'H': 84.00},
    50:  {'A': 37.50, 'B': 37.50, 'C': 39.10, 'D': 42.60, 'E': 48.80, 'F': 58.20, 'G': 70.40, 'H': 84.50},
    52:  {'A': 39.50, 'B': 39.50, 'C': 41.00, 'D': 44.40, 'E': 50.50, 'F': 59.60, 'G': 71.30, 'H': 85.00},
    54:  {'A': 41.60, 'B': 41.60, 'C': 43.00, 'D': 46.30, 'E': 52.10, 'F': 61.00, 'G': 72.30, 'H': 85.50},
    56:  {'A': 43.70, 'B': 43.70, 'C': 45.10, 'D': 48.20, 'E': 53.90, 'F': 62.40, 'G': 73.30, 'H': 86.00},
    58:  {'A': 45.80, 'B': 45.80, 'C': 47.20, 'D': 50.20, 'E': 55.60, 'F': 63.80, 'G': 74.30, 'H': 86.60},
    60:  {'A': 48.00, 'B': 48.00, 'C': 49.30, 'D': 52.20, 'E': 57.40, 'F': 65.30, 'G': 75.30, 'H': 87.10},
    62:  {'A': 50.20, 'B': 50.20, 'C': 51.50, 'D': 54.20, 'E': 59.20, 'F': 66.70, 'G': 76.40, 'H': 87.70},
    64:  {'A': 52.50, 'B': 52.50, 'C': 53.70, 'D': 56.30, 'E': 61.10, 'F': 68.30, 'G': 77.50, 'H': 88.20},
    66:  {'A': 54.80, 'B': 54.80, 'C': 55.90, 'D': 58.40, 'E': 63.00, 'F': 69.80, 'G': 78.60, 'H': 88.80},
    68:  {'A': 57.10, 'B': 57.10, 'C': 58.20, 'D': 60.60, 'E': 64.90, 'F': 71.40, 'G': 79.70, 'H': 89.40},
    70:  {'A': 59.50, 'B': 59.50, 'C': 60.50, 'D': 62.80, 'E': 66.80, 'F': 72.90, 'G': 80.80, 'H': 90.00},
    72:  {'A': 61.90, 'B': 61.90, 'C': 62.90, 'D': 65.00, 'E': 68.80, 'F': 74.60, 'G': 81.90, 'H': 90.60},
    74:  {'A': 64.40, 'B': 64.40, 'C': 65.30, 'D': 67.30, 'E': 70.80, 'F': 76.20, 'G': 76.20, 'H': 91.20},
    76:  {'A': 66.90, 'B': 66.90, 'C': 67.70, 'D': 69.60, 'E': 72.90, 'F': 77.90, 'G': 84.30, 'H': 91.80},
    78:  {'A': 69.40, 'B': 69.40, 'C': 70.20, 'D': 71.90, 'E': 74.90, 'F': 79.60, 'G': 85.50, 'H': 92.40},
    80:  {'A': 72.00, 'B': 72.00, 'C': 72.70, 'D': 74.30, 'E': 77.10, 'F': 81.30, 'G': 86.70, 'H': 93.10},
    82:  {'A': 74.60, 'B': 74.60, 'C': 75.30, 'D': 76.70, 'E': 79.20, 'F': 83.00, 'G': 88.00, 'H': 93.70},
    84:  {'A': 77.30, 'B': 77.30, 'C': 77.80, 'D': 79.10, 'E': 81.40, 'F': 84.80, 'G': 89.20, 'H': 94.40},
    86:  {'A': 80.00, 'B': 80.00, 'C': 80.50, 'D': 81.60, 'E': 83.60, 'F': 86.60, 'G': 90.50, 'H': 95.00},
    88:  {'A': 82.70, 'B': 82.70, 'C': 83.20, 'D': 84.10, 'E': 85.80, 'F': 88.50, 'G': 91.80, 'H': 95.70},
    90:  {'A': 85.50, 'B': 85.50, 'C': 85.90, 'D': 86.70, 'E': 88.10, 'F': 90.30, 'G': 93.10, 'H': 96.40},
    92:  {'A': 88.30, 'B': 88.30, 'C': 88.60, 'D': 89.30, 'E': 90.40, 'F': 92.20, 'G': 94.50, 'H': 97.10},
    94:  {'A': 91.20, 'B': 91.20, 'C': 91.40, 'D': 91.90, 'E': 92.80, 'F': 94.10, 'G': 95.80, 'H': 97.80},
    96:  {'A': 94.10, 'B': 94.10, 'C': 94.20, 'D': 94.60, 'E': 95.10, 'F': 96.00, 'G': 97.20, 'H': 98.50},
    98:  {'A': 97.00, 'B': 97.00, 'C': 97.10, 'D': 97.30, 'E': 97.60, 'F': 98.00, 'G': 98.60, 'H': 99.30},
    100:{'A': 100.0, 'B': 100.0, 'C': 100.0, 'D': 100.0, 'E': 100.0, 'F': 100.0, 'G': 100.0, 'H': 100.0}
}

def calcular_depreciacao(idade, vida, codigo):
    pct_vida = (idade / vida) * 100.0
    linha = int(round(pct_vida / 2.0) * 2)
    linha = min(max(linha, 2), 100)
    cod = codigo.strip().upper()
    k_pct = TABELA_ROSS[linha].get(cod, 26.4)
    fator = 1.0 - (k_pct / 100.0)
    return pct_vida, k_pct, fator

def arredondar_nbr(valor_exato):
    if valor_exato <= 0: return valor_exato
    valor_arred = round(valor_exato, -3)
    if abs(valor_arred - valor_exato) / valor_exato > 0.01:
        valor_arred = round(valor_exato, -2)
    return valor_arred

def fmt_br(val):
    return f"{val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

st.title("🏢 Gerador de Laudos de Avaliação de Imóveis")
st.caption("Conforme ABNT NBR 14653-2 & Método Evolutivo Ross-Heidecke")

formato_saida = st.radio("📄 Formato de Download Desejado:", ["PDF (.pdf)", "Word Editável (.docx)"], horizontal=True)

with st.form("laudo_form"):
    st.header("1. Identificação do Imóvel & Solicitante")
    col1, col2 = st.columns(2)
    with col1:
        nome_imovel = st.text_input("Nome do Imóvel / Uso", "Canil Municipal de Tarauacá")
        solicitante = st.text_input("Solicitante", "Prefeitura Municipal de Tarauacá")
        bairro = st.text_input("Bairro", "Centro")
        cidade_uf = st.text_input("Cidade / UF do Imóvel", "Tarauacá/AC")
    with col2:
        endereco_completo = st.text_input("Endereço Completo", "Rua A, nº 200, Bairro Centro")
        finalidade = st.text_input("Finalidade do Laudo", "fixação e repactuação de valor de locação mensal")
        data_vistoria = st.text_input("Data da Vistoria", "23 de Outubro de 2026")
        data_referencia = st.text_input("Data de Referência", "Outubro de 2026")

    st.header("2. Caracterização Física & Depreciação")
    col3, col4, col5 = st.columns(3)
    with col3:
        area_terreno = st.number_input("Área do Terreno (m²)", value=480.0, step=1.0)
        vu_ajustado_terreno = st.number_input("VU Ajustado Terreno (R$/m²)", value=142.50, step=0.5)
    with col4:
        area_construida = st.number_input("Área Construída (m²)", value=80.0, step=1.0)
        valor_cub_m2 = st.number_input("Valor CUB de Referência (R$/m²)", value=2100.0, step=10.0)
    with col5:
        idade_aparente = st.number_input("Idade Aparente (Anos)", value=18, step=1)
        vida_util_anos = st.number_input("Vida Útil de Referência (Anos)", value=50, step=1)

    padrao_cub_nome = st.selectbox("Padrão Construtivo / CUB", [
        "Galpão Industrial (GI)",
        "Residencial Unifamiliar Padrão Normal (R1-N)",
        "Comercial Salas e Lojas (CSL-8)"
    ])

    OPCOES_CONSERVACAO = {
        "A": "A - Novo / Excelente (Sem vícios)",
        "B": "B - Entre novo e regular (Pintura/Superficial)",
        "C": "C - Regular (Necessita reparos simples - telhado/instalações)",
        "D": "D - Entre regular e reparos simples",
        "E": "E - Necessita reparos simples",
        "F": "F - Reparos importantes",
        "G": "G - Reparos estruturais (Semi-ruína)",
        "H": "H - Ruína / Sem valor comercial"
    }
    estado_conservacao_cod = st.selectbox("Estado de Conservação (Ross-Heidecke)", list(OPCOES_CONSERVACAO.keys()), index=2, format_func=lambda x: OPCOES_CONSERVACAO[x])

    st.header("3. Taxa de Aluguel & Responsável Técnico")
    col6, col7 = st.columns(2)
    with col6:
        taxa_remuneracao_mensal = st.number_input("Taxa de Remuneração Mensal (% a.m.)", value=0.64, step=0.01)
        fonte_taxa = st.text_input("Fonte da Taxa", "Índice FipeZap Comercial (Setembro/2026)")
        cidade_emissao = st.text_input("Cidade de Emissão", "Tarauacá - AC")
    with col7:
        nome_engenheiro = st.text_input("Nome do Engenheiro", "Eng. Civil Fernando Silva de Brito")
        qualificacao_engenheiro = st.text_input("Qualificação", "Engenheiro Civil")
        crea_engenheiro = st.text_input("CREA", "CREA/AC nº 12345")

    st.header("4. Anexos Fotográficos & Croqui")
    croqui_file = st.file_uploader("Croqui do Terreno (Opcional)", type=["png", "jpg", "jpeg"])
    
    st.subheader("Fotos da Vistoria (Mínimo 1, Máximo 10)")
    foto_files = []
    
    col_f1, col_f2 = st.columns(2)
    for i in range(1, 11):
        target_col = col_f1 if i % 2 != 0 else col_f2
        with target_col:
            req_label = " (OBRIGATÓRIA)" if i == 1 else " (Opcional)"
            f_upload = st.file_uploader(f"Foto {i}{req_label}", type=["png", "jpg", "jpeg"], key=f"f_{i}")
            f_leg = st.text_input(f"Legenda Foto {i}", value=f"Vista da área {i}", key=f"l_{i}")
            if f_upload:
                foto_files.append((f_upload, f_leg))

    submit = st.form_submit_button("📄 Gerar Laudo de Avaliação", type="primary")

if submit:
    if not foto_files:
        st.error("⚠️ Atenção: É obrigatório anexar pelo menos 1 Foto de Vistoria (Foto 1) conforme NBR 14653.")
    else:
        with st.spinner("Calculando NBR 14653 e gerando o laudo..."):
            valor_total_terreno_vt = area_terreno * vu_ajustado_terreno
            cb0 = area_construida * valor_cub_m2
            pct_vida, k_pct, fator_dep = calcular_depreciacao(idade_aparente, vida_util_anos, estado_conservacao_cod)
            valor_benfeitoria_cb = cb0 * fator_dep
            custo_fisico_total = valor_total_terreno_vt + valor_benfeitoria_cb
            valor_patrimonial_arredondado = arredondar_nbr(custo_fisico_total)
            valor_aluguel_mensal = valor_patrimonial_arredondado * (taxa_remuneracao_mensal / 100.0)
            limite_inf = valor_aluguel_mensal * 0.85
            limite_sup = valor_aluguel_mensal * 1.15

            caminho_modelo = os.path.join('modelos', 'LAUTO AUTOMATICO rev4 - Copy.docx')
            doc = DocxTemplate(caminho_modelo)

            croqui_img = InlineImage(doc, croqui_file, width=Cm(15.0)) if croqui_file else None

            fotos_processadas = []
            for f_up, f_leg in foto_files:
                fotos_processadas.append({
                    "imagem": InlineImage(doc, f_up, width=Cm(15.0)),
                    "legenda": f_leg
                })

            MESES = ["janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"]
            hoje = datetime.now()
            data_extenso = f"{cidade_emissao}, {hoje.day} de {MESES[hoje.month-1]} de {hoje.year}"

            contexto = {
                "NOME_IMOVEL_USO": nome_imovel,
                "DESCRICAO_USO_IMOVEL": f"imóvel urbano destinado às instalações de {nome_imovel}",
                "BAIRRO_IMOVEL": bairro,
                "CIDADE_UF_IMOVEL": cidade_uf,
                "ENDERECO_COMPLETO_IMOVEL": endereco_completo,
                "NOME_SOLICITANTE": solicitante,
                "DATA_VISTORIA": data_vistoria,
                "DATA_REFERENCIA": data_referencia,
                "FINALIDADE_LAUDO": finalidade,

                "AREA_TERRENO": fmt_br(area_terreno),
                "AREA_CONSTRUIDA": fmt_br(area_construida),
                "PADRAO_CUB_NOME": padrao_cub_nome,
                "IDADE_APARENTE": str(idade_aparente),
                "ESTADO_CONSERVACAO": OPCOES_CONSERVACAO[estado_conservacao_cod],

                "VU_AJUSTADO_TERRENO": fmt_br(vu_ajustado_terreno),
                "VALOR_TOTAL_TERRENO_VT": fmt_br(valor_total_terreno_vt),
                "VALOR_CUB_M2": fmt_br(valor_cub_m2),
                "VALOR_INICIAL_BENFEITORIA_CB0": fmt_br(cb0),
                "VIDA_UTIL_ANOS": str(vida_util_anos),
                "PERCENTUAL_VIDA_DECORRIDA": fmt_br(pct_vida),
                "ESTADO_CONSERVACAO_NOME": OPCOES_CONSERVACAO[estado_conservacao_cod].split("-")[1].strip(),
                "ESTADO_CONSERVACAO_CODIGO": estado_conservacao_cod,
                "COEFICIENTE_K_ROSS_HEIDECKE": fmt_br(k_pct),
                "FATOR_DEPRECIACAO": f"{fator_dep:.4f}".replace(".", ","),
                "VALOR_BENFEITORIA_CB": fmt_br(valor_benfeitoria_cb),

                "CUSTO_FISICO_TOTAL": fmt_br(custo_fisico_total),
                "VALOR_PATRIMONIAL_ARREDONDADO": fmt_br(valor_patrimonial_arredondado),
                "VALOR_PATRIMONIAL_POR_EXTENSO": "Valor calculado conforme arredondamento NBR 14653-1",

                "TAXA_REMUNERACAO_MENSAL": fmt_br(taxa_remuneracao_mensal),
                "FONTE_TAXA_REMUNERACAO": fonte_taxa,
                "VALOR_ALUGUEL_MENSAL": fmt_br(valor_aluguel_mensal),
                "VALOR_ALUGUEL_POR_EXTENSO": "Valor calculado de locação mensal",
                "LIMITE_INFERIOR_ALUGUEL": fmt_br(limite_inf),
                "LIMITE_SUPERIOR_ALUGUEL": fmt_br(limite_sup),

                "CIDADE_UF_EMISSAO": cidade_emissao,
                "DATA_EMISSAO_COMPLETA": data_extenso,
                "NOME_ENGENHEIRO": nome_engenheiro,
                "QUALIFICACAO_ENGENHEIRO": qualificacao_engenheiro,
                "CREA_ENGENHEIRO": crea_engenheiro,

                "croqui_imagem": croqui_img,
                "croqui_legenda": "Croqui esquemático perimetral do imóvel.",
                "fotos": fotos_processadas
            }

            doc.render(contexto)

            with tempfile.TemporaryDirectory() as tmpdir:
                docx_path = os.path.join(tmpdir, f"Laudo_{bairro}.docx")
                doc.save(docx_path)

                if "PDF" in formato_saida:
                    try:
                        subprocess.run([
                            "libreoffice", "--headless", "--convert-to", "pdf",
                            docx_path, "--outdir", tmpdir
                        ], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                        
                        pdf_path = os.path.join(tmpdir, f"Laudo_{bairro}.pdf")
                        with open(pdf_path, "rb") as f:
                            pdf_bytes = f.read()

                        st.success("✅ Laudo em PDF gerado com sucesso!")
                        st.download_button(
                            label="⬇️ Baixar Laudo em PDF (.pdf)",
                            data=pdf_bytes,
                            file_name=f"Laudo_Avaliacao_{bairro}.pdf",
                            mime="application/pdf"
                        )
                    except Exception as err:
                        st.warning("⚠️ Conversão direta em PDF indisponível. Baixando em Word (.docx)...")
                        with open(docx_path, "rb") as f:
                            docx_bytes = f.read()
                        st.download_button(
                            label="⬇️ Baixar Laudo em Word (.docx)",
                            data=docx_bytes,
                            file_name=f"Laudo_Avaliacao_{bairro}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                        )
                else:
                    with open(docx_path, "rb") as f:
                        docx_bytes = f.read()

                    st.success("✅ Laudo em Word gerado com sucesso!")
                    st.download_button(
                        label="⬇️ Baixar Laudo em Word (.docx)",
                        data=docx_bytes,
                        file_name=f"Laudo_Avaliacao_{bairro}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
