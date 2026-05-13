import streamlit as st
import tempfile, os
from conversor_portaria import converter

st.set_page_config(page_title="Conversor de Portaria - MTE", page_icon="📄")

st.title("📄 Conversor de Portaria - MTE")
st.write("Envie um arquivo **.docx** para gerar o HTML no padrão do MTE.")

arquivo = st.file_uploader("Escolha o arquivo .docx", type=["docx"])

if arquivo:
    if st.button("Converter"):
        with st.spinner("Convertendo..."):
            try:
                with tempfile.TemporaryDirectory() as tmp:
                    entrada = os.path.join(tmp, "portaria.docx")
                    saida   = os.path.join(tmp, "portaria.html")
                    with open(entrada, "wb") as f:
                        f.write(arquivo.read())
                    converter(entrada, saida)
                    with open(saida, "rb") as f:
                        html = f.read()
                nome_saida = arquivo.name.replace(".docx", ".html")
                st.success("✅ Conversão concluída!")
                st.download_button(
                    label="⬇️ Baixar HTML",
                    data=html,
                    file_name=nome_saida,
                    mime="text/html"
                )
            except Exception as e:
                st.error(f"Erro: {e}")
