import pandas as pd
from pptx import Presentation
from pptx.util import Pt
import requests
import google.generativeai as genai
import os
from dotenv import load_dotenv
import streamlit as st
from io import BytesIO

load_dotenv()

client = genai.configure(api_key=os.getenv('API_KEY'))

# Função para gerar resumo usando a API do Gemini
def gerar_resumo(livro):
    response = client.models.generate_content(
        model="gemini-2.0-flash",
        contents=f"Faça um resumo de no máximo 445 caracteres sobre o livro: {livro}",
    )
    return response.text

def gerar_frase_motivacional(livros):
    temas = ", ".join(livros)
    response = client.models.generate_content(
        model="gemini-2.0-flash",
        contents=f"Crie apenas uma frase curta e motivacional de no máximo 100 caracteres que incentive a leitura (não quero sugestões, apenas retorne a frase sem mais nada além disso. não precisa deixar em negrito, então não coloque asteriscos '*'), baseada nos seguintes livros: {temas}",
    )
    return response.text

# Função para baixar imagens a partir de URLs
def baixar_imagem(url, caminho_local):
    resposta = requests.get(url)
    if resposta.status_code == 200:
        with open(caminho_local, "wb") as arquivo:
            arquivo.write(resposta.content)
    else:
        raise Exception(f"Erro ao baixar a imagem: {url}")

def main():
    st.image("imagens\header books.jpg")
    st.title("Gerador de Template da biblioteca corporativa")

    # Inputs para os nomes dos livros
    st.subheader("Digite os nomes dos livros")
    livro1 = st.text_input("Nome do Livro 1")
    livro2 = st.text_input("Nome do Livro 2")
    livro3 = st.text_input("Nome do Livro 3")

    # Inputs para os links das imagens
    st.subheader("Insira os links das imagens dos livros")
    link_imagem1 = st.text_input("Link da Imagem do Livro 1")
    link_imagem2 = st.text_input("Link da Imagem do Livro 2")
    link_imagem3 = st.text_input("Link da Imagem do Livro 3")

    # Verifica se todos os campos foram preenchidos
    if livro1 and livro2 and livro3 and link_imagem1 and link_imagem2 and link_imagem3:
        livros = [livro1, livro2, livro3]
        links_imagens = [link_imagem1, link_imagem2, link_imagem3]

        # Baixar as imagens e salvar localmente
        caminhos_imagens = []
        for i, url in enumerate(links_imagens):
            caminho_local = f"imagem_{i + 1}.jpg"  # Nome do arquivo local
            try:
                baixar_imagem(url, caminho_local)
                caminhos_imagens.append(caminho_local)
            except Exception as e:
                st.error(f"Erro ao baixar a imagem {i + 1}: {e}")
                return

        # Gerar resumos e frase motivacional
        resumos = [gerar_resumo(livro) for livro in livros]
        frase_motivacional = gerar_frase_motivacional(livros)

        # Exibir prévia dos resumos e imagens
        st.subheader("Prévia do Template")

        # Exibir resumos
        st.write("**Resumos dos Livros:**")
        for i, resumo in enumerate(resumos):
            st.write(f"**Livro {i + 1}:** {resumo}")

        # Exibir imagens
        st.write("**Imagens dos Livros:**")
        cols = st.columns(3)
        for i, caminho in enumerate(caminhos_imagens):
            cols[i].image(caminho, caption=f"Imagem {i + 1}", use_column_width=True)

        # Exibir frase motivacional
        st.write("**Frase Motivacional:**")
        st.write(frase_motivacional)

        # Botão para gerar a apresentação final
        if st.button("Gerar Apresentação"):
            # Carregar o template da apresentação
            pptx_path = 'minha_apresentacao.pptx'
            prs = Presentation(pptx_path)

            # Função para substituir texto nos placeholders
            def substituir_texto(slide, antigo_texto, novo_texto, tamanho_fonte=Pt(14)):
                for shape in slide.shapes:
                    if shape.has_text_frame:
                        for para in shape.text_frame.paragraphs:
                            if antigo_texto in para.text:
                                para.text = para.text.replace(antigo_texto, novo_texto)
                                for run in para.runs:
                                    run.font.size = tamanho_fonte

            # Função para substituir imagens nos placeholders
            def substituir_imagem_por_nome(slide, nome_placeholder, nova_imagem):
                for shape in slide.shapes:
                    if shape.name == nome_placeholder:
                        left = shape.left
                        top = shape.top
                        width = shape.width
                        height = shape.height
                        slide.shapes.add_picture(nova_imagem, left, top, width, height)
                        sp = shape._element
                        sp.getparent().remove(sp)

            # Substituir textos e imagens no template
            for slide in prs.slides:
                substituir_texto(slide, 'texto1', resumos[0])
                substituir_texto(slide, 'texto2', resumos[1])
                substituir_texto(slide, 'texto3', resumos[2])
                substituir_texto(slide, 'texto4', frase_motivacional)
                substituir_imagem_por_nome(slide, 'imagem1', caminhos_imagens[0])
                substituir_imagem_por_nome(slide, 'imagem2', caminhos_imagens[1])
                substituir_imagem_por_nome(slide, 'imagem3', caminhos_imagens[2])

            # Salvar a apresentação modificada
            output_path = 'apresentacao_modificada.pptx'
            prs.save(output_path)

            # Disponibilizar o download da apresentação
            with open(output_path, "rb") as file:
                btn = st.download_button(
                    label="Baixar Apresentação Modificada",
                    data=file,
                    file_name="apresentacao_modificada.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )

            # Limpar as imagens baixadas
            for caminho in caminhos_imagens:
                os.remove(caminho)

            st.success("Apresentação gerada com sucesso!")

if __name__ == "__main__":
    main()