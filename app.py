import os
import pandas as pd
from pptx import Presentation
from pptx.util import Pt
import requests
import google.generativeai as genai
from dotenv import load_dotenv
import streamlit as st
from io import BytesIO
from PIL import Image
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed

# Carregar variáveis de ambiente
load_dotenv()

# Configurar o cliente do Gemini
try:
    genai.configure(api_key=os.getenv('API_KEY'))
except Exception as e:
    st.error(f"Erro ao configurar a API do Gemini: {e}")
    st.stop()

# Função para gerar resumo usando a API do Gemini
def gerar_resumo(livro):
    try:
        model = genai.GenerativeModel('gemini-2.0-flash-lite')
        response = model.generate_content(
            f"Faça um resumo de no máximo 445 caracteres sobre o livro: {livro}"
        )
        st.write(f"Resposta da API para '{livro}': {response.text}")  # Log da resposta
        return response.text
    except Exception as e:
        st.error(f"Erro ao gerar resumo: {e}")
        return ""

def gerar_frase_motivacional(livros):
    try:
        temas = ", ".join(livros)
        model = genai.GenerativeModel('gemini-2.0-flash-lite')
        response = model.generate_content(
            f"Crie apenas uma frase curta e motivacional de no máximo 100 caracteres que incentive a leitura (não quero sugestões, apenas retorne a frase sem mais nada além disso. não precisa deixar em negrito, então não coloque asteriscos '*'), baseada nos seguintes livros: {temas}"
        )
        st.write(f"Resposta da API para frase motivacional: {response.text}")  # Log da resposta
        return response.text
    except Exception as e:
        st.error(f"Erro ao gerar frase motivacional: {e}")
        return ""

# Função para gerar resumo com timeout
def gerar_resumo_com_timeout(livro):
    with ThreadPoolExecutor() as executor:
        future = executor.submit(gerar_resumo, livro)
        try:
            return future.result(timeout=10)  # Timeout de 10 segundos
        except Exception as e:
            st.error(f"Timeout ao gerar resumo para '{livro}': {e}")
            return ""

# Função para baixar imagens a partir de URLs
def baixar_imagem(url, caminho_local):
    try:
        resposta = requests.get(url, timeout=10)  # Timeout de 10 segundos
        if resposta.status_code == 200:
            with open(caminho_local, "wb") as arquivo:
                arquivo.write(resposta.content)
        else:
            raise Exception(f"Erro ao baixar a imagem: {url}")
    except Exception as e:
        st.error(f"Erro ao baixar a imagem: {e}")
        raise

# Função para converter slides do PowerPoint para JPEG
def pptx_para_jpeg(pptx_path, output_folder):
    try:
        prs = Presentation(pptx_path)
        imagens = []
        for i, slide in enumerate(prs.slides):
            slide_image_path = os.path.join(output_folder, f"slide_{i + 1}.jpg")
            slide_image = Image.new("RGB", (1280, 720), (255, 255, 255))  # Tamanho padrão para slides
            slide_image.save(slide_image_path)
            imagens.append(slide_image_path)
        return imagens
    except Exception as e:
        st.error(f"Erro ao converter slides para JPEG: {e}")
        return []

def main():
    st.title("Gerador de Apresentações de Livros")

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
        with st.spinner("Gerando resumos..."):
            resumos = [gerar_resumo_com_timeout(livro) for livro in livros]

        with st.spinner("Gerando frase motivacional..."):
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
            try:
                cols[i].image(caminho, caption=f"Imagem {i + 1}", use_column_width=True)
            except Exception as e:
                st.error(f"Erro ao exibir a imagem {i + 1}: {e}")

        # Exibir frase motivacional
        st.write("**Frase Motivacional:**")
        st.write(frase_motivacional)

        # Botão para gerar a apresentação final
        if st.button("Gerar Apresentação"):
            # Verificar se o template da apresentação existe
            pptx_path = 'minha_apresentacao.pptx'
            if not os.path.exists(pptx_path):
                st.error(f"Erro: O arquivo {pptx_path} não foi encontrado.")
                return

            # Carregar o template da apresentação
            try:
                prs = Presentation(pptx_path)
            except Exception as e:
                st.error(f"Erro ao carregar o template da apresentação: {e}")
                return

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
            output_pptx_path = 'apresentacao_modificada.pptx'
            try:
                prs.save(output_pptx_path)
            except Exception as e:
                st.error(f"Erro ao salvar a apresentação modificada: {e}")
                return

            # Converter a apresentação para JPEG
            output_folder = "slides_jpeg"
            os.makedirs(output_folder, exist_ok=True)
            imagens = pptx_para_jpeg(output_pptx_path, output_folder)

            # Criar um arquivo ZIP com as imagens
            zip_path = "apresentacao_modificada.zip"
            try:
                with zipfile.ZipFile(zip_path, "w") as zipf:
                    for imagem in imagens:
                        zipf.write(imagem, os.path.basename(imagem))
            except Exception as e:
                st.error(f"Erro ao criar o arquivo ZIP: {e}")
                return

            # Disponibilizar o download do ZIP
            try:
                with open(zip_path, "rb") as file:
                    btn = st.download_button(
                        label="Baixar Apresentação em JPEG (ZIP)",
                        data=file,
                        file_name="apresentacao_modificada.zip",
                        mime="application/zip"
                    )
            except Exception as e:
                st.error(f"Erro ao disponibilizar o download: {e}")
                return

            # Limpar as imagens baixadas e os slides JPEG
            for caminho in caminhos_imagens:
                if os.path.exists(caminho):
                    os.remove(caminho)
            for imagem in imagens:
                if os.path.exists(imagem):
                    os.remove(imagem)
            if os.path.exists(output_folder):
                os.rmdir(output_folder)
            if os.path.exists(output_pptx_path):
                os.remove(output_pptx_path)
            if os.path.exists(zip_path):
                os.remove(zip_path)

            st.success("Apresentação gerada com sucesso!")

if __name__ == "__main__":
    main()