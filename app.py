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
        return response.text
    except Exception as e:
        st.error(f"Erro ao gerar frase motivacional: {e}")
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

# Função para salvar slides como imagens
def salvar_slides_como_imagens(pptx_path, output_folder="slides_imagens"):
    try:
        # Criar a pasta de saída se não existir
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        # Carregar a apresentação
        prs = Presentation(pptx_path)

        # Iterar sobre os slides e salvar como imagens
        for i, slide in enumerate(prs.slides):
            slide_image_path = os.path.join(output_folder, f"slide_{i + 1}.png")
            slide_image = slide.export()  # Exporta o slide como imagem
            Image.open(slide_image).save(slide_image_path)  # Salva a imagem usando Pillow

        st.success(f"Slides salvos como imagens na pasta: {output_folder}")
    except Exception as e:
        st.error(f"Erro ao salvar slides como imagens: {e}")

# Função para ler e atualizar o Excel
def ler_e_atualizar_excel(caminho_excel):
    try:
        # Ler o arquivo Excel
        df = pd.read_excel(caminho_excel, sheet_name='controle')

        # Verificar se as colunas necessárias existem
        if 'livros' not in df.columns or 'autores' not in df.columns or 'livros usados' not in df.columns:
            st.error("O arquivo Excel não contém as colunas necessárias: 'livros', 'autores', 'livros usados'.")
            return None, None

        # Filtrar livros não usados
        livros_nao_usados = df[df['livros usados'] != "Sim"]

        # Selecionar os primeiros 3 livros não usados
        livros_selecionados = livros_nao_usados.head(3)

        # Se não houver livros não usados suficientes, selecionar livros usados
        if len(livros_selecionados) < 3:
            st.warning("Não há livros suficientes não utilizados. Selecionando livros já usados...")
            livros_usados = df[df['livros usados'] == "Sim"]
            livros_selecionados = pd.concat([livros_selecionados, livros_usados.head(3 - len(livros_selecionados))])

        # Atualizar a coluna "livros usados" no DataFrame
        df.loc[df['livros'].isin(livros_selecionados['livros']), 'livros usados'] = "Sim"

        # Salvar o DataFrame atualizado de volta no Excel
        df.to_excel(caminho_excel, sheet_name='controle', index=False)

        return livros_selecionados, df
    except Exception as e:
        st.error(f"Erro ao ler/atualizar o Excel: {e}")
        return None, None

def main():
    st.title("Template de Livros")

    # Caminho do arquivo Excel
    caminho_excel = "livros.xlsx"  # Substitua pelo caminho do seu arquivo Excel

    # Verificar se o arquivo Excel existe
    if not os.path.exists(caminho_excel):
        st.warning("Arquivo Excel não encontrado. Criando um novo arquivo...")
        df = pd.DataFrame(columns=["livros", "autores", "livros usados"])
        df.to_excel(caminho_excel, sheet_name='controle', index=False)
        st.success(f"Arquivo Excel criado: {caminho_excel}")

    # Sidebar para navegação
    st.sidebar.title("Navegação")
    aba_selecionada = st.sidebar.radio("Selecione a aba", ["Apresentação", "Planilha"])

    if aba_selecionada == "Planilha":
        st.header("Planilha de Livros")
        try:
            # Ler e exibir a planilha
            df = pd.read_excel(caminho_excel, sheet_name='controle')
            st.write(df)
        except Exception as e:
            st.error(f"Erro ao carregar a planilha: {e}")
        return

    # Aba de Apresentação
    st.header("Gerar Apresentação")

    # Opção para escolher entre digitar manualmente ou selecionar do acervo
    metodo_selecao = st.radio(
        "Como deseja selecionar os livros?",
        ["Selecionar do acervo", "Digitar manualmente"]
    )

    livros_selecionados = None

    if metodo_selecao == "Selecionar do acervo":
        # Ler e atualizar o Excel
        livros_selecionados, df = ler_e_atualizar_excel(caminho_excel)

        if livros_selecionados is None:
            return

        # Exibir os livros selecionados
        st.subheader("Livros Selecionados para a Apresentação")
        st.write(livros_selecionados[['livros', 'autores']])

    elif metodo_selecao == "Digitar manualmente":
        st.subheader("Digite os nomes dos livros")
        livro1 = st.text_input("Nome do Livro 1")
        livro2 = st.text_input("Nome do Livro 2")
        livro3 = st.text_input("Nome do Livro 3")

        if livro1 and livro2 and livro3:
            livros_selecionados = pd.DataFrame({
                'livros': [livro1, livro2, livro3],
                'autores': ['Autor Desconhecido'] * 3  # Placeholder para autores
            })
        else:
            st.warning("Por favor, insira os nomes dos três livros.")
            return

    # Inputs para os links das imagens
    st.subheader("Insira os links das imagens dos livros")
    link_imagem1 = st.text_input("Link da Imagem do Livro 1")
    link_imagem2 = st.text_input("Link da Imagem do Livro 2")
    link_imagem3 = st.text_input("Link da Imagem do Livro 3")

    # Verifica se todos os campos foram preenchidos
    if livros_selecionados is not None and link_imagem1 and link_imagem2 and link_imagem3:
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
        resumos = [gerar_resumo(livro) for livro in livros_selecionados['livros']]
        frase_motivacional = gerar_frase_motivacional(livros_selecionados['livros'].tolist())

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

            # Converter slides em imagens
            salvar_slides_como_imagens(output_pptx_path)

            # Disponibilizar o download do PPT
            with open(output_pptx_path, "rb") as file:
                btn = st.download_button(
                    label="Baixar Apresentação (PPT)",
                    data=file,
                    file_name="apresentacao_modificada.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )

            # Limpar as imagens baixadas imediatamente após o uso
            for caminho in caminhos_imagens:
                if os.path.exists(caminho):
                    try:
                        os.remove(caminho)
                        st.success(f"Imagem {caminho} removida com sucesso.")
                    except Exception as e:
                        st.error(f"Erro ao remover a imagem {caminho}: {e}")

            st.success("Apresentação gerada com sucesso!")

if __name__ == "__main__":
    main()