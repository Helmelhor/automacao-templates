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
        resposta = requests.get(url, timeout=10)
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
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        prs = Presentation(pptx_path)

        for i, slide in enumerate(prs.slides):
            slide_image_path = os.path.join(output_folder, f"slide_{i + 1}.png")
            slide_image = slide.export()
            Image.open(slide_image).save(slide_image_path)

        st.success(f"Slides salvos como imagens na pasta: {output_folder}")
    except Exception as e:
        st.error(f"Erro ao salvar slides como imagens: {e}")

# Função para ler e atualizar o Excel
def ler_e_atualizar_excel(caminho_excel):
    try:
        df = pd.read_excel(caminho_excel, sheet_name='controle')

        if 'livros' not in df.columns or 'autores' not in df.columns or 'livros usados' not in df.columns:
            st.error("O arquivo Excel não contém as colunas necessárias.")
            return None, None

        livros_nao_usados = df[df['livros usados'] != "Sim"]
        livros_selecionados = livros_nao_usados.head(3)

        if len(livros_selecionados) < 3:
            st.warning("Não há livros suficientes não utilizados. Selecionando livros já usados...")
            livros_usados = df[df['livros usados'] == "Sim"]
            livros_selecionados = pd.concat([livros_selecionados, livros_usados.head(3 - len(livros_selecionados))])

        df.loc[df['livros'].isin(livros_selecionados['livros']), 'livros usados'] = "Sim"
        df.to_excel(caminho_excel, sheet_name='controle', index=False)

        return livros_selecionados, df
    except Exception as e:
        st.error(f"Erro ao ler/atualizar o Excel: {e}")
        return None, None

def main():
    st.title("Template de Livros")
    st.image("imagens\headerbooks.jpg")

    # Inicializar session state
    if 'form_data' not in st.session_state:
        st.session_state.form_data = {
            'metodo': "Selecionar do acervo",
            'livro1': "",
            'livro2': "",
            'livro3': "",
            'img1': "",
            'img2': "",
            'img3': "",
            'livros_df': None,
            'resumos': [],
            'frase_motivacional': ""
        }

    caminho_excel = "livros.xlsx"

    if not os.path.exists(caminho_excel):
        st.warning("Arquivo Excel não encontrado. Criando um novo arquivo...")
        df = pd.DataFrame(columns=["livros", "autores", "livros usados"])
        df.to_excel(caminho_excel, sheet_name='controle', index=False)
        st.success(f"Arquivo Excel criado: {caminho_excel}")

    # Sidebar navigation
    st.sidebar.title("Navegação")
    aba = st.sidebar.radio("Selecione a aba", ["Apresentação", "Planilha"])

    if aba == "Planilha":
        st.header("Planilha de Livros")
        try:
            df = pd.read_excel(caminho_excel, sheet_name='controle')
            st.dataframe(df)
        except Exception as e:
            st.error(f"Erro ao carregar a planilha: {e}")
        return

    # Aba de Apresentação
    st.header("Gerar Apresentação")

    # Método de seleção
    metodo = st.radio(
        "Como deseja selecionar os livros?",
        ["Selecionar do acervo", "Digitar manualmente"],
        index=0 if st.session_state.form_data['metodo'] == "Selecionar do acervo" else 1,
        key="metodo_selecao"
    )
    st.session_state.form_data['metodo'] = metodo

    if metodo == "Selecionar do acervo":
        livros_df, _ = ler_e_atualizar_excel(caminho_excel)
        st.session_state.form_data['livros_df'] = livros_df

        if livros_df is not None:
            st.subheader("Livros Selecionados")
            st.dataframe(livros_df[['livros', 'autores']])
    else:
        st.subheader("Digite os Livros Manualmente")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            livro1 = st.text_input("Livro 1", 
                                 value=st.session_state.form_data['livro1'],
                                 key="input_livro1",
                                 on_change=lambda: st.session_state.form_data.update({
                                     'livro1': st.session_state.input_livro1
                                 }))
        
        with col2:
            livro2 = st.text_input("Livro 2", 
                                 value=st.session_state.form_data['livro2'],
                                 key="input_livro2",
                                 on_change=lambda: st.session_state.form_data.update({
                                     'livro2': st.session_state.input_livro2
                                 }))
        
        with col3:
            livro3 = st.text_input("Livro 3", 
                                 value=st.session_state.form_data['livro3'],
                                 key="input_livro3",
                                 on_change=lambda: st.session_state.form_data.update({
                                     'livro3': st.session_state.input_livro3
                                 }))

        if livro1 and livro2 and livro3:
            st.session_state.form_data['livros_df'] = pd.DataFrame({
                'livros': [livro1, livro2, livro3],
                'autores': ['Autor Desconhecido'] * 3
            })

    # Links das imagens
    st.subheader("Links das Imagens dos Livros")
    
    img1 = st.text_input("Imagem Livro 1", 
                        value=st.session_state.form_data['img1'],
                        key="input_img1",
                        on_change=lambda: st.session_state.form_data.update({
                            'img1': st.session_state.input_img1
                        }))
    
    img2 = st.text_input("Imagem Livro 2", 
                        value=st.session_state.form_data['img2'],
                        key="input_img2",
                        on_change=lambda: st.session_state.form_data.update({
                            'img2': st.session_state.input_img2
                        }))
    
    img3 = st.text_input("Imagem Livro 3", 
                        value=st.session_state.form_data['img3'],
                        key="input_img3",
                        on_change=lambda: st.session_state.form_data.update({
                            'img3': st.session_state.input_img3
                        }))

    # Verificar se todos os campos estão preenchidos
    if (st.session_state.form_data['livros_df'] is not None and 
        img1 and img2 and img3):
        
        # Baixar imagens
        img_paths = []
        for i, url in enumerate([img1, img2, img3]):
            try:
                path = f"temp_img_{i}.jpg"
                baixar_imagem(url, path)
                img_paths.append(path)
            except Exception as e:
                st.error(f"Erro ao baixar imagem {i+1}: {e}")
                return

        # Gerar conteúdo
        livros = st.session_state.form_data['livros_df']['livros'].tolist()
        resumos = [gerar_resumo(livro) for livro in livros]
        frase = gerar_frase_motivacional(livros)
        
        st.session_state.form_data['resumos'] = resumos
        st.session_state.form_data['frase_motivacional'] = frase

        # Mostrar prévia
        st.subheader("Prévia da Apresentação")
        
        st.write("**Resumos:**")
        for i, resumo in enumerate(resumos):
            st.write(f"**Livro {i+1}:** {resumo}")
        
        st.write("**Frase Motivacional:**")
        st.write(frase)
        
        cols = st.columns(3)
        for i, path in enumerate(img_paths):
            try:
                cols[i].image(path, caption=f"Livro {i+1}", use_column_width=True)
            except Exception as e:
                st.error(f"Erro ao mostrar imagem {i+1}: {e}")

        # Botão para gerar apresentação
        if st.button("Gerar Apresentação"):
            try:
                prs = Presentation('minha_apresentacao.pptx')
                
                # Funções auxiliares para substituição
                def replace_text(slide, old, new, size=Pt(14)):
                    for shape in slide.shapes:
                        if shape.has_text_frame:
                            for para in shape.text_frame.paragraphs:
                                if old in para.text:
                                    para.text = para.text.replace(old, new)
                                    for run in para.runs:
                                        run.font.size = size
                
                def replace_img(slide, name, new_img):
                    for shape in slide.shapes:
                        if shape.name == name:
                            left, top, width, height = shape.left, shape.top, shape.width, shape.height
                            slide.shapes.add_picture(new_img, left, top, width, height)
                            sp = shape._element
                            sp.getparent().remove(sp)
                
                # Aplicar substituições
                for slide in prs.slides:
                    replace_text(slide, 'texto1', resumos[0])
                    replace_text(slide, 'texto2', resumos[1])
                    replace_text(slide, 'texto3', resumos[2])
                    replace_text(slide, 'texto4', frase)
                    replace_img(slide, 'imagem1', img_paths[0])
                    replace_img(slide, 'imagem2', img_paths[1])
                    replace_img(slide, 'imagem3', img_paths[2])
                
                # Salvar e oferecer download
                output_path = 'apresentacao_final.pptx'
                prs.save(output_path)
                
                with open(output_path, "rb") as f:
                    st.download_button(
                        "Baixar Apresentação",
                        f,
                        file_name="apresentacao_livros.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
                
                # Limpar imagens temporárias
                for path in img_paths:
                    if os.path.exists(path):
                        os.remove(path)
                
                st.success("Apresentação gerada com sucesso!")
                
            except Exception as e:
                st.error(f"Erro ao gerar apresentação: {e}")

if __name__ == "__main__":
    main()