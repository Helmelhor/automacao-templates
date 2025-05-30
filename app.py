import os
import pandas as pd
from pptx import Presentation
from pptx.util import Pt
import requests
import google.generativeai as genai
from dotenv import load_dotenv
import streamlit as st
from io import BytesIO
# PIL Image não é mais necessário se não vamos salvar slides como imagens
# from PIL import Image

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
        model = genai.GenerativeModel('gemini-1.5-flash-latest')
        response = model.generate_content(
            f"Faça um resumo de no máximo 445 caracteres sobre o livro: {livro}"
        )
        return response.text
    except Exception as e:
        st.error(f"Erro ao gerar resumo para '{livro}': {e}")
        return f"Erro ao gerar resumo para {livro}."

def gerar_frase_motivacional(livros):
    try:
        temas = ", ".join(livros)
        model = genai.GenerativeModel('gemini-1.5-flash-latest')
        response = model.generate_content(
            f"Crie apenas uma frase curta e motivacional de no máximo 100 caracteres que incentive a leitura (não quero sugestões, apenas retorne a frase sem mais nada além disso. não precisa deixar em negrito, então não coloque asteriscos '*'), baseada nos seguintes livros: {temas}"
        )
        return response.text
    except Exception as e:
        st.error(f"Erro ao gerar frase motivacional: {e}")
        return "A leitura transforma vidas."

# Função para baixar imagens a partir de URLs
def baixar_imagem(url, caminho_local):
    try:
        resposta = requests.get(url, timeout=10)
        resposta.raise_for_status()
        with open(caminho_local, "wb") as arquivo:
            arquivo.write(resposta.content)
    except requests.exceptions.RequestException as e:
        st.error(f"Erro ao baixar a imagem de {url}: {e}")
        raise
    except Exception as e:
        st.error(f"Erro inesperado ao processar a imagem de {url}: {e}")
        raise

# --- Funções Modificadas/Novas para Interação com Excel ---

def carregar_planilha(caminho_excel):
    try:
        df = pd.read_excel(caminho_excel, sheet_name='controle')
        colunas_necessarias = ['livros', 'autores', 'livros usados']
        if not all(col in df.columns for col in colunas_necessarias):
            st.error(f"O arquivo Excel '{caminho_excel}' não contém todas as colunas necessárias: {colunas_necessarias}.")
            return None
        # Normalizar a coluna 'livros usados' para strings e maiúsculas para consistência
        df['livros usados'] = df['livros usados'].astype(str).str.upper().fillna("NÃO")
        return df
    except FileNotFoundError:
        st.error(f"Arquivo Excel '{caminho_excel}' não encontrado.")
        return None
    except Exception as e:
        st.error(f"Erro ao ler o arquivo Excel '{caminho_excel}': {e}")
        return None

def selecionar_livros_do_acervo(caminho_excel, aleatorio=False):
    df = carregar_planilha(caminho_excel)
    if df is None:
        return None

    livros_nao_usados = df[df['livros usados'] != "SIM"]
    livros_selecionados_df = pd.DataFrame() # Inicializa como DataFrame vazio

    if not livros_nao_usados.empty:
        if aleatorio:
            num_a_selecionar = min(3, len(livros_nao_usados))
            if num_a_selecionar > 0:
                livros_selecionados_df = livros_nao_usados.sample(n=num_a_selecionar, random_state=None)
        else:
            livros_selecionados_df = livros_nao_usados.head(3)
    
    num_selecionados = len(livros_selecionados_df)
    if num_selecionados < 3:
        st.warning(f"Encontrados {num_selecionados} livro(s) 'não usado(s)'. Tentando completar com livros já usados...")
        livros_ja_usados = df[df['livros usados'] == "SIM"]
        # Evitar duplicatas se um livro já foi selecionado dos não usados
        livros_ja_usados_para_completar = livros_ja_usados[~livros_ja_usados['livros'].isin(livros_selecionados_df['livros'])]
        
        num_faltando = 3 - num_selecionados
        if not livros_ja_usados_para_completar.empty and num_faltando > 0:
            if aleatorio: # Se o modo é aleatório, tenta pegar aleatórios dos já usados também
                num_a_pegar_dos_usados = min(num_faltando, len(livros_ja_usados_para_completar))
                livros_para_adicionar = livros_ja_usados_para_completar.sample(n=num_a_pegar_dos_usados, random_state=None)
            else:
                livros_para_adicionar = livros_ja_usados_para_completar.head(num_faltando)
            
            livros_selecionados_df = pd.concat([livros_selecionados_df, livros_para_adicionar]).reset_index(drop=True)
        elif num_faltando > 0:
            st.warning("Não há livros 'já usados' suficientes para completar a seleção de 3 livros.")

    if livros_selecionados_df.empty:
        st.warning("Nenhum livro disponível para seleção no acervo.")
        return None
    
    return livros_selecionados_df.head(3) # Garante no máximo 3 e lida com DataFrame vazio

def marcar_livros_como_usados(caminho_excel, livros_df_para_marcar):
    if livros_df_para_marcar is None or livros_df_para_marcar.empty:
        st.info("Nenhum livro do acervo para marcar como usado.")
        return

    df = carregar_planilha(caminho_excel) # Recarrega para garantir dados mais recentes antes de escrever
    if df is None:
        return

    nomes_dos_livros_para_marcar = livros_df_para_marcar['livros'].tolist()
    indices_para_marcar = df['livros'].isin(nomes_dos_livros_para_marcar)
    df.loc[indices_para_marcar, 'livros usados'] = "SIM"

    try:
        df.to_excel(caminho_excel, sheet_name='controle', index=False)
        st.success(f"Planilha '{caminho_excel}' atualizada. Livros selecionados marcados como 'SIM'.")
    except Exception as e:
        st.error(f"Erro ao salvar as atualizações na planilha Excel '{caminho_excel}': {e}")

# --- Fim das Funções de Interação com Excel ---

def limpar_estado_pos_geracao_ou_nova_sugestao(limpar_tudo=False):
    """Limpa o estado do formulário e arquivos temporários."""
    if limpar_tudo: # Usado após download bem-sucedido
        st.session_state.form_data = {
            'metodo': "Selecionar do acervo", 'livro1': "", 'livro2': "", 'livro3': "",
            'img1': "", 'img2': "", 'img3': "", 'livros_df': None, 'resumos': [], 'frase_motivacional': ""
        }
        st.session_state.livros_acervo_selecionados_df = None
        st.session_state.metodo_anterior = "Selecionar do acervo"
        st.session_state.ppt_file_info = None # Limpa info do arquivo PPT
    else: # Usado ao pedir novas sugestões de livros
        st.session_state.form_data.update({
            'img1': "", 'img2': "", 'img3': "",
            'resumos': [], 'frase_motivacional': ""
        })

    # Limpar imagens temporárias da prévia
    for i in range(3):
        path_temp_preview = f"temp_preview_img_{i}.jpg"
        if os.path.exists(path_temp_preview):
            try:
                os.remove(path_temp_preview)
            except Exception as e:
                st.warning(f"Não foi possível remover o arquivo temporário {path_temp_preview}: {e}")


def main():
    st.set_page_config(layout="wide") # Opcional: usar layout mais largo
    st.title("Gerador de Apresentações de Livros")
    
    # Carregar imagem do cabeçalho localmente se existir
    caminho_header = "imagens/headerbooks.jpg"
    if os.path.exists(caminho_header):
        st.image(caminho_header, use_column_width=True)
    else:
        st.warning(f"Imagem de cabeçalho não encontrada em: {caminho_header}")

    # Inicializar session state
    if 'form_data' not in st.session_state:
        st.session_state.form_data = {
            'metodo': "Selecionar do acervo", 'livro1': "", 'livro2': "", 'livro3': "",
            'img1': "", 'img2': "", 'img3': "", 'livros_df': None, 
            'resumos': [], 'frase_motivacional': ""
        }
    if 'livros_acervo_selecionados_df' not in st.session_state:
        st.session_state.livros_acervo_selecionados_df = None
    if 'metodo_anterior' not in st.session_state:
        st.session_state.metodo_anterior = st.session_state.form_data['metodo']
    if 'ppt_file_info' not in st.session_state: # Para armazenar {'path': ..., 'name': ...}
        st.session_state.ppt_file_info = None


    caminho_excel = "livros.xlsx"
    if not os.path.exists(caminho_excel):
        st.warning(f"Arquivo Excel '{caminho_excel}' não encontrado. Criando um novo arquivo de exemplo...")
        try:
            df_exemplo = pd.DataFrame({
                "livros": ["O Pequeno Príncipe", "Dom Casmurro", "1984", "A Revolução dos Bichos", "Orgulho e Preconceito", "Cem Anos de Solidão", "O Senhor dos Anéis", "Harry Potter e a Pedra Filosofal", "O Hobbit", "Crônica de Nárnia"],
                "autores": ["Antoine de Saint-Exupéry", "Machado de Assis", "George Orwell", "George Orwell", "Jane Austen", "Gabriel García Márquez", "J.R.R. Tolkien", "J.K. Rowling", "J.R.R. Tolkien", "C.S. Lewis"],
                "livros usados": ["Não", "Não", "Não", "Não", "Não", "Não", "Não", "Não", "Não", "Não"]
            })
            df_exemplo.to_excel(caminho_excel, sheet_name='controle', index=False)
            st.success(f"Arquivo Excel de exemplo criado: {caminho_excel}")
        except Exception as e:
            st.error(f"Não foi possível criar o arquivo Excel de exemplo: {e}")
            st.stop()

    st.sidebar.title("Navegação")
    aba_selecionada = st.sidebar.radio("Selecione a aba", ["Gerar Apresentação", "Visualizar Planilha"])

    if aba_selecionada == "Visualizar Planilha":
        st.header("Planilha de Controle de Livros")
        df_display = carregar_planilha(caminho_excel)
        if df_display is not None:
            st.dataframe(df_display, height=600) # Aumentar altura do dataframe
        return

    # --- Aba de Gerar Apresentação ---
    st.header("Configurar Conteúdo da Apresentação")

    # Se um PPT foi gerado, mostrar o botão de download primeiro
    if st.session_state.ppt_file_info:
        ppt_path = st.session_state.ppt_file_info['path']
        ppt_name = st.session_state.ppt_file_info['name']
        if os.path.exists(ppt_path):
            with open(ppt_path, "rb") as f_ppt:
                st.download_button(
                    label="Clique aqui para Baixar a Apresentação PPTX Gerada",
                    data=f_ppt,
                    file_name=ppt_name,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    on_click=limpar_estado_pos_geracao_ou_nova_sugestao,
                    kwargs={'limpar_tudo': True},
                    key="download_ppt_final"
                )
            st.success(f"Sua apresentação '{ppt_name}' está pronta para download!")
            st.info("Após o download, a página será recarregada para uma nova geração.")
            # O on_click já lida com a limpeza. O rerun implícito do Streamlit acontece.
        else:
            st.error("O arquivo PPTX gerado não foi encontrado. Por favor, tente gerar novamente.")
            st.session_state.ppt_file_info = None # Limpa para evitar loop
        return # Não mostrar o resto do formulário se o download estiver pendente

    metodo_atual = st.radio(
        "Como deseja selecionar os livros para a apresentação?",
        ["Selecionar do acervo", "Digitar manualmente"],
        index=0 if st.session_state.form_data['metodo'] == "Selecionar do acervo" else 1,
        key="metodo_selecao_radio"
    )

    if metodo_atual != st.session_state.form_data['metodo']:
        st.session_state.form_data['metodo'] = metodo_atual
        st.session_state.livros_acervo_selecionados_df = None
        st.session_state.form_data['livros_df'] = None
        if metodo_atual == "Selecionar do acervo":
            st.session_state.form_data.update({'livro1': "", 'livro2': "", 'livro3': ""})
        limpar_estado_pos_geracao_ou_nova_sugestao(limpar_tudo=False) # Limpa campos de imagem/IA

    if st.session_state.form_data['metodo'] == "Selecionar do acervo":
        if st.session_state.livros_acervo_selecionados_df is None:
            st.session_state.livros_acervo_selecionados_df = selecionar_livros_do_acervo(caminho_excel, aleatorio=False) # Primeira seleção não aleatória
        
        st.session_state.form_data['livros_df'] = st.session_state.livros_acervo_selecionados_df

        if st.session_state.form_data['livros_df'] is not None and not st.session_state.form_data['livros_df'].empty:
            st.subheader("Livros Selecionados do Acervo:")
            st.dataframe(st.session_state.form_data['livros_df'][['livros', 'autores']])
            
            if st.button("🔄 Sugerir outros livros do acervo", key="sugerir_outros_acervo_btn"):
                st.session_state.livros_acervo_selecionados_df = selecionar_livros_do_acervo(caminho_excel, aleatorio=True)
                st.session_state.form_data['livros_df'] = st.session_state.livros_acervo_selecionados_df
                limpar_estado_pos_geracao_ou_nova_sugestao(limpar_tudo=False) # Limpa imagens e IA, mas mantém método e livros
                st.rerun()
        else:
            st.warning("Não foi possível selecionar livros do acervo no momento ou não há livros disponíveis.")
    
    else: # Digitar manualmente
        st.subheader("Digite os Livros Manualmente")
        if st.session_state.metodo_anterior == "Selecionar do acervo":
             st.session_state.form_data['livros_df'] = None

        col1, col2, col3 = st.columns(3)
        with col1: st.session_state.form_data['livro1'] = st.text_input("Livro 1", value=st.session_state.form_data['livro1'], key="man_livro1")
        with col2: st.session_state.form_data['livro2'] = st.text_input("Livro 2", value=st.session_state.form_data['livro2'], key="man_livro2")
        with col3: st.session_state.form_data['livro3'] = st.text_input("Livro 3", value=st.session_state.form_data['livro3'], key="man_livro3")

        if st.session_state.form_data['livro1'] and st.session_state.form_data['livro2'] and st.session_state.form_data['livro3']:
            st.session_state.form_data['livros_df'] = pd.DataFrame({
                'livros': [st.session_state.form_data['livro1'], st.session_state.form_data['livro2'], st.session_state.form_data['livro3']],
                'autores': ['Autor Desconhecido'] * 3
            })
        else:
             st.session_state.form_data['livros_df'] = None

    st.session_state.metodo_anterior = st.session_state.form_data['metodo']

    st.subheader("Links das Imagens dos Livros (obrigatório)")
    col_img_input1, col_img_input2, col_img_input3 = st.columns(3)
    with col_img_input1: st.session_state.form_data['img1'] = st.text_input("URL Imagem Livro 1", value=st.session_state.form_data['img1'], key="url_img1")
    with col_img_input2: st.session_state.form_data['img2'] = st.text_input("URL Imagem Livro 2", value=st.session_state.form_data['img2'], key="url_img2")
    with col_img_input3: st.session_state.form_data['img3'] = st.text_input("URL Imagem Livro 3", value=st.session_state.form_data['img3'], key="url_img3")

    inputs_validos = False
    if st.session_state.form_data['livros_df'] is not None and not st.session_state.form_data['livros_df'].empty:
        if len(st.session_state.form_data['livros_df']) == 3:
            if st.session_state.form_data['img1'] and st.session_state.form_data['img2'] and st.session_state.form_data['img3']:
                inputs_validos = True
            else:
                st.warning("Por favor, forneça os links das imagens para os 3 livros.")
        # else: st.warning("São necessários 3 livros para gerar a apresentação. Seleção atual pode estar incompleta.") # Já coberto
    else:
        if st.session_state.form_data['metodo'] == "Selecionar do acervo" and (st.session_state.form_data['livros_df'] is None or st.session_state.form_data['livros_df'].empty):
            pass # Avisos já dados pela função de seleção
        else:
            st.info("Selecione ou digite 3 livros e forneça os links das imagens para continuar.")

    if inputs_validos:
        st.markdown("---")
        st.subheader("Prévia do Conteúdo Gerado")
        
        livros_para_ia = st.session_state.form_data['livros_df']['livros'].tolist()
        
        # Gerar resumos e frase se ainda não existem ou se os livros mudaram (simplificado)
        if not st.session_state.form_data['resumos'] or len(st.session_state.form_data['resumos']) != 3 :
            with st.spinner("Gerando resumos e frase motivacional com IA..."):
                st.session_state.form_data['resumos'] = [gerar_resumo(livro) for livro in livros_para_ia]
                st.session_state.form_data['frase_motivacional'] = gerar_frase_motivacional(livros_para_ia)

        for i, resumo in enumerate(st.session_state.form_data['resumos']):
            st.write(f"**Resumo Livro {i+1} ({livros_para_ia[i]}):** {resumo}")
        st.write(f"**Frase Motivacional:** {st.session_state.form_data['frase_motivacional']}")
        
        st.markdown("---")
        st.subheader("Prévia das Imagens (após download)")
        
        img_paths_preview = []
        img_urls = [st.session_state.form_data['img1'], st.session_state.form_data['img2'], st.session_state.form_data['img3']]
        
        cols_img_preview = st.columns(3)
        imagens_ok_para_ppt = True
        for i, url in enumerate(img_urls):
            if url: # Só tenta baixar se a URL não estiver vazia
                try:
                    path_temp = f"temp_preview_img_{i}.jpg"
                    baixar_imagem(url, path_temp)
                    cols_img_preview[i].image(path_temp, caption=f"Prévia Livro {i+1}", use_column_width=True)
                    img_paths_preview.append(path_temp)
                except Exception as e_img_preview:
                    cols_img_preview[i].error(f"Erro imagem {i+1}") # Mensagem já dada por baixar_imagem
                    imagens_ok_para_ppt = False
            else: # URL vazia
                cols_img_preview[i].warning(f"URL da imagem {i+1} não fornecida.")
                imagens_ok_para_ppt = False
        
        if not imagens_ok_para_ppt:
            st.error("Corrija os problemas com as imagens (URLs vazias ou erros de download) antes de gerar a apresentação.")

        if imagens_ok_para_ppt and st.button("Gerar Apresentação PPTX", key="gerar_ppt_final_btn"):
            with st.spinner("Montando a apresentação... Por favor, aguarde."):
                try:
                    # Tentar carregar o template. Se não existir, criar um básico.
                    template_path = 'minha_apresentacao.pptx'
                    if not os.path.exists(template_path):
                        st.warning(f"Template '{template_path}' não encontrado. Gerando um template básico.")
                        prs = Presentation()
                        # Adicionar alguns layouts básicos se estiver criando do zero
                        slide_layout = prs.slide_layouts[5] # Layout em branco
                        for _ in range(3): # Adicionar 3 slides de exemplo
                            slide = prs.slides.add_slide(slide_layout)
                            slide.shapes.add_textbox(Pt(50), Pt(50), Pt(300), Pt(50)).text_frame.text = "Placeholder Título"
                            # Adicione mais placeholders se necessário para 'texto1', 'imagem1', etc.
                    else:
                         prs = Presentation(template_path)
                    
                    def replace_text_in_slide(slide, old_text_placeholder, new_text, font_size=Pt(14)):
                        for shape in slide.shapes:
                            if shape.has_text_frame and old_text_placeholder in shape.text_frame.text:
                                shape.text_frame.text = shape.text_frame.text.replace(old_text_placeholder, new_text)
                                for para in shape.text_frame.paragraphs:
                                    for run in para.runs:
                                        run.font.size = font_size
                    
                    def replace_image_in_slide(slide, image_placeholder_name, new_image_path):
                        pic_shape_to_replace = None
                        for shape in slide.shapes:
                            if shape.name == image_placeholder_name:
                                pic_shape_to_replace = shape
                                break
                        if pic_shape_to_replace:
                            slide.shapes.add_picture(new_image_path, 
                                                     pic_shape_to_replace.left, pic_shape_to_replace.top, 
                                                     pic_shape_to_replace.width, pic_shape_to_replace.height)
                            sp = pic_shape_to_replace._element
                            sp.getparent().remove(sp)
                        else:
                            st.warning(f"Placeholder de imagem '{image_placeholder_name}' não encontrado no slide.")

                    # Placeholders esperados no seu template
                    placeholders_text_map = {
                        "texto1": st.session_state.form_data['resumos'][0] if len(st.session_state.form_data['resumos']) > 0 else "",
                        "texto2": st.session_state.form_data['resumos'][1] if len(st.session_state.form_data['resumos']) > 1 else "",
                        "texto3": st.session_state.form_data['resumos'][2] if len(st.session_state.form_data['resumos']) > 2 else "",
                        "texto4": st.session_state.form_data['frase_motivacional']
                    }
                    placeholders_img_map = {
                        "imagem1": img_paths_preview[0] if len(img_paths_preview) > 0 else None,
                        "imagem2": img_paths_preview[1] if len(img_paths_preview) > 1 else None,
                        "imagem3": img_paths_preview[2] if len(img_paths_preview) > 2 else None,
                    }

                    for slide in prs.slides:
                        for placeholder, text_content in placeholders_text_map.items():
                            replace_text_in_slide(slide, placeholder, text_content)
                        for placeholder_name, img_path in placeholders_img_map.items():
                            if img_path and os.path.exists(img_path): # Verifica se o caminho da imagem é válido
                                replace_image_in_slide(slide, placeholder_name, img_path)
                    
                    output_filename = 'apresentacao_gerada.pptx'
                    prs.save(output_filename)
                    
                    if st.session_state.form_data['metodo'] == "Selecionar do acervo":
                        marcar_livros_como_usados(caminho_excel, st.session_state.form_data['livros_df'])
                    
                    st.session_state.ppt_file_info = {'path': output_filename, 'name': 'apresentacao_livros_final.pptx'}
                    st.rerun() # Para mostrar o botão de download no topo

                except Exception as e_ppt:
                    st.error(f"Erro ao gerar a apresentação PPTX: {e_ppt}")
                    # Limpar imagens temporárias da prévia mesmo em caso de erro
                    for path in img_paths_preview:
                        if os.path.exists(path):
                            try: os.remove(path)
                            except: pass
    # elif st.session_state.form_data['livros_df'] is None or st.session_state.form_data['livros_df'].empty or len(st.session_state.form_data['livros_df']) < 3 :
    #     st.info("Aguardando seleção/digitação de 3 livros válidos.")
    # elif not (st.session_state.form_data['img1'] and st.session_state.form_data['img2'] and st.session_state.form_data['img3']):
    #     st.info("Aguardando os links das 3 imagens.")

if __name__ == "__main__":
    main()