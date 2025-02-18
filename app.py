from pptx import Presentation
from pptx.util import Inches, Pt

# Caminho para o arquivo PowerPoint existente
pptx_path = 'minha_apresentacao.pptx'

# Abra a apresentação existente
prs = Presentation(pptx_path)

# Função para substituir texto e definir o tamanho da fonte em todos os placeholders
def substituir_texto(slide, antigo_texto, novo_texto, tamanho_fonte=Pt(14)):
    for shape in slide.shapes:
        if shape.has_text_frame:
            for para in shape.text_frame.paragraphs:
                if antigo_texto in para.text:
                    para.text = para.text.replace(antigo_texto, novo_texto)
                    for run in para.runs:
                        run.font.size = tamanho_fonte

# Função para substituir imagens baseadas nos nomes dos placeholders
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

# Loop pelos slides para modificar texto e imagens
for slide in prs.slides:
    substituir_texto(slide, 'texto1', 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.')
    substituir_texto(slide, 'texto2', 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.')
    substituir_texto(slide, 'texto3', 'Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do eiusmod tempor incididunt ut labore et dolore magna aliqua. Ut enim ad minim veniam, quis nostrud exercitation ullamco laboris nisi ut aliquip ex ea commodo consequat. Duis aute irure dolor in reprehenderit in voluptate velit esse cillum dolore eu fugiat nulla pariatur. Excepteur sint occaecat cupidatat non proident, sunt in culpa qui officia deserunt mollit anim id est laborum.')
    substituir_imagem_por_nome(slide, 'imagem1', r'imagens\conexão com DB.png')
    substituir_imagem_por_nome(slide, 'imagem2', r'imagens\conexão DB com looker studio.png')
    substituir_imagem_por_nome(slide, 'imagem3', r'imagens\contabilidade geral.png')

# Salve a apresentação modificada
prs.save('apresentacao_modificada.pptx')