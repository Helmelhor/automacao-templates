# Gerador Automático de Apresentações de Livros

Este projeto é uma aplicação desenvolvida para a **ACLF Empreendimentos** que automatiza a criação de apresentações de livros. A ferramenta permite que o usuário insira os nomes de três livros e os links das imagens correspondentes, e gera automaticamente uma apresentação no formato PowerPoint (`.pptx`) com resumos dos livros, imagens e uma frase motivacional.

A aplicação utiliza a API do **Google Gemini** para gerar os resumos e a frase motivacional, e é hospedada no **Streamlit Community Cloud** para fácil acesso e uso.

---

## Como Utilizar

### Acessando a Aplicação

1. **Acesse o Link**:
   - A aplicação está disponível no Streamlit Community Cloud. Clique no link abaixo para acessar:
     [Link da Aplicação](https://automacao-templates-d84tnvyn58ewodczp7dapp5.streamlit.app/)

2. **Preencha os Campos**:
   - Na página inicial, você verá campos para inserir:
     - **Nomes dos Livros**: Insira o nome de três livros.
     - **Links das Imagens**: Insira os links das imagens correspondentes aos livros (URLs de imagens da web).

3. **Gere a Apresentação**:
   - Após preencher todos os campos, clique no botão **"Gerar Apresentação"**.
   - A aplicação irá:
     - Baixar as imagens a partir dos links fornecidos.
     - Gerar resumos dos livros e uma frase motivacional usando a API do Gemini.
     - Criar uma apresentação no formato PowerPoint com os dados fornecidos.
     - Disponibilizar o download da apresentação em formato ZIP contendo os slides em JPEG.

4. **Faça o Download**:
   - Após a geração da apresentação, um botão de download será exibido. Clique nele para baixar o arquivo ZIP.

---

### Requisitos

Para utilizar a aplicação, você precisará de:

- **Nomes dos Livros**: Três títulos de livros.
- **Links das Imagens**: URLs de imagens correspondentes aos livros (formato JPEG ou PNG).
- **Acesso à Internet**: A aplicação utiliza a API do Gemini e precisa de conexão com a internet.

---

## Desenvolvimento

### Tecnologias Utilizadas

- **Streamlit**: Para a interface web.
- **Google Gemini API**: Para gerar resumos e frases motivacionais.
- **Python-pptx**: Para manipulação de arquivos PowerPoint.
- **Pillow**: Para manipulação de imagens.
- **Requests**: Para baixar imagens a partir de URLs.

### Estrutura do Projeto

```
meu_projeto/
├── app.py                  # Código principal da aplicação
├── requirements.txt        # Lista de dependências
├── .env                    # Arquivo de configuração com a chave da API
├── minha_apresentacao.pptx # Template da apresentação
├── README.md               # Documentação do projeto
└── imagens/                # Pasta para imagens (opcional)
```

---

## Como Executar Localmente

Se você deseja executar o projeto localmente, siga os passos abaixo:

1. **Clone o Repositório**:
   ```bash
   git clone https://github.com/seu-usuario/automacao-templates.git
   cd automacao-templates
   ```

2. **Instale as Dependências**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Configure a Chave da API**:
   - Crie um arquivo `.env` na raiz do projeto e adicione a chave da API do Gemini:
     ```plaintext
     API_KEY=sua_chave_aqui
     ```

4. **Execute a Aplicação**:
   ```bash
   streamlit run app.py
   ```

5. **Acesse a Aplicação**:
   - Abra o navegador e acesse o endereço fornecido pelo Streamlit (geralmente `http://localhost:8501`).

---

## Sobre o Projeto

Este projeto foi desenvolvido especificamente para a **ACLF Empreendimentos** com o objetivo de automatizar a criação de apresentações de livros, facilitando a geração de conteúdo visual e textual de forma rápida e eficiente.

### Desenvolvedor
- **Nome**: [Seu Nome]
- **Contato**: [Seu Email ou LinkedIn]

---

## Licença

Este projeto está licenciado sob a licença MIT. Consulte o arquivo [LICENSE](LICENSE) para mais detalhes.
