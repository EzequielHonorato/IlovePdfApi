# 📄 iLovePDF Automação - PDF para Word

Este projeto é uma **API REST** que automatiza a conversão de arquivos PDF para Word usando o site [iLovePDF](https://www.ilovepdf.com/pt/pdf_para_word), sem custos de API.

## 🚀 Como Funciona

A API usa **Selenium** em background para:
1. Receber o PDF via endpoint REST
2. Abrir o navegador Chrome (modo headless)
3. Acessar a página de conversão PDF → Word do iLovePDF
4. Fazer upload do arquivo PDF
5. Iniciar a conversão
6. Baixar o arquivo Word convertido
7. Disponibilizar para download via API

## 📋 Pré-requisitos

- Python 3.8+
- Google Chrome instalado
- Conexão com a internet

## 🔧 Instalação

1. Clone ou baixe este projeto

2. Crie e ative o ambiente virtual (opcional, mas recomendado):
```bash
python -m venv .venv
source .venv/bin/activate  # Linux/macOS
# ou
.venv\Scripts\activate  # Windows
```

3. Instale as dependências:
```bash
pip install -r requirements.txt
```

## 📖 Uso

### Iniciar a API
```bash
python api.py
```

A API estará disponível em `http://localhost:8000`

### Documentação Interativa
Acesse `http://localhost:8000/docs` para ver a documentação Swagger.

### Endpoints

#### 1. Enviar PDF para conversão
```bash
curl -X POST "http://localhost:8000/convert" \
  -H "Content-Type: multipart/form-data" \
  -F "file=@documento.pdf"
```

**Resposta:**
```json
{
  "id": "abc123-uuid",
  "status": "pending",
  "message": "Conversão iniciada. Use /status/{id} para acompanhar."
}
```

#### 2. Verificar status da conversão
```bash
curl "http://localhost:8000/status/{id}"
```

**Respostas possíveis:**
```json
{
  "id": "abc123-uuid",
  "status": "processing",
  "message": "Convertendo PDF para Word..."
}
```

```json
{
  "id": "abc123-uuid",
  "status": "completed",
  "message": "Conversão concluída!",
  "url": "/download/abc123-uuid",
  "filename": "documento.docx"
}
```

#### 3. Baixar arquivo convertido
```bash
curl -O "http://localhost:8000/download/{id}"
```

### Exemplo de uso com JavaScript/Fetch
```javascript
// 1. Enviar PDF
const formData = new FormData();
formData.append('file', pdfFile);

const response = await fetch('http://localhost:8000/convert', {
  method: 'POST',
  body: formData
});
const { id } = await response.json();

// 2. Verificar status (polling)
const checkStatus = async () => {
  const res = await fetch(`http://localhost:8000/status/${id}`);
  const data = await res.json();
  
  if (data.status === 'completed') {
    // 3. Baixar arquivo
    window.location.href = `http://localhost:8000${data.url}`;
  } else if (data.status === 'processing') {
    setTimeout(checkStatus, 2000); // Verifica novamente em 2s
  }
};

checkStatus();
```

### Uso via linha de comando (script original)
```bash
python main.py caminho/do/arquivo.pdf
```

## ⚙️ Configurações

### Modo Headless (sem interface gráfica)

Para executar sem abrir a janela do navegador, descomente a linha no arquivo `main.py`:

```python
chrome_options.add_argument("--headless")
```

### Diretório de Download Padrão

Por padrão, os arquivos são salvos na pasta `Downloads` do usuário. Você pode alterar isso passando o segundo argumento ou modificando o código.

## 🔍 Estrutura do Projeto

```
scriptAPI/
├── .github/
│   └── copilot-instructions.md
├── .venv/                    # Ambiente virtual (não versionado)
├── uploads/                  # PDFs temporários (criado automaticamente)
├── outputs/                  # Arquivos convertidos (criado automaticamente)
├── api.py                    # API REST FastAPI
├── main.py                   # Script de linha de comando
├── requirements.txt          # Dependências
└── README.md                 # Este arquivo
```

## ⚠️ Observações

- O script depende da estrutura atual do site iLovePDF. Se o site mudar, pode ser necessário atualizar os seletores CSS.
- Em caso de problemas, verifique se o Chrome está atualizado.
- O ChromeDriver é baixado automaticamente pelo `webdriver-manager`.

## 🐛 Solução de Problemas

### Erro: "Chrome not found"
Certifique-se de que o Google Chrome está instalado no seu sistema.

### Erro: Timeout ao aguardar elemento
O site pode estar lento ou ter mudado sua estrutura. Tente executar novamente.

### Download não completa
Verifique se há espaço suficiente no disco e se o diretório de destino existe.

## 📝 Licença

Este projeto é para uso pessoal e educacional.
# IlovePdfApi
# py-convertpdf
