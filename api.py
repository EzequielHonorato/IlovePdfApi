#!/usr/bin/env python3
"""
API REST para converter PDF para Word usando o site iLovePDF.
Fornece endpoints para upload, status e download do arquivo convertido.
"""

import os
import uuid
import time
import threading
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, UploadFile, File, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager


# ============ Configuração ============
UPLOAD_DIR = Path("./uploads")
OUTPUT_DIR = Path("./outputs")
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

# Armazena o status das conversões
# Em produção, use Redis ou banco de dados
conversions: dict = {}


# ============ Modelos ============
class ConversionStatus(BaseModel):
    id: str
    status: str  # "pending", "processing", "completed", "error"
    message: Optional[str] = None
    url: Optional[str] = None
    filename: Optional[str] = None


# ============ API ============
app = FastAPI(
    title="iLovePDF Converter API",
    description="API para converter PDF para Word usando iLovePDF",
    version="1.0.0"
)

# CORS - permite requisições de qualquer origem
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# ============ Conversor ============
class ILovePDFConverter:
    """Classe para automatizar a conversão de PDF para Word no iLovePDF."""
    
    URL_PDF_TO_WORD = "https://www.ilovepdf.com/pt/pdf_para_word"
    
    def __init__(self, download_dir: str):
        self.download_dir = download_dir
        self.driver = None
    
    def _setup_driver(self):
        """Configura o WebDriver do Chrome."""
        chrome_options = Options()
        
        prefs = {
            "download.default_directory": self.download_dir,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True,
            "plugins.always_open_pdf_externally": True
        }
        chrome_options.add_experimental_option("prefs", prefs)
        chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
        chrome_options.add_experimental_option("useAutomationExtension", False)
        
        # Modo headless NOVO para Chrome 109+
        chrome_options.add_argument("--headless=new")
        chrome_options.add_argument("--no-sandbox")
        chrome_options.add_argument("--disable-dev-shm-usage")
        chrome_options.add_argument("--disable-gpu")
        chrome_options.add_argument("--window-size=1920,1080")
        chrome_options.add_argument("--disable-notifications")
        chrome_options.add_argument("--disable-extensions")
        chrome_options.add_argument("--disable-blink-features=AutomationControlled")
        chrome_options.add_argument("--user-agent=Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/144.0.0.0 Safari/537.36")
        
        # Usa webdriver-manager para pegar a versão correta
        service = Service(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=service, options=chrome_options)
        self.driver.implicitly_wait(10)
        
        # Configuração extra para download em headless
        self.driver.execute_cdp_cmd("Page.setDownloadBehavior", {
            "behavior": "allow",
            "downloadPath": self.download_dir
        })
    
    def _wait_for_element(self, by: By, value: str, timeout: int = 30):
        """Aguarda um elemento estar presente e visível."""
        wait = WebDriverWait(self.driver, timeout)
        return wait.until(EC.element_to_be_clickable((by, value)))
    
    def _wait_for_download(self, timeout: int = 120) -> Optional[str]:
        """Aguarda o download ser concluído e retorna o nome do arquivo."""
        end_time = time.time() + timeout
        
        while time.time() < end_time:
            files = os.listdir(self.download_dir)
            
            # Verifica se há arquivos .crdownload (download em progresso)
            downloading = any(f.endswith('.crdownload') for f in files)
            
            if not downloading:
                # Procura por arquivos .docx
                docx_files = [f for f in files if f.endswith('.docx')]
                if docx_files:
                    # Retorna o arquivo mais recente
                    docx_files.sort(
                        key=lambda x: os.path.getmtime(os.path.join(self.download_dir, x)),
                        reverse=True
                    )
                    return docx_files[0]
            
            time.sleep(1)
        
        return None
    
    def convert(self, pdf_path: str) -> Optional[str]:
        """
        Converte um arquivo PDF para Word.
        
        Returns:
            Nome do arquivo convertido ou None em caso de erro.
        """
        try:
            self._setup_driver()
            print(f"[INFO] Driver iniciado, convertendo: {pdf_path}")
            
            # Abre a página de conversão
            self.driver.get(self.URL_PDF_TO_WORD)
            print("[INFO] Página carregada")
            
            # Fecha popup de cookies se existir
            try:
                cookie_btn = WebDriverWait(self.driver, 5).until(
                    EC.element_to_be_clickable((By.ID, "c-p-bn"))
                )
                cookie_btn.click()
                print("[INFO] Cookies aceitos")
            except:
                pass
            
            # Upload do arquivo
            file_input = self.driver.find_element(By.CSS_SELECTOR, "input[type='file']")
            file_input.send_keys(pdf_path)
            print("[INFO] Arquivo enviado")
            
            # Aguarda upload completar
            time.sleep(5)
            
            # Clica no botão de converter
            try:
                convert_btn = self._wait_for_element(
                    By.CSS_SELECTOR, 
                    "#processTask"
                )
                convert_btn.click()
                print("[INFO] Botão converter clicado")
            except Exception as e:
                print(f"[WARN] Tentando seletor alternativo: {e}")
                convert_btn = self._wait_for_element(
                    By.XPATH,
                    "//button[contains(@class, 'process')]"
                )
                convert_btn.click()
            
            # Aguarda botão de download
            print("[INFO] Aguardando conversão...")
            download_btn = self._wait_for_element(
                By.CSS_SELECTOR,
                "a.downloader__btn, #downloadFile, .download__btn",
                timeout=90
            )
            print("[INFO] Conversão concluída, iniciando download")
            
            # Clica no download
            download_btn.click()
            
            # Aguarda download completar
            filename = self._wait_for_download()
            print(f"[INFO] Download concluído: {filename}")
            return filename
                
        except Exception as e:
            print(f"Erro na conversão: {e}")
            # Salva screenshot para debug
            try:
                self.driver.save_screenshot(f"{self.download_dir}/error_screenshot.png")
                print(f"[DEBUG] Screenshot salvo em {self.download_dir}/error_screenshot.png")
            except:
                pass
            return None
            
        finally:
            if self.driver:
                time.sleep(2)
                self.driver.quit()


# ============ Funções de Background ============
def process_conversion(conversion_id: str, pdf_path: str):
    """Processa a conversão em background."""
    try:
        # Atualiza status para processing
        conversions[conversion_id]["status"] = "processing"
        conversions[conversion_id]["message"] = "Convertendo PDF para Word..."
        
        # Configura diretório de saída
        output_dir = str(OUTPUT_DIR.absolute())
        
        # Executa conversão
        converter = ILovePDFConverter(download_dir=output_dir)
        filename = converter.convert(pdf_path)
        
        if filename:
            # Sucesso
            conversions[conversion_id]["status"] = "completed"
            conversions[conversion_id]["message"] = "Conversão concluída!"
            conversions[conversion_id]["filename"] = filename
            conversions[conversion_id]["url"] = f"/download/{conversion_id}"
        else:
            # Erro
            conversions[conversion_id]["status"] = "error"
            conversions[conversion_id]["message"] = "Falha na conversão"
            
    except Exception as e:
        conversions[conversion_id]["status"] = "error"
        conversions[conversion_id]["message"] = str(e)
    
    finally:
        # Remove o PDF original após processamento
        try:
            os.remove(pdf_path)
        except:
            pass


# ============ Endpoints ============
@app.get("/")
async def root():
    """Endpoint raiz."""
    return {
        "message": "iLovePDF Converter API",
        "docs": "/docs",
        "endpoints": {
            "POST /convert": "Envia PDF para conversão",
            "GET /status/{id}": "Verifica status da conversão",
            "GET /download/{id}": "Baixa o arquivo convertido"
        }
    }


@app.post("/convert", response_model=ConversionStatus)
async def convert_pdf(file: UploadFile = File(...)):
    """
    Envia um arquivo PDF para conversão.
    
    Retorna um ID para acompanhar o status.
    """
    # Valida o arquivo
    if not file.filename.lower().endswith('.pdf'):
        raise HTTPException(status_code=400, detail="O arquivo deve ser um PDF")
    
    # Gera ID único
    conversion_id = str(uuid.uuid4())
    
    # Salva o arquivo
    pdf_path = UPLOAD_DIR / f"{conversion_id}.pdf"
    
    try:
        content = await file.read()
        with open(pdf_path, "wb") as f:
            f.write(content)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Erro ao salvar arquivo: {e}")
    
    # Inicializa o status
    conversions[conversion_id] = {
        "id": conversion_id,
        "status": "pending",
        "message": "Aguardando processamento...",
        "url": None,
        "filename": None,
        "original_filename": file.filename
    }
    
    # Inicia conversão em background
    thread = threading.Thread(
        target=process_conversion,
        args=(conversion_id, str(pdf_path.absolute()))
    )
    thread.start()
    
    return ConversionStatus(
        id=conversion_id,
        status="pending",
        message="Conversão iniciada. Use /status/{id} para acompanhar."
    )


@app.get("/status/{conversion_id}", response_model=ConversionStatus)
async def get_status(conversion_id: str):
    """
    Retorna o status atual da conversão.
    
    Status possíveis:
    - pending: Aguardando processamento
    - processing: Convertendo...
    - completed: Concluído (url disponível)
    - error: Erro na conversão
    """
    if conversion_id not in conversions:
        raise HTTPException(status_code=404, detail="Conversão não encontrada")
    
    conv = conversions[conversion_id]
    
    return ConversionStatus(
        id=conv["id"],
        status=conv["status"],
        message=conv.get("message"),
        url=conv.get("url"),
        filename=conv.get("filename")
    )


@app.get("/download/{conversion_id}")
async def download_file(conversion_id: str):
    """
    Baixa o arquivo Word convertido.
    """
    if conversion_id not in conversions:
        raise HTTPException(status_code=404, detail="Conversão não encontrada")
    
    conv = conversions[conversion_id]
    
    if conv["status"] != "completed":
        raise HTTPException(
            status_code=400, 
            detail=f"Conversão não concluída. Status: {conv['status']}"
        )
    
    filename = conv.get("filename")
    if not filename:
        raise HTTPException(status_code=404, detail="Arquivo não encontrado")
    
    file_path = OUTPUT_DIR / filename
    
    if not file_path.exists():
        raise HTTPException(status_code=404, detail="Arquivo não encontrado no servidor")
    
    # Nome original sem .pdf + .docx
    original = conv.get("original_filename", "documento.pdf")
    download_name = original.rsplit('.', 1)[0] + ".docx"
    
    return FileResponse(
        path=str(file_path),
        filename=download_name,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )


@app.get("/conversions")
async def list_conversions():
    """Lista todas as conversões (para debug)."""
    return {"conversions": list(conversions.values())}


@app.delete("/conversion/{conversion_id}")
async def delete_conversion(conversion_id: str):
    """Remove uma conversão e seus arquivos."""
    if conversion_id not in conversions:
        raise HTTPException(status_code=404, detail="Conversão não encontrada")
    
    conv = conversions[conversion_id]
    
    # Remove arquivo de saída
    if conv.get("filename"):
        try:
            os.remove(OUTPUT_DIR / conv["filename"])
        except:
            pass
    
    # Remove do registro
    del conversions[conversion_id]
    
    return {"message": "Conversão removida"}


# ============ Main ============
if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
