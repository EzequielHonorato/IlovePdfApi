#!/usr/bin/env python3
"""
Automação para converter PDF para Word usando o site iLovePDF.
Este script usa Selenium para automatizar o navegador e realizar a conversão.
"""

import os
import sys
import time
from pathlib import Path

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager


class ILovePDFConverter:
    """Classe para automatizar a conversão de PDF para Word no iLovePDF."""
    
    URL_PDF_TO_WORD = "https://www.ilovepdf.com/pt/pdf_para_word"
    
    def __init__(self, download_dir: str = None):
        """
        Inicializa o conversor.
        
        Args:
            download_dir: Diretório para salvar os arquivos convertidos.
                         Se não especificado, usa o diretório padrão de Downloads.
        """
        self.download_dir = download_dir or str(Path.home() / "Downloads")
        self.driver = None
    
    def _setup_driver(self):
        """Configura o WebDriver do Chrome."""
        chrome_options = Options()
        
        # Configurações para download automático
        prefs = {
            "download.default_directory": self.download_dir,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        }
        chrome_options.add_experimental_option("prefs", prefs)
        
        # Descomente a linha abaixo para executar em modo headless (sem interface)
        # chrome_options.add_argument("--headless")
        
        chrome_options.add_argument("--start-maximized")
        chrome_options.add_argument("--disable-notifications")
        
        service = Service(ChromeDriverManager().install())
        self.driver = webdriver.Chrome(service=service, options=chrome_options)
        self.driver.implicitly_wait(10)
    
    def _wait_for_element(self, by: By, value: str, timeout: int = 30):
        """Aguarda um elemento estar presente e visível."""
        wait = WebDriverWait(self.driver, timeout)
        return wait.until(EC.element_to_be_clickable((by, value)))
    
    def _wait_for_download(self, timeout: int = 120):
        """Aguarda o download ser concluído."""
        print("⏳ Aguardando download...")
        end_time = time.time() + timeout
        
        while time.time() < end_time:
            # Verifica se há arquivos .crdownload (download em progresso)
            downloading = any(
                f.endswith('.crdownload') 
                for f in os.listdir(self.download_dir)
            )
            if not downloading:
                # Verifica se um novo arquivo .docx foi criado
                docx_files = [
                    f for f in os.listdir(self.download_dir) 
                    if f.endswith('.docx')
                ]
                if docx_files:
                    return True
            time.sleep(1)
        
        return False
    
    def convert_pdf_to_word(self, pdf_path: str) -> bool:
        """
        Converte um arquivo PDF para Word usando o iLovePDF.
        
        Args:
            pdf_path: Caminho absoluto para o arquivo PDF.
            
        Returns:
            True se a conversão foi bem sucedida, False caso contrário.
        """
        pdf_path = os.path.abspath(pdf_path)
        
        if not os.path.exists(pdf_path):
            print(f"❌ Erro: Arquivo não encontrado: {pdf_path}")
            return False
        
        if not pdf_path.lower().endswith('.pdf'):
            print("❌ Erro: O arquivo deve ter extensão .pdf")
            return False
        
        print(f"📄 Convertendo: {pdf_path}")
        print(f"📁 Salvando em: {self.download_dir}")
        
        try:
            self._setup_driver()
            
            # Abre a página de conversão
            print("🌐 Abrindo iLovePDF...")
            self.driver.get(self.URL_PDF_TO_WORD)
            
            # Aguarda e fecha popup de cookies se existir
            try:
                cookie_btn = WebDriverWait(self.driver, 5).until(
                    EC.element_to_be_clickable((By.ID, "c-p-bn"))
                )
                cookie_btn.click()
                print("🍪 Cookies aceitos")
            except:
                pass  # Popup de cookies pode não aparecer
            
            # Encontra o input de arquivo (está oculto, mas funciona)
            print("📤 Fazendo upload do arquivo...")
            file_input = self.driver.find_element(By.CSS_SELECTOR, "input[type='file']")
            file_input.send_keys(pdf_path)
            
            # Aguarda o upload completar e o botão de conversão aparecer
            print("⏳ Aguardando upload...")
            time.sleep(3)  # Aguarda processamento inicial
            
            # Clica no botão de converter
            print("🔄 Iniciando conversão...")
            convert_btn = self._wait_for_element(
                By.CSS_SELECTOR, 
                "#processTask, button[type='submit'], .process__btn"
            )
            convert_btn.click()
            
            # Aguarda a conversão e o botão de download aparecer
            print("⏳ Processando conversão...")
            download_btn = self._wait_for_element(
                By.CSS_SELECTOR,
                "a.downloader__btn, #downloadFile, .download__btn, a[href*='download']",
                timeout=60
            )
            
            # Clica no botão de download
            print("⬇️ Baixando arquivo convertido...")
            download_btn.click()
            
            # Aguarda o download completar
            if self._wait_for_download():
                print("✅ Conversão concluída com sucesso!")
                return True
            else:
                print("⚠️ Tempo limite de download excedido")
                return False
                
        except Exception as e:
            print(f"❌ Erro durante a conversão: {e}")
            return False
            
        finally:
            if self.driver:
                time.sleep(2)  # Aguarda um pouco antes de fechar
                self.driver.quit()
                print("🔒 Navegador fechado")
    
    def close(self):
        """Fecha o navegador."""
        if self.driver:
            self.driver.quit()


def main():
    """Função principal."""
    if len(sys.argv) < 2:
        print("Uso: python main.py <caminho_do_arquivo.pdf> [diretorio_destino]")
        print("\nExemplo:")
        print("  python main.py documento.pdf")
        print("  python main.py documento.pdf ~/Desktop")
        sys.exit(1)
    
    pdf_path = sys.argv[1]
    download_dir = sys.argv[2] if len(sys.argv) > 2 else None
    
    converter = ILovePDFConverter(download_dir=download_dir)
    success = converter.convert_pdf_to_word(pdf_path)
    
    sys.exit(0 if success else 1)


if __name__ == "__main__":
    main()
