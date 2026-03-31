import webview
import os
import traceback
from excel_service import ExcelService
from pdf_service import PdfService
from logger import log_erro

class ApiController:
    """Controller que intermedeia os chamados em JavaScript (UI) com as Regras de Negócio (Python)."""
    
    def __init__(self):
        self.excel_service = ExcelService()

    def carregar_dados_excel(self):
        return self.excel_service.carregar_dados()

    def salvar_dados_excel(self, dados_enf, dados_uti, dados_upa):
        return self.excel_service.salvar_dados(dados_enf, dados_uti, dados_upa)

    def _pedir_caminho_salvar(self, nome_sugerido):
        """Abre o explorador de arquivos do sistema operacional nativamente para escolha de pasta."""
        try:
            caminho = webview.windows[0].create_file_dialog(
                webview.SAVE_DIALOG, 
                directory='', 
                save_filename=nome_sugerido, 
                file_types=('Arquivos PDF (*.pdf)',)
            )
            if not caminho: return None
            if isinstance(caminho, (tuple, list)):
                if len(caminho) > 0: caminho = caminho[0]
                else: return None
            caminho = str(caminho)
            if not caminho.lower().endswith('.pdf'): caminho += '.pdf'
            return caminho
        except Exception as e:
            log_erro(f"Erro diálogo: {e}")
            return None

    def imprimir_etiquetas(self, lista_pacientes):
        if not lista_pacientes: return "Fila vazia!"
        
        caminho = self._pedir_caminho_salvar("etiquetas.pdf")
        if not caminho: return "Cancelado."

        try:
            PdfService.gerar_etiquetas(lista_pacientes, caminho)
            os.startfile(caminho)
            return "PDF de etiquetas salvo e aberto com sucesso!"
        except PermissionError: 
            return "ERRO: O PDF atual está aberto em algum visualizador. Feche-o para gerar um novo."
        except Exception as e: 
            log_erro(traceback.format_exc())
            return f"Erro de geração: {e}"

    # --- RELATÓRIOS ---
    def gerar_relatorio_enf(self, tipo):
        df = self.excel_service.df_pacientes_enf if tipo == 'simples' else self.excel_service.df_completo_enf
        
        if df is None or df.empty: return "Sem dados de Enfermarias para gerar."
        
        if 'ENFERMARIA' in df.columns:
            try:
                df = df.sort_values(by=['ENFERMARIA', 'LEITO'])
            except: pass

        nome_arquivo = f"relatorio_enf_{tipo}.pdf"
        titulo = "PACIENTES OCUPADOS" if tipo == 'simples' else "MAPA GERAL (AUDITORIA)"
        
        caminho = self._pedir_caminho_salvar(nome_arquivo)
        if not caminho: return "Cancelado."

        try:
            PdfService.gerar_relatorio_mesclado(df, caminho, titulo)
            os.startfile(caminho)
            return "Relatório gerado com sucesso!"
        except PermissionError: 
            return "ERRO: O arquivo PDF está aberto. Feche-o."
        except Exception as e: 
            log_erro(traceback.format_exc())
            return f"Erro ao gerar: {e}"

    def gerar_relatorio_uti(self, tipo):
        df = self.excel_service.df_pacientes_uti if tipo == 'simples' else self.excel_service.df_completo_uti
        if df is None or df.empty: return "Sem dados de UTI para gerar."
        
        caminho = self._pedir_caminho_salvar(f"relatorio_uti_{tipo}.pdf")
        if not caminho: return "Cancelado."

        try:
            titulo = "NUTRIÇÃO - CORRIDA DE LEITO - UTI HRMSS"
            PdfService.gerar_relatorio_simples(df, caminho, titulo)
            os.startfile(caminho)
            return "Relatório da UTI gerado com sucesso!"
        except PermissionError: return "ERRO: O arquivo PDF está aberto. Feche-o."
        except Exception as e: return f"Erro ao gerar: {e}"

    def gerar_relatorio_upa(self, tipo):
        df = self.excel_service.df_pacientes_upa if tipo == 'simples' else self.excel_service.df_completo_upa
        if df is None or df.empty: return "Sem dados da UPA para gerar."
        
        caminho = self._pedir_caminho_salvar(f"relatorio_upa_{tipo}.pdf")
        if not caminho: return "Cancelado."

        try:
            titulo = "NUTRIÇÃO - CORRIDA DE LEITO - UPA / SALA VERMELHA / AMARELA"
            PdfService.gerar_relatorio_simples(df, caminho, titulo)
            os.startfile(caminho)
            return "Relatório da UPA gerado com sucesso!"
        except PermissionError: return "ERRO: O arquivo PDF está aberto. Feche-o."
        except Exception as e: return f"Erro ao gerar: {e}"


if __name__ == '__main__':
    api = ApiController()
    
    # Inicia a UI do Webview conectada ao nosso Controller em Python
    webview.create_window(
        title='Sistema NutriBem +', 
        url='web/index.html', 
        js_api=api, 
        width=1200, 
        height=800, 
        resizable=True
    )
    webview.start()