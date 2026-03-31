import pandas as pd
import os
import shutil
import traceback
from datetime import datetime
from logger import log_erro

class ExcelService:
    def __init__(self, arquivo="pacientes.xlsx"):
        self.arquivo = arquivo
        
        # O estado da aplicação (DataFrames) agora vive na instância do serviço
        # eliminando o anti-pattern de variáveis globais (global keyword)
        self.df_pacientes_enf = None
        self.df_completo_enf = None
        self.df_pacientes_uti = None
        self.df_completo_uti = None
        self.df_pacientes_upa = None
        self.df_completo_upa = None

    def criar_excel_padrao(self):
        """Gera um Excel em branco com as colunas corretas se o arquivo não existir."""
        try:
            with pd.ExcelWriter(self.arquivo, engine='openpyxl') as writer:
                cols_enf = ['ENFERMARIA', 'LEITO', 'NOME DO PACIENTE', 'DIETA', 'OBSERVAÇÕES']
                cols_res = ['LEITO', 'NOME DO PACIENTE', 'DIETA', 'OBSERVAÇÕES']
                pd.DataFrame(columns=cols_enf).to_excel(writer, sheet_name="Enfermaria", index=False)
                pd.DataFrame(columns=cols_res).to_excel(writer, sheet_name="UTI", index=False)
                pd.DataFrame(columns=cols_res).to_excel(writer, sheet_name="UPA", index=False)
        except Exception as e: 
            log_erro(f"Erro criar excel: {e}")

    def carregar_dados(self):
        """Lê as abas do Excel, higieniza os dados e popula os DataFrames de estado."""
        if not os.path.exists(self.arquivo): 
            self.criar_excel_padrao()

        try:
            def limpar_leito(val):
                if pd.isna(val) or str(val).strip() == "": return ""
                try: return str(int(float(val)))
                except: return str(val).strip().upper()

            def normalizar_colunas(df):
                df.columns = [str(col).upper().strip() for col in df.columns]
                return df

            # --- 1. ENFERMARIA ---
            try: df_enf = pd.read_excel(self.arquivo, sheet_name="Enfermaria")
            except: 
                try: df_enf = pd.read_excel(self.arquivo, sheet_name=0)
                except: df_enf = pd.DataFrame()

            df_enf = normalizar_colunas(df_enf)
            for col in ['ENFERMARIA', 'LEITO', 'NOME DO PACIENTE', 'DIETA', 'OBSERVAÇÕES']:
                if col not in df_enf.columns: df_enf[col] = ""

            if 'ENFERMARIA' in df_enf.columns: df_enf['ENFERMARIA'] = df_enf['ENFERMARIA'].ffill()
            df_enf['LEITO'] = df_enf['LEITO'].apply(limpar_leito)
            
            self.df_completo_enf = df_enf.copy()
            self.df_pacientes_enf = df_enf.dropna(subset=['NOME DO PACIENTE']).copy()
            self.df_pacientes_enf = self.df_pacientes_enf[self.df_pacientes_enf['NOME DO PACIENTE'].astype(str).str.strip() != '']

            # --- 2. UTI ---
            try: df_uti = pd.read_excel(self.arquivo, sheet_name="UTI")
            except: df_uti = pd.DataFrame()

            df_uti = normalizar_colunas(df_uti)
            for col in ['LEITO', 'NOME DO PACIENTE', 'DIETA', 'OBSERVAÇÕES']:
                if col not in df_uti.columns: df_uti[col] = ""

            df_uti['LEITO'] = df_uti['LEITO'].apply(limpar_leito)
            self.df_completo_uti = df_uti.copy()
            self.df_pacientes_uti = df_uti.dropna(subset=['NOME DO PACIENTE']).copy()
            self.df_pacientes_uti = self.df_pacientes_uti[self.df_pacientes_uti['NOME DO PACIENTE'].astype(str).str.strip() != '']
            
            self.df_pacientes_uti['ENFERMARIA'] = "UTI HRMSS"
            self.df_completo_uti['ENFERMARIA'] = "UTI HRMSS" 

            # --- 3. UPA ---
            try: df_upa = pd.read_excel(self.arquivo, sheet_name="UPA")
            except: df_upa = pd.DataFrame()

            df_upa = normalizar_colunas(df_upa)
            for col in ['LEITO', 'NOME DO PACIENTE', 'DIETA', 'OBSERVAÇÕES']:
                if col not in df_upa.columns: df_upa[col] = ""

            df_upa['LEITO'] = df_upa['LEITO'].apply(limpar_leito)
            self.df_completo_upa = df_upa.copy()
            self.df_pacientes_upa = df_upa.dropna(subset=['NOME DO PACIENTE']).copy()
            self.df_pacientes_upa = self.df_pacientes_upa[self.df_pacientes_upa['NOME DO PACIENTE'].astype(str).str.strip() != '']
            
            self.df_pacientes_upa['ENFERMARIA'] = "UPA"
            self.df_completo_upa['ENFERMARIA'] = "UPA"

            # Retorna para o Pywebview
            return {
                "sucesso": True, 
                "dados_enf": self.df_pacientes_enf.fillna('').to_dict(orient='records'),
                "dados_uti": self.df_pacientes_uti.fillna('').to_dict(orient='records'),
                "dados_upa": self.df_pacientes_upa.fillna('').to_dict(orient='records'),
                "editor_enf": self.df_completo_enf.fillna('').to_dict(orient='records'),
                "editor_uti": self.df_completo_uti.fillna('').to_dict(orient='records'),
                "editor_upa": self.df_completo_upa.fillna('').to_dict(orient='records')
            }
            
        except PermissionError: 
            return {"sucesso": False, "erro": "Excel está aberto em outro programa. Feche-o para continuar."}
        except Exception as e:
            log_erro(traceback.format_exc())
            return {"sucesso": False, "erro": f"Erro de leitura: {str(e)}"}

    def _fazer_backup(self):
        """Cria um backup com timestamp antes de sobrescrever os dados para prevenir perda de informações."""
        try:
            if not os.path.exists(self.arquivo): return
            
            backup_dir = ".backups"
            if not os.path.exists(backup_dir):
                os.makedirs(backup_dir)
                
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            caminho_backup = os.path.join(backup_dir, f"pacientes_bkp_{timestamp}.xlsx")
            shutil.copy2(self.arquivo, caminho_backup)
        except Exception as e:
            log_erro(f"Aviso: Falha ao criar backup: {e}")

    def salvar_dados(self, dados_enf, dados_uti, dados_upa):
        """Grava as modificações de volta no Excel após gerar o backup de segurança."""
        try:
            self._fazer_backup()

            with pd.ExcelWriter(self.arquivo, engine='openpyxl') as writer:
                # Enf
                df_enf = pd.DataFrame(dados_enf)
                cols_enf = ['ENFERMARIA', 'LEITO', 'NOME DO PACIENTE', 'DIETA', 'OBSERVAÇÕES']
                for c in cols_enf: 
                    if c not in df_enf.columns: df_enf[c] = ""
                df_enf[cols_enf].to_excel(writer, sheet_name="Enfermaria", index=False)
                
                # UTI & UPA
                cols_rest = ['LEITO', 'NOME DO PACIENTE', 'DIETA', 'OBSERVAÇÕES']
                df_uti = pd.DataFrame(dados_uti)
                for c in cols_rest: 
                    if c not in df_uti.columns: df_uti[c] = ""
                df_uti[cols_rest].to_excel(writer, sheet_name="UTI", index=False)

                df_upa = pd.DataFrame(dados_upa)
                for c in cols_rest: 
                    if c not in df_upa.columns: df_upa[c] = ""
                df_upa[cols_rest].to_excel(writer, sheet_name="UPA", index=False)

            # Recarrega os DFs em memória
            self.carregar_dados()
            return {"sucesso": True, "msg": "Salvo com sucesso!"}
            
        except PermissionError: 
            return {"sucesso": False, "msg": "Erro: Feche o Excel antes de salvar."}
        except Exception as e: 
            return {"sucesso": False, "msg": f"Erro: {str(e)}"}