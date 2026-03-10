import subprocess
import os
import sys
import time
import logging
from pathlib import Path
import winreg

# Configurar logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('desinstalacao_office.log'),
        logging.StreamHandler()
    ]
)

class DesinstaladorOffice:
    def __init__(self):
        self.username = "dom"
        self.password = "Stefanini@2026"
        
    def verificar_admin(self):
        """Verifica se está rodando como administrador"""
        try:
            import ctypes
            return ctypes.windll.shell32.IsUserAnAdmin() != 0
        except:
            return False
    
    def parar_processos_office(self):
        """Para todos os processos do Office"""
        logging.info("Parando processos do Office...")
        
        processos = [
            "WINWORD.EXE",
            "EXCEL.EXE",
            "POWERPNT.EXE",
            "OUTLOOK.EXE",
            "ONENOTE.EXE",
            "MSACCESS.EXE",
            "MSPUB.EXE",
            "TEAMS.EXE",
            "lync.exe",
            "OfficeClickToRun.exe",
            "OfficeC2RClient.exe",
            "setup.exe"
        ]
        
        for processo in processos:
            try:
                cmd = f'taskkill /F /IM {processo} /T'
                subprocess.run(cmd, shell=True, capture_output=True, timeout=10)
            except:
                pass
        
        logging.info("✓ Processos finalizados")
    
    def parar_servicos_office(self):
        """Para serviços do Office"""
        logging.info("Parando serviços do Office...")
        
        servicos = [
            "ClickToRunSvc",
            "OfficeSvc",
            "OfficeClientTelemetry"
        ]
        
        for servico in servicos:
            try:
                cmd = f'sc stop {servico}'
                subprocess.run(cmd, shell=True, capture_output=True, timeout=10)
                
                cmd = f'sc config {servico} start= disabled'
                subprocess.run(cmd, shell=True, capture_output=True, timeout=10)
            except:
                pass
        
        logging.info("✓ Serviços parados")
    
    def encontrar_desinstalador_office(self):
        """Procura o desinstalador do Office no sistema"""
        logging.info("Procurando desinstalador do Office...")
        
        # Caminhos comuns do Office
        caminhos_possiveis = [
            r"C:\Program Files\Microsoft Office",
            r"C:\Program Files (x86)\Microsoft Office",
            r"C:\Program Files\Microsoft Office 15",
            r"C:\Program Files (x86)\Microsoft Office 15",
            r"C:\Program Files\Microsoft Office 16",
            r"C:\Program Files (x86)\Microsoft Office 16",
            r"C:\Program Files\Common Files\microsoft shared\ClickToRun",
        ]
        
        for caminho in caminhos_possiveis:
            if os.path.exists(caminho):
                logging.info(f"Office encontrado em: {caminho}")
                return caminho
        
        return None
    
    def desinstalar_via_registro(self):
        """Desinstala Office usando informações do registro"""
        logging.info("Procurando Office no registro do Windows...")
        
        try:
            # Chaves do registro onde o Office pode estar
            chaves_registro = [
                (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"),
                (winreg.HKEY_LOCAL_MACHINE, r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"),
            ]
            
            produtos_office = []
            
            for hkey, caminho in chaves_registro:
                try:
                    chave = winreg.OpenKey(hkey, caminho)
                    i = 0
                    
                    while True:
                        try:
                            subchave_nome = winreg.EnumKey(chave, i)
                            subchave = winreg.OpenKey(chave, subchave_nome)
                            
                            try:
                                nome = winreg.QueryValueEx(subchave, "DisplayName")[0]
                                
                                # Verificar se é Office
                                if "Microsoft Office" in nome or "Microsoft 365" in nome or "Office 365" in nome:
                                    try:
                                        uninstall = winreg.QueryValueEx(subchave, "UninstallString")[0]
                                        produtos_office.append({
                                            'nome': nome,
                                            'uninstall': uninstall
                                        })
                                        logging.info(f"Encontrado: {nome}")
                                    except:
                                        pass
                            except:
                                pass
                            
                            winreg.CloseKey(subchave)
                            i += 1
                            
                        except WindowsError:
                            break
                    
                    winreg.CloseKey(chave)
                    
                except Exception as e:
                    pass
            
            return produtos_office
            
        except Exception as e:
            logging.error(f"Erro ao acessar registro: {e}")
            return []
    
    def desinstalar_office_clicktorun(self):
        """Desinstala Office Click-to-Run (Office 365/2016+)"""
        logging.info("Tentando desinstalar Office Click-to-Run...")
        
        try:
            # Caminho do desinstalador OfficeClickToRun
            caminhos_c2r = [
                r"C:\Program Files\Common Files\microsoft shared\ClickToRun\OfficeClickToRun.exe",
                r"C:\Program Files (x86)\Common Files\microsoft shared\ClickToRun\OfficeClickToRun.exe"
            ]
            
            for caminho in caminhos_c2r:
                if os.path.exists(caminho):
                    logging.info(f"Executando desinstalador: {caminho}")
                    
                    # Comando de desinstalação
                    cmd = f'"{caminho}" scenario=install scenariosubtype=ARP sourcetype=None productstoremove=O365ProPlusRetail.16_en-us_x-none culture=en-us version.16 DisplayLevel=False'
                    
                    subprocess.run(cmd, shell=True, timeout=300)
                    logging.info("✓ Desinstalação Click-to-Run iniciada")
                    return True
            
            return False
            
        except Exception as e:
            logging.error(f"Erro ao desinstalar Click-to-Run: {e}")
            return False
    
    def desinstalar_via_powershell(self):
        """Desinstala Office usando PowerShell"""
        logging.info("Desinstalando Office via PowerShell...")
        
        ps_script = ''' # Parar processos do Office Get-Process | Where-Object {$_.ProcessName -like "*office*" -or $_.ProcessName -like "*excel*" -or $_.ProcessName -like "*word*" -or $_.ProcessName -like "*outlook*"} | Stop-Process -Force -ErrorAction SilentlyContinue # Obter produtos Office instalados $produtos = Get-WmiObject -Class Win32_Product | Where-Object {$_.Name -like "*Microsoft Office*" -or $_.Name -like "*Microsoft 365*"} foreach ($produto in $produtos) { Write-Host "Desinstalando: $($produto.Name)" $produto.Uninstall() | Out-Null } Write-Host "Desinstalacao concluida" '''
        
        try:
            # Salvar script temporário
            script_temp = Path(os.environ['TEMP']) / "desinstalar_office.ps1"
            with open(script_temp, 'w', encoding='utf-8') as f:
                f.write(ps_script)
            
            # Executar PowerShell
            cmd = f'powershell -ExecutionPolicy Bypass -File "{script_temp}"'
            processo = subprocess.Popen(cmd, shell=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
            
            stdout, stderr = processo.communicate(timeout=600)  # 10 minutos
            
            # Remover script temporário
            try:
                script_temp.unlink()
            except:
                pass
            
            logging.info("✓ Desinstalação via PowerShell concluída")
            return True
            
        except Exception as e:
            logging.error(f"Erro ao desinstalar via PowerShell: {e}")
            return False
    
    def limpar_pastas_office(self):
        """Remove pastas do Office"""
        logging.info("Limpando pastas do Office...")
        
        pastas = [
            r"C:\Program Files\Microsoft Office",
            r"C:\Program Files (x86)\Microsoft Office",
            r"C:\Program Files\Microsoft Office 15",
            r"C:\Program Files (x86)\Microsoft Office 15",
            r"C:\Program Files\Microsoft Office 16",
            r"C:\Program Files (x86)\Microsoft Office 16",
            Path.home() / "AppData" / "Local" / "Microsoft" / "Office",
            Path.home() / "AppData" / "Roaming" / "Microsoft" / "Office",
            Path.home() / "AppData" / "Local" / "Microsoft" / "Teams",
        ]
        
        for pasta in pastas:
            try:
                caminho = Path(pasta)
                if caminho.exists():
                    logging.info(f"Removendo: {caminho}")
                    
                    # Usar rmdir do Windows para forçar remoção
                    cmd = f'rmdir /S /Q "{caminho}"'
                    subprocess.run(cmd, shell=True, capture_output=True)
            except Exception as e:
                logging.warning(f"Não foi possível remover {pasta}: {e}")
        
        logging.info("✓ Limpeza de pastas concluída")
    
    def limpar_registro_office(self):
        """Remove entradas do Office no registro"""
        logging.info("Limpando registro do Office...")
        
        chaves = [
            r"HKEY_CURRENT_USER\Software\Microsoft\Office",
            r"HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office",
            r"HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\Office",
        ]
        
        for chave in chaves:
            try:
                cmd = f'reg delete "{chave}" /f'
                subprocess.run(cmd, shell=True, capture_output=True)
            except:
                pass
        
        logging.info("✓ Limpeza do registro concluída")
    
    def usar_ferramenta_microsoft(self):
        """Baixa e usa a ferramenta oficial de desinstalação da Microsoft"""
        logging.info("Baixando ferramenta oficial de desinstalação...")
        
        try:
            # URL da ferramenta SaRA (Support and Recovery Assistant)
            url = "https://aka.ms/SaRA-OfficeUninstall"
            destino = Path(os.environ['TEMP']) / "SaRA_OfficeUninstall.exe"
            
            # Baixar ferramenta
            ps_download = f'Invoke-WebRequest -Uri "{url}" -OutFile "{destino}"'
            subprocess.run(f'powershell -Command "{ps_download}"', shell=True, timeout=120)
            
            if destino.exists():
                logging.info("Executando ferramenta Microsoft...")
                subprocess.Popen(str(destino), shell=True)
                return True
            
            return False
            
        except Exception as e:
            logging.error(f"Erro ao usar ferramenta Microsoft: {e}")
            return False
    
    def executar(self):
        """Executa o processo completo de desinstalação"""
        logging.info("=" * 70)
        logging.info(" DESINSTALADOR COMPLETO DO OFFICE")
        logging.info("=" * 70)
        
        # Verificar privilégios
        if not self.verificar_admin():
            logging.error("⚠ ATENÇÃO: Este script precisa ser executado como Administrador!")
            input("\nPressione ENTER para sair...")
            return False
        
        print("\n⚠ ATENÇÃO: Este processo irá desinstalar completamente o Microsoft Office!")
        print("Todos os programas do Office serão removidos.")
        resposta = input("\nDeseja continuar? (S/N): ")
        
        if resposta.upper() != 'S':
            logging.info("Desinstalação cancelada pelo usuário")
            return False
        
        # Passo 1: Parar processos
        logging.info("\n[1/7] Parando processos do Office...")
        self.parar_processos_office()
        time.sleep(2)
        
        # Passo 2: Parar serviços
        logging.info("\n[2/7] Parando serviços do Office...")
        self.parar_servicos_office()
        time.sleep(2)
        
        # Passo 3: Procurar no registro
        logging.info("\n[3/7] Procurando Office no registro...")
        produtos = self.desinstalar_via_registro()
        
        if produtos:
            logging.info(f"Encontrados {len(produtos)} produtos Office")
            for produto in produtos:
                logging.info(f" - {produto['nome']}")
        
        # Passo 4: Desinstalar Click-to-Run
        logging.info("\n[4/7] Desinstalando Office Click-to-Run...")
        self.desinstalar_office_clicktorun()
        time.sleep(5)
        
        # Passo 5: Desinstalar via PowerShell
        logging.info("\n[5/7] Desinstalando via PowerShell...")
        self.desinstalar_via_powershell()
        time.sleep(3)
        
        # Passo 6: Limpar pastas
        logging.info("\n[6/7] Limpando pastas do Office...")
        self.limpar_pastas_office()
        
        # Passo 7: Limpar registro
        logging.info("\n[7/7] Limpando registro...")
        self.limpar_registro_office()
        
        logging.info("\n" + "=" * 70)
        logging.info(" ✓ DESINSTALAÇÃO CONCLUÍDA!")
        logging.info("=" * 70)
        logging.info("\nRecomendações:")
        logging.info(" 1. Reinicie o computador")
        logging.info(" 2. Verifique se algum programa do Office ainda está instalado")
        logging.info(" 3. Execute a ferramenta oficial da Microsoft se necessário")
        
        input("\nPressione ENTER para sair...")
        return True


if __name__ == "__main__":
    desinstalador = DesinstaladorOffice()
    desinstalador.executar()