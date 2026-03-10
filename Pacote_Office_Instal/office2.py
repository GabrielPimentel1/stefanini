import subprocess
import shutil
import sys
import os
from pathlib import Path
import logging

logging.basicConfig(level=logging.INFO, format='%(message)s')


def extrair_pasta_office():
    """Extrai a pasta Office embutida no EXE"""
    try:
        # Quando empacotado com PyInstaller, os arquivos ficam em sys._MEIPASS
        if getattr(sys, 'frozen', False):
            # Rodando como EXE
            pasta_temp = sys._MEIPASS
            pasta_office = Path(pasta_temp) / "Office"

            if pasta_office.exists() and pasta_office.is_dir():
                logging.info(f"✓ Pasta Office encontrada no pacote")
                return pasta_office
            else:
                logging.error("✗ Pasta Office não encontrada no pacote!")
                return None
        else:
            # Rodando como script Python
            pasta_atual = Path(__file__).parent
            pasta_office = pasta_atual / "Office"

            if pasta_office.exists() and pasta_office.is_dir():
                return pasta_office
            else:
                logging.error("✗ Pasta Office não encontrada!")
                return None

    except Exception as e:
        logging.error(f"Erro ao localizar pasta Office: {e}")
        return None


def copiar_pasta_e_executar(pasta_origem):
    """Copia pasta Office para Downloads e executa o OfficeSetup_E3.exe"""
    try:
        pasta_downloads = Path.home() / "Downloads"
        destino = pasta_downloads / "Office"

        # Remove pasta destino se já existir
        if destino.exists():
            logging.info("Removendo instalação anterior...")
            shutil.rmtree(destino)

        # Copia toda a pasta
        logging.info("Copiando pasta Office para Downloads...")
        shutil.copytree(pasta_origem, destino)
        logging.info(f"✓ Pasta copiada para: {destino}")

        # Caminho do instalador específico
        instalador_exe = destino / "OfficeSetup_E3.exe"

        if not instalador_exe.exists():
            logging.error(f"✗ OfficeSetup_E3.exe não encontrado em: {destino}")
            return False

        logging.info(f"Executando instalador: OfficeSetup_E3.exe")
        subprocess.Popen(str(instalador_exe), shell=True, cwd=str(destino))
        logging.info("✓ Instalador iniciado!")

        return True

    except Exception as e:
        logging.error(f"Erro ao copiar/executar: {e}")
        return False


def main():
    print("=" * 60)
    print(" INSTALADOR OFFICE E3")
    print("=" * 60)

    # 1. Verificar admin
    try:
        import ctypes
        if not ctypes.windll.shell32.IsUserAnAdmin():
            print("\n⚠ ERRO: Execute como Administrador!")
            input("\nPressione ENTER para sair...")
            return
    except:
        pass

    # 2. Extrair pasta Office
    print("\n[2/3] Localizando pasta Office no pacote...")
    pasta_office = extrair_pasta_office()

    if not pasta_office:
        print("\n✗ ERRO: Pasta Office não encontrada!")
        input("\nPressione ENTER para sair...")
        return

    # 3. Copiar pasta e executar
    print("\n[3/3] Copiando pasta Office e executando instalador...")
    sucesso = copiar_pasta_e_executar(pasta_office)

    if not sucesso:
        print("\n✗ ERRO: Falha ao executar instalador!")
        input("\nPressione ENTER para sair...")
        return

    print("\n" + "=" * 60)
    print(" ✓ PROCESSO CONCLUÍDO!")
    print("=" * 60)
    print("\nO Office está sendo instalado...")
    print(f"\nPasta copiada para: {Path.home() / 'Downloads' / 'Office'}")

    input("\nPressione ENTER para sair...")


if __name__ == "__main__":
    main()
