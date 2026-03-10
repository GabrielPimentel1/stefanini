import subprocess
import shutil
import sys
import os
import tempfile
from pathlib import Path
import logging

logging.basicConfig(level=logging.INFO, format='%(message)s')


def extrair_instalador():
    """Extrai o instalador embutido no EXE"""
    try:
        # Quando empacotado com PyInstaller, os arquivos ficam em sys._MEIPASS
        if getattr(sys, 'frozen', False):
            # Rodando como EXE
            pasta_temp = sys._MEIPASS
            instalador_temp = Path(pasta_temp) / "OfficeSetup_E3.exe"

            if instalador_temp.exists():
                logging.info(f"✓ Instalador encontrado no pacote")
                return instalador_temp
            else:
                logging.error("✗ Instalador não encontrado no pacote!")
                return None
        else:
            # Rodando como script Python
            pasta_atual = Path(__file__).parent
            instalador = pasta_atual / "OfficeSetup_E3.exe"

            if instalador.exists():
                return instalador
            else:
                logging.error("✗ Instalador não encontrado!")
                return None

    except Exception as e:
        logging.error(f"Erro ao extrair instalador: {e}")
        return None


def copiar_e_executar_instalador(instalador_origem):
    """Copia instalador para Downloads e executa"""
    pasta_downloads = Path.home() / "Downloads"
    destino = pasta_downloads / "OfficeSetup_E3.exe"

    logging.info("Copiando instalador para Downloads...")
    shutil.copy2(instalador_origem, destino)
    logging.info(f"✓ Copiado para: {destino}")

    logging.info("Executando instalador...")
    subprocess.Popen(str(destino), shell=True)
    logging.info("✓ Instalador iniciado!")

    return True


def main():
    print("=" * 60)
    print(" INSTALADOR OFFICE E3")
    print("=" * 60)

    # Verificar admin
    try:
        import ctypes
        if not ctypes.windll.shell32.IsUserAnAdmin():
            print("\n⚠ ERRO: Execute como Administrador!")
            input("\nPressione ENTER para sair...")
            return
    except:
        pass

    # 2. Extrair instalador
    print("\n[2/3] Extraindo instalador do pacote...")
    instalador = extrair_instalador()

    if not instalador:
        print("\n✗ ERRO: Instalador não encontrado!")
        input("\nPressione ENTER para sair...")
        return

    # 3. Copiar e executar
    print("\n[3/3] Copiando e executando instalador...")
    copiar_e_executar_instalador(instalador)

    print("\n" + "=" * 60)
    print(" ✓ PROCESSO CONCLUÍDO!")
    print("=" * 60)
    print("\nO Office está sendo instalado...")

    input("\nPressione ENTER para sair...")


if __name__ == "__main__":
    main()
