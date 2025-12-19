# cert_image_selector.py
import os
import time
from typing import Optional

import pyautogui as pg

pg.FAILSAFE = True
pg.PAUSE = 0.2


def _caminho_imagem(base_dir: str, nome_arquivo: str) -> str:
    """Monta caminho absoluto da imagem do certificado."""
    return os.path.join(base_dir, nome_arquivo)


def selecionar_certificado_por_imagem(
    base_dir_imagens: str,
    nome_img_cert: str,
    nome_img_botao_ok: str = "btn_ok_cert.png",  # ignorado, mantido só pela assinatura
    timeout_cert: int = 30,
    timeout_ok: int = 20,                         # ignorado
    confidence: float = 0.8,
    debug: bool = False,
) -> bool:
    """
    Fluxo:
    1) Espera aparecer na tela a imagem da linha do certificado (nome_img_cert) e clica no centro.
    2) Em seguida, envia a tecla ENTER (no lugar de procurar o botão OK).

    Retorna True se conseguiu clicar no certificado e mandar ENTER, False caso contrário.
    """
    if not nome_img_cert:
        print("[ERRO] nome_img_cert vazio em selecionar_certificado_por_imagem.")
        return False

    caminho_cert = _caminho_imagem(base_dir_imagens, nome_img_cert)

    if not os.path.exists(caminho_cert):
        print(f"[ERRO] Imagem do certificado não encontrada: {caminho_cert}")
        return False

    print(f"[INFO] Vou procurar a imagem do certificado na tela: {caminho_cert}")
    print(f"[INFO] Timeout para achar certificado: {timeout_cert} segundos | confidence={confidence}")

    inicio = time.time()
    tentativa = 0
    pos_cert: Optional[pg.Box] = None

    while time.time() - inicio < timeout_cert:
        tentativa += 1
        if debug:
            print(f"[DEBUG] Tentativa {tentativa} para localizar o certificado...")

        try:
            pos_cert = pg.locateOnScreen(caminho_cert, confidence=confidence)
        except Exception as e:
            print(f"[ERRO] locateOnScreen (certificado) falhou: {e}")
            return False

        if pos_cert:
            if debug:
                print(f"[DEBUG] Certificado encontrado! Bounding box: {pos_cert}")
            x, y = pg.center(pos_cert)
            print(f"[INFO] Certificado localizado, vou clicar em ({x}, {y}).")
            try:
                pg.moveTo(x, y, duration=0.5)
                pg.click()
            except Exception as e:
                print(f"[ERRO] Falha ao clicar na posição do certificado: {e}")
                return False

            # Depois do clique, ENTER para confirmar (OK)
            time.sleep(0.5)
            try:
                pg.press("enter")
                print("[INFO] Tecla ENTER enviada no popup de certificado.")
                return True
            except Exception as e:
                print(f"[ERRO] Falha ao enviar ENTER após clicar no certificado: {e}")
                return False

        # não achou ainda
        if debug:
            print("[DEBUG] Ainda não localizei o certificado, aguardando 1s e tentando de novo...")
        time.sleep(1.0)

    print("[ERRO] Não localizei a imagem do certificado na tela dentro do timeout.")
    return False
