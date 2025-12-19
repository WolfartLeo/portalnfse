# ocr_ui.py
import time
from typing import Optional, Tuple, List, Dict

import pyautogui as pg
from PIL import Image
import pytesseract

# Ajuste aqui o caminho do executável do Tesseract no seu Windows, se for diferente:
# Exemplo padrão:
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

pg.FAILSAFE = True
pg.PAUSE = 0.2


def _ocr_linhas(im: Image.Image) -> List[Dict]:
    """
    Roda OCR na imagem e devolve uma lista de linhas com bounding box consolidado.

    Cada item:
    {
        'text': 'texto da linha',
        'left': int,
        'top': int,
        'right': int,
        'bottom': int
    }
    """
    # Sem lang="por" porque o pacote por.traineddata não está instalado;
    # o idioma padrão (geralmente eng) é suficiente para ler nomes e "OK".
    data = pytesseract.image_to_data(
        im,
        output_type=pytesseract.Output.DICT
    )

    linhas: Dict[int, Dict] = {}
    n = len(data["text"])

    for i in range(n):
        txt = (data["text"][i] or "").strip()
        if not txt:
            continue

        line_num = data["line_num"][i]
        left = data["left"][i]
        top = data["top"][i]
        w = data["width"][i]
        h = data["height"][i]

        if line_num not in linhas:
            linhas[line_num] = {
                "text": txt,
                "left": left,
                "top": top,
                "right": left + w,
                "bottom": top + h,
            }
        else:
            linhas[line_num]["text"] += " " + txt
            linhas[line_num]["left"] = min(linhas[line_num]["left"], left)
            linhas[line_num]["top"] = min(linhas[line_num]["top"], top)
            linhas[line_num]["right"] = max(linhas[line_num]["right"], left + w)
            linhas[line_num]["bottom"] = max(linhas[line_num]["bottom"], top + h)

    # retorna lista ordenada por posição vertical
    return sorted(linhas.values(), key=lambda x: x["top"])


def clicar_texto_na_tela(
    texto_alvo: str,
    timeout: int = 15,
    region: Optional[Tuple[int, int, int, int]] = None,
    debug: bool = False,
) -> bool:
    """
    Procura um TEXTO na tela via OCR e clica no centro da linha onde esse texto aparece.

    texto_alvo: string que deve aparecer na linha (case-insensitive).
    region: (left, top, width, height) ou None pra tela toda.
    """
    texto_alvo_norm = texto_alvo.strip().lower()
    if not texto_alvo_norm:
        print("[ERRO] texto_alvo vazio em clicar_texto_na_tela.")
        return False

    inicio = time.time()
    while (time.time() - inicio) < timeout:
        # screenshot
        im = pg.screenshot(region=region)

        # OCR por linhas
        linhas = _ocr_linhas(im)

        for linha in linhas:
            linha_txt_norm = linha["text"].lower()
            if texto_alvo_norm in linha_txt_norm:
                # bounding box da linha
                left = linha["left"]
                top = linha["top"]
                right = linha["right"]
                bottom = linha["bottom"]

                # se tiver region, somar offset
                if region is not None:
                    rx, ry, _, _ = region
                    left += rx
                    right += rx
                    top += ry
                    bottom += ry

                x_centro = (left + right) // 2
                y_centro = (top + bottom) // 2

                if debug:
                    print(f"[DEBUG] Encontrado texto '{texto_alvo}' na linha: {linha['text']}")
                    print(f"[DEBUG] BBox linha: left={left}, top={top}, right={right}, bottom={bottom}")
                    print(f"[DEBUG] Clicando em ({x_centro}, {y_centro})")

                try:
                    pg.moveTo(x_centro, y_centro, duration=0.3)
                    pg.click()
                    return True
                except Exception as e:
                    print(f"[ERRO] Falha ao clicar no texto '{texto_alvo}': {e}")
                    return False

        time.sleep(0.8)

    print(f"[ERRO] Não encontrei o texto '{texto_alvo}' na tela dentro do timeout.")
    return False
