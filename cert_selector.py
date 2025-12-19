# cert_selector.py
import time
from typing import Optional

from pywinauto import Desktop

POSSIVEIS_TITULOS = [
    "Selecione um certificado",
    "Selecionar um certificado",
    "Select a Certificate",
]


def _encontrar_janela_certificado(timeout: int = 15):
    """
    Varre todas as janelas e retorna aquela que contém algum controle com
    o texto 'Selecione um certificado' (ou similar).
    """
    desktop = Desktop(backend="uia")
    inicio = time.time()
    janela_alvo = None

    while (time.time() - inicio) < timeout and janela_alvo is None:
        for w in desktop.windows():
            try:
                # varre todos os controles dentro dessa janela
                for ctrl in w.descendants():
                    try:
                        txt = ctrl.window_text() or ""
                    except Exception:
                        continue

                    for alvo in POSSIVEIS_TITULOS:
                        if alvo in txt:
                            janela_alvo = w
                            break

                    if janela_alvo is not None:
                        break
            except Exception:
                continue

            if janela_alvo is not None:
                break

        if janela_alvo is None:
            time.sleep(0.5)

    return janela_alvo


def selecionar_certificado(ident_cert: str, timeout: int = 15, debug: bool = False) -> bool:
    """
    Seleciona o certificado pelo TEXTO que aparece na linha da lista (coluna Tema).

    Fluxo:
      1) Selenium já clicou em 'Acesso via certificado digital'.
      2) Aqui:
         - encontra a janela que contém 'Selecione um certificado';
         - procura qualquer controle cujo texto contenha ident_cert;
         - clica nele;
         - procura qualquer botão com texto 'OK', 'Confirmar', etc. e clica.

    ident_cert: trecho que aparece na linha do certificado
                ex.: 'INNOVE CONTABILIDADE E SOLL'
                    'LCG CONTABILIDADE SS LTDA:36'
    """

    ident_cert = (ident_cert or "").strip()
    if not ident_cert:
        print("[ERRO] IDENT_CERT vazio ao selecionar certificado.")
        return False

    # 1) achar a janela do certificado
    janela = _encontrar_janela_certificado(timeout=timeout)
    if janela is None:
        print("[ERRO] Não encontrei nenhuma janela com 'Selecione um certificado' dentro do timeout.")
        return False

    try:
        janela.set_focus()
    except Exception:
        pass

    if debug:
        try:
            titulo = janela.window_text()
        except Exception:
            titulo = ""
        print(f"[DEBUG] Janela de certificado localizada. Título: {repr(titulo)}")

    # 2) achar o item do certificado pelo texto
    item_alvo = None
    try:
        for ctrl in janela.descendants():
            try:
                txt = (ctrl.window_text() or "").strip()
            except Exception:
                continue

            if not txt:
                continue

            if debug and ident_cert.lower() in txt.lower():
                print(f"[DEBUG] Match de IDENT_CERT em: {repr(txt)}")

            if ident_cert.lower() in txt.lower():
                item_alvo = ctrl
                break
    except Exception as e:
        print(f"[ERRO] Exceção ao varrer itens da janela: {e}")
        return False

    if item_alvo is None:
        print(f"[ERRO] Nenhum controle na UI contém o texto IDENT_CERT='{ident_cert}'.")
        if not debug:
            print("[DICA] Confirme se o IDENT_CERT na planilha está escrito igual ao texto que aparece na coluna Tema.")
        return False

    # 3) clicar / selecionar o item do certificado
    try:
        try:
            item_alvo.select()
        except Exception:
            item_alvo.click_input()
    except Exception as e:
        print(f"[ERRO] Falha ao clicar no item do certificado: {e}")
        return False

    time.sleep(0.5)

    # 4) localizar botão OK e clicar
    botao_ok = None
    try:
        for ctrl in janela.descendants():
            try:
                txt_btn = (ctrl.window_text() or "").strip().lower()
            except Exception:
                continue

            if txt_btn in ("ok", "ok.", "continuar", "confirmar", "concluir"):
                botao_ok = ctrl
                break
    except Exception:
        pass

    if botao_ok is None:
        print("[ERRO] Botão OK/Confirmar não encontrado na janela do certificado.")
        return False

    try:
        botao_ok.click_input()
    except Exception as e:
        print(f"[ERRO] Falha ao clicar no botão OK do certificado: {e}")
        return False

    time.sleep(1.0)
    print(f"[INFO] Certificado selecionado com sucesso para IDENT_CERT='{ident_cert}'.")
    return True
