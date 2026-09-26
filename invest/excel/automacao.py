"""Etapa feita pelo próprio Excel (automação COM, Windows): tabelas dinâmicas, macros e botões; salva como .xlsm.

Para inserir macros, o Excel exige "Confiar no acesso ao modelo de objeto do projeto VBA" (AccessVBOM).
A opção é ligada só durante esta etapa e a configuração original do usuário é sempre restaurada.
"""
import time
import winreg
from contextlib import contextmanager
from pathlib import Path

XL_XLSM = 52
CHAVE = r"Software\Microsoft\Office\16.0\Excel\Security"


@contextmanager
def acesso_vba_temporario():
    """Liga AccessVBOM e devolve o valor original ao final (inclusive em caso de erro)."""
    k = winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, CHAVE, 0, winreg.KEY_READ | winreg.KEY_WRITE)
    try:
        original = winreg.QueryValueEx(k, "AccessVBOM")[0]
    except FileNotFoundError:
        original = None
    winreg.SetValueEx(k, "AccessVBOM", 0, winreg.REG_DWORD, 1)
    try:
        yield
    finally:
        if original is None:
            winreg.DeleteValue(k, "AccessVBOM")
        else:
            winreg.SetValueEx(k, "AccessVBOM", 0, winreg.REG_DWORD, original)
        winreg.CloseKey(k)


def _tentar(f, vezes=90):
    """O Excel recusa chamadas enquanto calcula/atualiza; espera e tenta de novo."""
    for _ in range(vezes):
        try:
            return f()
        except Exception as e:
            if "-2147418111" not in str(e) and "rejeitada" not in str(e).lower():
                raise
            time.sleep(1)
    return f()


def finalizar(origem: Path, destino: Path, linhas_botoes, dinamicas=True, macros=True):
    """Abre o .xlsx gerado pelo openpyxl, aplica as etapas do Excel e salva como .xlsm em destino."""
    try:
        import pythoncom
        import win32com.client as win32
    except ImportError:
        return ["etapa do Excel ignorada (pywin32 não instalado)"], False
    from . import dinamicas as dyn
    from . import macros as mac
    from . import navegacao as nav

    msgs = []
    pythoncom.CoInitialize()
    xl = None
    ctx = acesso_vba_temporario() if macros else _nada()
    try:
        with ctx:
            xl = win32.DispatchEx("Excel.Application")          # instância própria: não mexe no Excel do usuário
            xl.Visible = xl.DisplayAlerts = False
            wb = _tentar(lambda: xl.Workbooks.Open(str(origem)))
            _tentar(xl.CalculateFull)
            # formas (menu e botões) antes das tabelas dinâmicas: criá-las depois derruba o Excel
            msgs.append(nav.inserir(wb))
            if macros:
                msgs.append(mac.inserir(wb, linhas_botoes))
            if dinamicas:
                msgs.append(dyn.criar(xl, wb))
            wb.Worksheets("Painel").Activate()
            if destino.exists():
                destino.unlink()
            _tentar(lambda: wb.SaveAs(str(destino), FileFormat=XL_XLSM))
            _tentar(lambda: wb.Close(False))
        return msgs, True
    except Exception as e:
        msgs.append(f"etapa do Excel falhou ({e})")
        return msgs, False
    finally:
        if xl is not None:
            try:
                xl.Quit()
            except Exception:
                pass
        pythoncom.CoUninitialize()


@contextmanager
def _nada():
    yield
