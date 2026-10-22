"""VBA-Steuerung einbetten: Mappe als .xlsm mit den Modulen aus prognosemodell/vba.

openpyxl kann kein VBA-Projekt erzeugen, nur ein vorhandenes übernehmen. Deshalb
baut LibreOffice (headless, per UNO) aus den Quelltexten ein VBA-Projekt
(vbaProject.bin) in einer leeren Hilfsmappe mit denselben Codenamen; openpyxl
übernimmt daraus nur das VBA-Projekt und schreibt die eigentliche Mappe wie bisher.
LibreOffice wird nur beim Bauen gebraucht, die fertige Datei braucht nur Excel.

Die UNO-Brücke (Modul uno) bringt unter Windows und macOS nur das Python mit, das
in LibreOffice steckt. Fehlt sie im aufrufenden Python, baut dieses LibreOffice-Python
die Hilfsmappe in einem eigenen Prozess.

Dieselbe LibreOffice-Sitzung führt im Prüfskript die Makros aus.
"""

import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
import uuid
import zipfile
from pathlib import Path

from .modelle import CODENAME_MAPPE

VBA_ORDNER = Path(__file__).parent / "vba"
MODULE = ["modStart", "modRechnen", "modPruefung", "modObjekte", "modVarianten"]
PROJEKTNAME = "VBAProject"
FILTER_XLSM = "Calc MS Excel 2007 VBA XML"
FILTER_XLSX = "Calc MS Excel 2007 XML"
# com.sun.star.script.ModuleType
MODUL_NORMAL, MODUL_DOKUMENT = 1, 4
PAKETORDNER = Path(__file__).resolve().parent.parent


def _programmordner() -> list:
    """Mögliche Programmordner von LibreOffice: aus dem PATH, sonst die Standardorte."""
    ordner = []
    for name in ("soffice", "libreoffice"):
        gefunden = shutil.which(name)
        if gefunden:
            ordner.append(Path(gefunden).resolve().parent)
    for variable in ("ProgramFiles", "ProgramFiles(x86)", "ProgramW6432"):
        if os.environ.get(variable):
            ordner.append(Path(os.environ[variable]) / "LibreOffice" / "program")
    ordner += [Path("/Applications/LibreOffice.app/Contents/MacOS"),
               Path("/usr/lib/libreoffice/program"), Path("/opt/libreoffice/program")]
    return ordner


def soffice_pfad() -> str:
    """Pfad zu soffice; unter Windows liegt es meist nicht im PATH."""
    gefunden = shutil.which("soffice") or shutil.which("libreoffice")
    if gefunden:
        return gefunden
    for ordner in _programmordner():
        for name in ("soffice.exe", "soffice"):
            if (ordner / name).is_file():
                return str(ordner / name)
    return None


def libreoffice_python() -> str:
    """Das mit LibreOffice gelieferte Python (Windows, macOS); dort ist uno importierbar."""
    for ordner in _programmordner():
        for kandidat in (ordner / "python.exe", ordner / "python",
                         ordner.parent / "Resources" / "python"):
            if kandidat.is_file():
                return str(kandidat)
    return None


def uno_verfuegbar() -> bool:
    try:
        import uno  # noqa: F401
    except ImportError:
        return False
    return True


def quelltext(datei: Path) -> str:
    """VBA-Quelltext ohne Attribute-Zeilen; LibreOffice setzt VB_Name beim Export selbst."""
    text = datei.read_text(encoding="ascii")  # bewusst ohne Umlaute, siehe modStart
    zeilen = [z for z in text.splitlines() if not z.startswith("Attribute ")]
    return "\n".join(zeilen) + "\n"


def _pv(name: str, wert):
    from com.sun.star.beans import PropertyValue
    p = PropertyValue()
    p.Name, p.Value = name, wert
    return p


class LibreOffice:
    """LibreOffice headless mit UNO-Verbindung über eine benannte Pipe."""

    def __init__(self, arbeitsordner: Path):
        self.ordner = Path(arbeitsordner)
        self.pipe = f"prognosemodell_{uuid.uuid4().hex}"
        self.proc = None
        self.desktop = None

    def __enter__(self):
        import uno
        soffice = soffice_pfad()
        if not soffice:
            raise RuntimeError("LibreOffice (soffice) nicht gefunden; für die Makros nötig.")
        self.ordner.mkdir(parents=True, exist_ok=True)
        self.proc = subprocess.Popen(
            [soffice, f"-env:UserInstallation={(self.ordner / 'lo-profil').as_uri()}",
             "--headless", "--invisible", "--norestore",
             f"--accept=pipe,name={self.pipe};urp;"],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        lokal = uno.getComponentContext()
        resolver = lokal.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", lokal)
        for _ in range(120):
            try:
                ctx = resolver.resolve(
                    f"uno:pipe,name={self.pipe};urp;StarOffice.ComponentContext")
                break
            except Exception:
                if self.proc.poll() is not None:
                    raise RuntimeError("LibreOffice ist beim Start beendet worden.")
                time.sleep(0.5)
        else:
            raise RuntimeError("Keine Verbindung zu LibreOffice.")
        self.desktop = ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", ctx)
        return self

    def __exit__(self, *_):
        try:
            self.desktop.terminate()
        except Exception:
            pass
        try:
            self.proc.wait(timeout=30)
        except subprocess.TimeoutExpired:
            self.proc.kill()

    def laden(self, pfad: Path, makros: bool = False):
        # MacroExecutionMode 4: Makros ohne Rückfrage ausführen (nur im Prüfskript)
        args = [_pv("Hidden", True)]
        if makros:
            args.append(_pv("MacroExecutionMode", 4))
        return self.desktop.loadComponentFromURL(Path(pfad).resolve().as_uri(), "_blank", 0,
                                                 tuple(args))

    def neu(self):
        return self.desktop.loadComponentFromURL("private:factory/scalc", "_blank", 0,
                                                 (_pv("Hidden", True),))

    @staticmethod
    def speichern(doc, pfad: Path, filter_: str = FILTER_XLSX) -> None:
        doc.storeToURL(Path(pfad).resolve().as_uri(), (_pv("FilterName", filter_),))

    @staticmethod
    def makro(doc, modul: str, prozedur: str, *argumente):
        """Makro im VBA-Projekt der Mappe aufrufen; liefert den Rückgabewert."""
        skript = doc.getScriptProvider().getScript(
            f"vnd.sun.star.script:{PROJEKTNAME}.{modul}.{prozedur}"
            "?language=Basic&location=document")
        ergebnis, _, _ = skript.invoke(tuple(argumente), (), ())
        return ergebnis


def _vba_hilfsmappe(lo: LibreOffice, blaetter: list, ziel: Path) -> None:
    """Leere Mappe mit den Codenamen der echten Mappe und allen Modulen als .xlsm.

    blaetter: (Blattname, Codename) in der Reihenfolge der Mappe.
    """
    from com.sun.star.script import ModuleInfo
    doc = lo.neu()
    try:
        sheets = doc.Sheets
        while sheets.Count < len(blaetter):
            sheets.insertNewByName(f"Blatt{sheets.Count + 1}", sheets.Count)
        for i, (titel, codename) in enumerate(blaetter):
            blatt = sheets.getByIndex(i)
            blatt.Name = titel
            blatt.CodeName = codename
        doc.CodeName = CODENAME_MAPPE

        bibliotheken = doc.BasicLibraries
        bibliotheken.VBACompatibilityMode = True
        bibliotheken.ProjectName = PROJEKTNAME
        bibliothek = bibliotheken.createLibrary(PROJEKTNAME)
        # liefert das Workbook-Objekt; nur damit exportiert LibreOffice das Modul
        # ThisWorkbook als Arbeitsmappen-Modul statt als Tabellenmodul
        objekte = doc.createInstance("ooo.vba.VBAObjectModuleObjectProvider")

        def modul(name: str, code: str, typ: int, objekt=None) -> None:
            info = ModuleInfo()
            info.ModuleType = typ
            if objekt is not None:
                info.ModuleObject = objekt
            bibliothek.insertModuleInfo(name, info)
            bibliothek.insertByName(name, "Option VBASupport 1\n" + code)

        modul(CODENAME_MAPPE, quelltext(VBA_ORDNER / "ThisWorkbook.cls"), MODUL_DOKUMENT,
              objekte.getByName(CODENAME_MAPPE))
        for _, codename in blaetter:
            modul(codename, "", MODUL_DOKUMENT)
        for name in MODULE:
            modul(name, quelltext(VBA_ORDNER / f"{name}.bas"), MODUL_NORMAL)
        lo.speichern(doc, ziel, FILTER_XLSM)
    finally:
        doc.close(True)


def _hilfsmappe_extern(blaetter: list, ziel: Path) -> None:
    """Hilfsmappe mit dem Python von LibreOffice bauen, wenn hier uno fehlt."""
    python = libreoffice_python()
    if not python:
        raise RuntimeError(
            "Für --makros wird LibreOffice mit der Python-UNO-Brücke gebraucht (Modul uno). "
            "Unter Windows: LibreOffice installieren (Standardordner C:\\Program Files"
            "\\LibreOffice), das mitgelieferte Python wird dann automatisch verwendet. "
            "Unter Debian/Ubuntu: Pakete libreoffice-calc und python3-uno. "
            "Ohne --makros entsteht die Mappe als .xlsx ohne Schaltflächen.")
    blattliste = Path(ziel).with_suffix(".json")
    blattliste.write_text(json.dumps(blaetter), encoding="utf-8")
    # sys.path statt PYTHONPATH: der Starter von LibreOffice setzt PYTHONPATH selbst
    code = (f"import sys; sys.path.insert(0, {str(PAKETORDNER)!r}); "
            "from prognosemodell.makros import _hilfsmappe_kommandozeile; "
            "_hilfsmappe_kommandozeile(sys.argv[1:])")
    lauf = subprocess.run([python, "-c", code, str(blattliste), str(ziel)],
                          capture_output=True, text=True, errors="replace")
    if lauf.returncode != 0 or not Path(ziel).is_file():
        raise RuntimeError(f"LibreOffice-Python ({python}) konnte das VBA-Projekt nicht "
                           f"bauen:\n{lauf.stderr.strip() or lauf.stdout.strip()}")


def _hilfsmappe_kommandozeile(argumente: list) -> None:
    """Einstieg für das LibreOffice-Python: Blattliste (JSON) und Ziel der Hilfsmappe."""
    blattliste, ziel = Path(argumente[0]), Path(argumente[1])
    blaetter = [tuple(b) for b in json.loads(blattliste.read_text(encoding="utf-8"))]
    with LibreOffice(ziel.parent / "lo") as lo:
        _vba_hilfsmappe(lo, blaetter, ziel)


def speichere_mit_makros(wb, ziel: Path, lo: LibreOffice = None) -> Path:
    """Mappe als .xlsm mit eingebettetem VBA-Projekt speichern."""
    ziel = Path(ziel)
    blaetter = [(ws.title, ws.sheet_properties.codeName) for ws in wb.worksheets]
    with tempfile.TemporaryDirectory() as tmp:
        hilfsmappe = Path(tmp) / "vba.xlsm"
        if lo is None and not uno_verfuegbar():
            _hilfsmappe_extern(blaetter, hilfsmappe)
        elif lo is None:
            with LibreOffice(Path(tmp) / "lo") as eigenes:
                _vba_hilfsmappe(eigenes, blaetter, hilfsmappe)
        else:
            _vba_hilfsmappe(lo, blaetter, hilfsmappe)
        with zipfile.ZipFile(hilfsmappe) as archiv:
            if "xl/vbaProject.bin" not in archiv.namelist():
                raise RuntimeError("LibreOffice hat kein VBA-Projekt geschrieben.")
            wb.vba_archive = archiv
            try:
                wb.save(ziel)
            finally:
                wb.vba_archive = None
    return ziel
