"""Testobjekt aus dem Projektplan (Abschnitt 7: Gebäude 800.000, 2 %, Kauf 2007)."""

from .modelle import Modell, Objekt


def testobjekt() -> Objekt:
    return Objekt(
        objekt_id="OBJ-001",
        name="Testobjekt",
        ak_gebaeude=800_000,
        ak_gub=200_000,
        kaufjahr=2007,
        afa_satz=0.02,
        # 20 volle Jahre AfA 2007 bis 2026 à 16.000
        restbuchwert=480_000,
        verkehrswert=1_400_000,
        vk_quote_gebaeude=0.5,
        miete=60_000,
        erhaltung=8_000,
        baujahr=2007,      # als Neubau gekauft
        san_jahr=0,        # keine Großmaßnahme geplant
        san_betrag=0,
    )


def testmodell() -> Modell:
    return Modell(objekte=[testobjekt()])
