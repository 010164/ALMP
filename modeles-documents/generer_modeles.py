# -*- coding: utf-8 -*-
"""Génère les modèles de documents Word pour SARL ALM PATRIMOINE CONSEILS."""

import os

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

DOSSIER = os.path.dirname(os.path.abspath(__file__))
LOGO_PATH = os.path.join(DOSSIER, "logo-ALM.png")

ENTETE_TITRE = "SARL ALM PATRIMOINE CONSEILS"
ENTETE_LIGNE = (
    "au capital de 750€ – SIREN : 899 178 321 – RCS Pau – Enregistrée à l’Orias "
    "sous le N° 21004946 : en qualité de CIF, COA, MIA, MIOBSP – www.orias.fr"
)
MENTIONS = (
    "Conseiller en Investissements Financiers, adhérent de la CNCIF (association agréée "
    "par l’Autorité des Marchés Financiers) – Courtier en Assurances et Mandataire "
    "d’intermédiaire d’assurance, sous le contrôle de l’ACPR, 4 place de Budapest, "
    "CS 92459 75 436 Paris cedex 09 – Mandataire d’Intermédiaire en Opérations de Banques "
    "et en Services de Paiement – Transaction en immeubles et fonds de commerce sans "
    "manipulation de fonds, Carte Professionnelle Immobilière n° 6402 2021 000 000 006 "
    "délivrée par la CCI Pau Béarn – Garanties financières et Responsabilité Civile "
    "Professionnelle : MMA IARD Assurances Mutuelles n° 127 128 662 – conforme au code "
    "des assurances"
)
MEDIATION = (
    "Conformément aux articles L.616-1 et R.616-1 du code de la consommation, notre "
    "société a mis en place un dispositif de médiation de la consommation. L’entité de "
    "médiation retenue est : SAS CNPM - MÉDIATION - CONSOMMATION. En cas de litige, le "
    "consommateur pourra déposer sa réclamation sur le site : "
    "http://cnpm-mediation-consommation.eu ou par voie postale en écrivant à "
    "CNPM MEDIATION CONSOMMATION, Centre d’Affaires Stéphanois SAS, IMMEUBLE L’HORIZON – "
    "ESPLANADE DE France, 3, RUE J. CONSTANT MILLERET – 42000 SAINT-ÉTIENNE"
)

GRIS = RGBColor(0x59, 0x59, 0x59)
BLEU = RGBColor(0x1F, 0x38, 0x64)


def bordure(paragraph, cote, sz=8, couleur="9CA3AF"):
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = OxmlElement("w:pBdr")
    elem = OxmlElement("w:" + cote)
    elem.set(qn("w:val"), "single")
    elem.set(qn("w:sz"), str(sz))
    elem.set(qn("w:space"), "4")
    elem.set(qn("w:color"), couleur)
    pBdr.append(elem)
    pPr.append(pBdr)


def configurer_page(doc):
    s = doc.sections[0]
    s.page_width = Cm(21.0)
    s.page_height = Cm(29.7)
    s.top_margin = Cm(3.2)
    s.bottom_margin = Cm(3.9)
    s.left_margin = Cm(2.0)
    s.right_margin = Cm(2.0)
    s.header_distance = Cm(0.8)
    s.footer_distance = Cm(0.7)
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(11)
    return s


def ajouter_entete(section):
    h = section.header
    h.is_linked_to_previous = False
    p = h.paragraphs[0]
    p.text = ""
    if os.path.exists(LOGO_PATH):
        p.add_run().add_picture(LOGO_PATH, height=Cm(1.4))
        p.paragraph_format.tab_stops.add_tab_stop(Cm(2.3), WD_TAB_ALIGNMENT.LEFT)
        r = p.add_run("\t" + ENTETE_TITRE)
    else:
        r = p.add_run(ENTETE_TITRE)
    r.bold = True
    r.font.size = Pt(15)
    r.font.color.rgb = BLEU
    p2 = h.add_paragraph()
    r2 = p2.add_run(ENTETE_LIGNE)
    r2.font.size = Pt(7.5)
    r2.font.color.rgb = GRIS
    bordure(p2, "bottom")


def ajouter_pied(section):
    f = section.footer
    f.is_linked_to_previous = False
    p0 = f.paragraphs[0]
    p0.text = ""
    bordure(p0, "top")
    r0 = p0.add_run(MENTIONS)
    r0.font.size = Pt(6.5)
    r0.font.color.rgb = GRIS
    p1 = f.add_paragraph()
    r1 = p1.add_run(MEDIATION)
    r1.font.size = Pt(6.5)
    r1.font.color.rgb = GRIS


def nouveau_doc(titre):
    doc = Document()
    section = configurer_page(doc)
    ajouter_entete(section)
    ajouter_pied(section)
    doc.core_properties.title = titre
    doc.core_properties.author = ENTETE_TITRE
    return doc


def paragraphe(doc, texte, taille=11, gras=False, italique=False, couleur=None,
                alignement=None, espace_apres=6):
    p = doc.add_paragraph()
    r = p.add_run(texte)
    r.font.size = Pt(taille)
    r.bold = gras
    r.italic = italique
    if couleur is not None:
        r.font.color.rgb = couleur
    if alignement is not None:
        p.alignment = alignement
    p.paragraph_format.space_after = Pt(espace_apres)
    return p


def modele_papier_entete():
    doc = nouveau_doc("Papier à en-tête – SARL ALM PATRIMOINE CONSEILS")
    paragraphe(
        doc,
        "[Rédigez votre texte ici – cet en-tête et ce pied de page s’affichent "
        "automatiquement sur chaque page.]",
        taille=10, italique=True, couleur=GRIS,
    )
    doc.save(os.path.join(DOSSIER, "01-papier-a-entete.docx"))


def modele_courrier():
    doc = nouveau_doc("Modèle de courrier – SARL ALM PATRIMOINE CONSEILS")
    paragraphe(doc, "[Nom et prénom du destinataire]\n[Adresse]\n[Code postal] [Ville]")
    paragraphe(doc, "[Ville], le [date]", alignement=WD_ALIGN_PARAGRAPH.RIGHT)
    paragraphe(doc, "Objet : [Objet du courrier]", gras=True, espace_apres=12)
    paragraphe(
        doc,
        "Madame, Monsieur,\n\n[Corps du courrier : exposez votre demande ou votre "
        "information de manière claire et concise.]\n\n[Veuillez préciser les pièces "
        "à joindre et les délais applicables, le cas échéant.]",
        espace_apres=12,
    )
    paragraphe(
        doc,
        "Veuillez agréer, Madame, Monsieur, l’expression de nos salutations distinguées.",
        espace_apres=24,
    )
    paragraphe(doc, "[Prénom NOM]\n[Qualité]\n" + ENTETE_TITRE, gras=False)
    doc.save(os.path.join(DOSSIER, "02-modele-courrier.docx"))


def modele_devis():
    doc = nouveau_doc("Modèle de devis – SARL ALM PATRIMOINE CONSEILS")
    paragraphe(doc, "DEVIS", taille=18, gras=True, couleur=BLEU,
               alignement=WD_ALIGN_PARAGRAPH.CENTER, espace_apres=2)
    paragraphe(doc, "N° [AAAA-NNN]", taille=10, alignement=WD_ALIGN_PARAGRAPH.CENTER,
               espace_apres=2)
    paragraphe(doc, "[Ville], le [date]", taille=10,
               alignement=WD_ALIGN_PARAGRAPH.CENTER, espace_apres=14)

    paragraphe(doc, "Client :\n[Nom / Société du client]\n[Adresse]\n[Code postal] [Ville]",
               espace_apres=14)
    paragraphe(doc, "Objet : [Objet de la prestation]", gras=True, espace_apres=10)

    tableau = doc.add_table(rows=5, cols=4)
    tableau.style = "Table Grid"
    tableau.alignment = WD_TABLE_ALIGNMENT.CENTER
    en_tetes = ["Désignation", "Qté", "Prix unitaire HT", "Montant HT"]
    for i, texte in enumerate(en_tetes):
        cellule = tableau.rows[0].cells[i]
        cellule.text = ""
        r = cellule.paragraphs[0].add_run(texte)
        r.bold = True
        r.font.size = Pt(10)
    for ligne in range(1, 4):
        for col in range(4):
            cellule = tableau.rows[ligne].cells[col]
            cellule.text = ""
            r = cellule.paragraphs[0].add_run("[…]")
            r.font.size = Pt(10)
            r.font.color.rgb = GRIS
    libelles = ["Total HT", "TVA (taux : […] %)", "Total TTC"]
    for libelle in libelles:
        ligne = tableau.add_row()
        c0 = ligne.cells[0]
        c0.merge(ligne.cells[1]).merge(ligne.cells[2])
        c0.text = ""
        r = c0.paragraphs[0].add_run(libelle)
        r.bold = (libelle == "Total TTC")
        r.font.size = Pt(10)
        c3 = ligne.cells[3]
        c3.text = ""
        r3 = c3.paragraphs[0].add_run("[…] €")
        r3.bold = (libelle == "Total TTC")
        r3.font.size = Pt(10)

    paragraphe(doc, "", espace_apres=6)
    paragraphe(doc, "Validité du devis : [30 jours] à compter de la date d’émission.",
               taille=10, espace_apres=4)
    paragraphe(doc, "Modalités de règlement : [acompte de … % à la commande, solde à "
               "l’issue de la prestation].", taille=10, espace_apres=4)
    paragraphe(doc, "Prestation réalisée conformément aux conditions générales remises "
               "au client.", taille=10, espace_apres=16)

    paragraphe(doc, "Bon pour accord", gras=True, espace_apres=18)
    paragraphe(doc, "Date : [………]          Signature du client :", taille=10, espace_apres=24)
    paragraphe(doc, "[Prénom NOM] – " + ENTETE_TITRE, taille=10)
    doc.save(os.path.join(DOSSIER, "03-modele-devis.docx"))


if __name__ == "__main__":
    modele_papier_entete()
    modele_courrier()
    modele_devis()
    print("Documents générés : 01-papier-a-entete.docx, 02-modele-courrier.docx, "
          "03-modele-devis.docx")
