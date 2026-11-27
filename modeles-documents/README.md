# Modèles de documents – SARL ALM PATRIMOINE CONSEILS

Ensemble de documents Word (.docx, format A4) avec en-tête et mentions légales intégrés
automatiquement sur chaque page.

## Fichiers

| Fichier | Usage |
|---|---|
| `01-papier-a-entete.docx` | Papier à en-tête maître : dupliquez ce fichier pour tout nouveau document (l'en-tête et le pied de page suivent). |
| `02-modele-courrier.docx` | Modèle de courrier avec champs à remplacer : `[destinataire]`, `[date]`, `[objet]`, corps, signature. |
| `03-modele-devis.docx` | Modèle de devis : tableau de prestations (désignation, quantité, PU HT, montant HT), totaux HT / TVA / TTC, conditions, « Bon pour accord ». |
| `blocs-de-texte.md` | Blocs de texte prêts à copier-coller (en-tête, mentions légales, médiation CNPM, signature e-mail) pour e-mails, PDF ou autres logiciels. |
| `generer_modeles.py` | Script qui régénère les trois fichiers .docx (Python 3 + `python-docx`). |

## Contenu intégré dans chaque document

- **En-tête** : SARL ALM PATRIMOINE CONSEILS – au capital de 750€ – SIREN : 899 178 321 – RCS Pau – Orias n° 21004946 (CIF, COA, MIA, MIOBSP) – www.orias.fr
- **Pied de page** : statut CIF / adhésion CNCIF, courtier en assurances (ACPR), MIOBSP, carte professionnelle immobilière CCI Pau Béarn n° 6402 2021 000 000 006, garanties financières et RCP MMA IARD n° 127 128 662, et médiation de la consommation CNPM (articles L.616-1 et R.616-1 du code de la consommation).

## Personnalisation

- **Logo** : dans Word, double-cliquez sur l'en-tête puis `Insertion > Images` pour placer votre logo à côté du nom de la société.
- **Couleur** : le nom de la société est en bleu foncé (`#1F3864`) ; modifiez `BLEU` dans `generer_modeles.py` pour changer la couleur de marque, puis relancez `python3 generer_modeles.py`.
- **Champs à remplacer** : tous les champs entre crochets `[...]` sont à personnaliser dans chaque document.
