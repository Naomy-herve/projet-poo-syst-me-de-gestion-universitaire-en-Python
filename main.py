"""
==============================================================
   SYSTEME DE GESTION UNIVERSITAIRE
   Application de bureau réalisée avec Tkinter, Pandas, Openpyxl & Reportlab
==============================================================

Fonctionnalités :
  - Onglet "Inscription" : saisie des informations de l'ETUDIANT
    (n°, nom, prénom, classe, adresse, contact) et du PARENT
    (nom, prénom, téléphone, adresse, message), puis enregistrement
    dans un fichier Excel.
  - Onglet "Notes" : vérification de l'étudiant (par n° + nom) puis
    saisie et enregistrement de ses notes dans 6 matières
    (Mathématique, Physique, Anglais, Français, Histoire, Géographie).
  - Onglet "Comptabilité" : vérification de l'étudiant, enregistrement
    des paiements (frais de scolarité) et génération d'un reçu (PDF).
  - Onglet "Tableau de Bord" : statistiques générales (nombre
    d'étudiants, répartition par classe, total des recettes) et
    génération de la liste des étudiants par classe (PDF).
  - Bouton "Ouvrir le fichier Excel" : ouvre le classeur Excel avec
    l'application associée du système d'exploitation.
  - Bouton "Quitter le programme" : ferme l'application après
    confirmation.

Bibliothèques utilisées : tkinter, pandas, openpyxl, reportlab, os, sys,
subprocess, datetime
==============================================================
"""

import os
import sys
import subprocess
from datetime import date
import tkinter as tk
from tkinter import ttk, messagebox

import pandas as pd

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet


# --------------------------------------------------------------------------
#   CONFIGURATION GENERALE
# --------------------------------------------------------------------------

DOSSIER_APPLICATION = os.path.dirname(os.path.abspath(__file__))
FICHIER_EXCEL = os.path.join(DOSSIER_APPLICATION, "gestion_universitaire.xlsx")

DOSSIER_DOCUMENTS = os.path.join(DOSSIER_APPLICATION, "documents_generes")
os.makedirs(DOSSIER_DOCUMENTS, exist_ok=True)

FEUILLE_ETUDIANTS = "Etudiants"
FEUILLE_NOTES = "Notes"
FEUILLE_COMPTABILITE = "Comptabilite"

COLONNES_ETUDIANTS = [
    "N°", "Nom Étudiant", "Prénom Étudiant", "Classe",
    "Adresse Étudiant", "Contact Étudiant",
    "Nom Parent", "Prénom Parent", "Téléphone Parent",
    "Adresse Parent", "Message"
]

MATIERES = ["Mathématique", "Physique", "Anglais", "Français", "Histoire", "Géographie"]

COLONNES_NOTES = ["N°", "Nom Étudiant", "Prénom Étudiant"] + MATIERES

COLONNES_COMPTABILITE = [
    "N°", "Nom Étudiant", "Prénom Étudiant", "Date",
    "Montant Payé (Ar)", "Mode de Paiement", "Observation"
]

# Classes disponibles (repérées sur le tableau : 7ème, 1ère, Tle) + niveaux intermédiaires usuels
CLASSES_DISPONIBLES = ["7ème", "6ème", "5ème", "4ème", "3ème", "2nde", "1ère", "Tle"]
CAPACITE_MAX_PAR_CLASSE = 25  # Capacité maximale par classe (indiquée sur le tableau : "25")

MODES_PAIEMENT = ["Espèces", "Chèque", "Virement bancaire", "Mobile Money"]


# --------------------------------------------------------------------------
#   GESTION DU FICHIER EXCEL
# --------------------------------------------------------------------------

def initialiser_fichier_excel():
    """Crée le fichier Excel avec les trois feuilles s'il n'existe pas déjà."""
    if not os.path.exists(FICHIER_EXCEL):
        df_etudiants = pd.DataFrame(columns=COLONNES_ETUDIANTS)
        df_notes = pd.DataFrame(columns=COLONNES_NOTES)
        df_compta = pd.DataFrame(columns=COLONNES_COMPTABILITE)
        with pd.ExcelWriter(FICHIER_EXCEL, engine="openpyxl") as writer:
            df_etudiants.to_excel(writer, sheet_name=FEUILLE_ETUDIANTS, index=False)
            df_notes.to_excel(writer, sheet_name=FEUILLE_NOTES, index=False)
            df_compta.to_excel(writer, sheet_name=FEUILLE_COMPTABILITE, index=False)


def lire_feuille(nom_feuille, colonnes):
    """Lit une feuille du classeur Excel. Retourne un DataFrame vide en cas de problème."""
    try:
        return pd.read_excel(FICHIER_EXCEL, sheet_name=nom_feuille, engine="openpyxl")
    except Exception:
        return pd.DataFrame(columns=colonnes)


def ecrire_toutes_les_feuilles(df_etudiants, df_notes, df_compta):
    """Réécrit le classeur Excel complet (les trois feuilles) avec Pandas + Openpyxl."""
    with pd.ExcelWriter(FICHIER_EXCEL, engine="openpyxl") as writer:
        df_etudiants.to_excel(writer, sheet_name=FEUILLE_ETUDIANTS, index=False)
        df_notes.to_excel(writer, sheet_name=FEUILLE_NOTES, index=False)
        df_compta.to_excel(writer, sheet_name=FEUILLE_COMPTABILITE, index=False)


def ouvrir_fichier(chemin):
    """Ouvre un fichier (Excel, PDF, ...) avec l'application par défaut du système."""
    try:
        if sys.platform.startswith("win"):
            os.startfile(chemin)  # Windows
        elif sys.platform == "darwin":
            subprocess.call(["open", chemin])  # macOS
        else:
            subprocess.call(["xdg-open", chemin])  # Linux
    except Exception as erreur:
        messagebox.showerror("Erreur", f"Impossible d'ouvrir le fichier :\n{erreur}")


def ouvrir_fichier_excel():
    """Ouvre le fichier Excel avec l'application par défaut du système."""
    if not os.path.exists(FICHIER_EXCEL):
        initialiser_fichier_excel()
    ouvrir_fichier(FICHIER_EXCEL)


# --------------------------------------------------------------------------
#   GENERATION DE DOCUMENTS PDF (Reportlab)
# --------------------------------------------------------------------------

def generer_liste_classe_pdf(classe, df_etudiants):
    """Génère un PDF listant tous les étudiants d'une classe donnée."""
    df_classe = df_etudiants[df_etudiants["Classe"].astype(str) == str(classe)]

    chemin_pdf = os.path.join(DOSSIER_DOCUMENTS, f"Liste_classe_{classe}.pdf")
    styles = getSampleStyleSheet()
    document = SimpleDocTemplate(chemin_pdf, pagesize=A4,
                                  topMargin=2 * cm, bottomMargin=2 * cm)
    elements = []

    elements.append(Paragraph("SYSTEME DE GESTION UNIVERSITAIRE", styles["Title"]))
    elements.append(Paragraph(f"Liste des étudiants — Classe {classe}", styles["Heading2"]))
    elements.append(Paragraph(f"Date d'édition : {date.today().strftime('%d/%m/%Y')}", styles["Normal"]))
    elements.append(Paragraph(f"Effectif : {len(df_classe)} / {CAPACITE_MAX_PAR_CLASSE}", styles["Normal"]))
    elements.append(Spacer(1, 0.6 * cm))

    donnees = [["N°", "Nom", "Prénom", "Contact", "Adresse"]]
    for _, ligne in df_classe.iterrows():
        donnees.append([
            str(ligne.get("N°", "")),
            str(ligne.get("Nom Étudiant", "")),
            str(ligne.get("Prénom Étudiant", "")),
            str(ligne.get("Contact Étudiant", "")),
            str(ligne.get("Adresse Étudiant", "")),
        ])

    if len(donnees) == 1:
        elements.append(Paragraph("Aucun étudiant enregistré dans cette classe.", styles["Normal"]))
    else:
        tableau = Table(donnees, repeatRows=1, colWidths=[2 * cm, 4 * cm, 4 * cm, 3.5 * cm, 4 * cm])
        tableau.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1d3557")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 9),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#eef2f7")]),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ]))
        elements.append(tableau)

    document.build(elements)
    return chemin_pdf


def generer_recu_pdf(paiement):
    """Génère un reçu de paiement (PDF) à partir d'un dictionnaire de données."""
    numero_recu = f"{paiement['N°']}_{paiement['Date'].replace('/', '-')}"
    chemin_pdf = os.path.join(DOSSIER_DOCUMENTS, f"Recu_{numero_recu}.pdf")

    styles = getSampleStyleSheet()
    document = SimpleDocTemplate(chemin_pdf, pagesize=A4,
                                  topMargin=2 * cm, bottomMargin=2 * cm)
    elements = []

    elements.append(Paragraph("SYSTEME DE GESTION UNIVERSITAIRE", styles["Title"]))
    elements.append(Paragraph("REÇU DE PAIEMENT", styles["Heading2"]))
    elements.append(Spacer(1, 0.5 * cm))

    donnees = [
        ["Date", paiement["Date"]],
        ["N° Étudiant", paiement["N°"]],
        ["Nom & Prénom", f"{paiement['Nom Étudiant']} {paiement['Prénom Étudiant']}"],
        ["Montant payé", f"{paiement['Montant Payé (Ar)']} Ar"],
        ["Mode de paiement", paiement["Mode de Paiement"]],
        ["Observation", paiement.get("Observation", "")],
    ]
    tableau = Table(donnees, colWidths=[5 * cm, 9 * cm])
    tableau.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (0, -1), colors.HexColor("#eef2f7")),
        ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("FONTSIZE", (0, 0), (-1, -1), 10),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    elements.append(tableau)
    elements.append(Spacer(1, 1.2 * cm))
    elements.append(Paragraph("Signature :", styles["Normal"]))

    document.build(elements)
    return chemin_pdf


# --------------------------------------------------------------------------
#   APPLICATION PRINCIPALE
# --------------------------------------------------------------------------

class ApplicationGestionUniversitaire(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title("SYSTEME DE GESTION UNIVERSITAIRE")
        self.geometry("820x680")
        self.resizable(False, False)
        self.configure(bg="#eef2f7")

        initialiser_fichier_excel()

        # Etats de vérification (indépendants entre les onglets Notes et Comptabilité)
        self.etudiant_verifie_notes = None
        self.etudiant_verifie_compta = None
        self.dernier_paiement = None

        self._construire_entete()
        self._construire_onglets()
        self._construire_barre_boutons()

    # ------------------------------------------------------------------
    #   CONSTRUCTION DE L'INTERFACE GENERALE
    # ------------------------------------------------------------------

    def _construire_entete(self):
        entete = tk.Frame(self, bg="#1d3557", height=70)
        entete.pack(fill="x", side="top")
        tk.Label(
            entete,
            text="SYSTEME DE GESTION UNIVERSITAIRE",
            bg="#1d3557",
            fg="white",
            font=("Segoe UI", 18, "bold"),
        ).pack(pady=15)

    def _construire_onglets(self):
        style = ttk.Style()
        style.configure("TNotebook.Tab", font=("Segoe UI", 10, "bold"), padding=[12, 8])

        self.onglets = ttk.Notebook(self)
        self.onglets.pack(fill="both", expand=True, padx=10, pady=10)

        self.onglet_inscription = tk.Frame(self.onglets, bg="#eef2f7")
        self.onglet_notes = tk.Frame(self.onglets, bg="#eef2f7")
        self.onglet_comptabilite = tk.Frame(self.onglets, bg="#eef2f7")
        self.onglet_tableau_bord = tk.Frame(self.onglets, bg="#eef2f7")

        self.onglets.add(self.onglet_inscription, text="  Inscription  ")
        self.onglets.add(self.onglet_notes, text="  Notes  ")
        self.onglets.add(self.onglet_comptabilite, text="  Comptabilité  ")
        self.onglets.add(self.onglet_tableau_bord, text="  Tableau de Bord  ")

        self._construire_onglet_inscription()
        self._construire_onglet_notes()
        self._construire_onglet_comptabilite()
        self._construire_onglet_tableau_bord()

        # Actualiser le tableau de bord dès l'ouverture, et à chaque changement d'onglet
        self.onglets.bind("<<NotebookTabChanged>>", lambda e: self._actualiser_tableau_bord())
        self._actualiser_tableau_bord()

    def _construire_barre_boutons(self):
        barre = tk.Frame(self, bg="#eef2f7")
        barre.pack(fill="x", padx=10, pady=(0, 12))

        bouton_excel = tk.Button(
            barre, text="📂  Ouvrir le fichier Excel", font=("Segoe UI", 10, "bold"),
            bg="#2a9d8f", fg="white", relief="flat", padx=12, pady=8,
            command=ouvrir_fichier_excel
        )
        bouton_excel.pack(side="left", padx=5)

        bouton_quitter = tk.Button(
            barre, text="⏻  Quitter le programme", font=("Segoe UI", 10, "bold"),
            bg="#e63946", fg="white", relief="flat", padx=12, pady=8,
            command=self.quitter_programme
        )
        bouton_quitter.pack(side="right", padx=5)

    def _champ(self, parent, libelle, variable, ligne, largeur=28):
        """Crée une paire (Label + Entry) dans une grille."""
        tk.Label(parent, text=libelle, bg="#eef2f7", font=("Segoe UI", 10)).grid(
            row=ligne, column=0, sticky="w", pady=6
        )
        tk.Entry(parent, textvariable=variable, width=largeur, font=("Segoe UI", 10)).grid(
            row=ligne, column=1, pady=6, padx=5
        )

    # ------------------------------------------------------------------
    #   ONGLET 1 : INSCRIPTION (ETUDIANT + PARENT)
    # ------------------------------------------------------------------

    def _construire_onglet_inscription(self):
        conteneur = tk.Frame(self.onglet_inscription, bg="#eef2f7")
        conteneur.pack(fill="both", expand=True, padx=10, pady=10)

        # ---------------- Cadre ETUDIANT ----------------
        cadre_etudiant = tk.LabelFrame(
            conteneur, text="  ETUDIANT  ", font=("Segoe UI", 11, "bold"),
            bg="#eef2f7", fg="#1d3557", padx=15, pady=15
        )
        cadre_etudiant.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        self.var_insc_numero = tk.StringVar()
        self.var_insc_nom = tk.StringVar()
        self.var_insc_prenom = tk.StringVar()
        self.var_insc_classe = tk.StringVar()
        self.var_insc_adresse = tk.StringVar()
        self.var_insc_contact = tk.StringVar()

        self._champ(cadre_etudiant, "N° Étudiant :", self.var_insc_numero, 0)
        self._champ(cadre_etudiant, "Nom :", self.var_insc_nom, 1)
        self._champ(cadre_etudiant, "Prénom :", self.var_insc_prenom, 2)

        tk.Label(cadre_etudiant, text="Classe :", bg="#eef2f7", font=("Segoe UI", 10)).grid(
            row=3, column=0, sticky="w", pady=6)
        combo_classe = ttk.Combobox(
            cadre_etudiant, textvariable=self.var_insc_classe,
            values=CLASSES_DISPONIBLES, width=25, font=("Segoe UI", 10)
        )
        combo_classe.grid(row=3, column=1, pady=6, padx=5)

        self._champ(cadre_etudiant, "Adresse :", self.var_insc_adresse, 4)
        self._champ(cadre_etudiant, "Contact :", self.var_insc_contact, 5)

        # ---------------- Cadre PARENT ----------------
        cadre_parent = tk.LabelFrame(
            conteneur, text="  PARENT / TUTEUR  ", font=("Segoe UI", 11, "bold"),
            bg="#eef2f7", fg="#1d3557", padx=15, pady=15
        )
        cadre_parent.grid(row=0, column=1, sticky="nsew", padx=(10, 0))

        self.var_parent_nom = tk.StringVar()
        self.var_parent_prenom = tk.StringVar()
        self.var_parent_tel = tk.StringVar()
        self.var_parent_adresse = tk.StringVar()

        self._champ(cadre_parent, "Nom :", self.var_parent_nom, 0)
        self._champ(cadre_parent, "Prénom :", self.var_parent_prenom, 1)
        self._champ(cadre_parent, "Téléphone (Contact Parents) :", self.var_parent_tel, 2)
        self._champ(cadre_parent, "Adresse :", self.var_parent_adresse, 3)

        tk.Label(cadre_parent, text="Message :", bg="#eef2f7",
                 font=("Segoe UI", 10)).grid(row=4, column=0, sticky="nw", pady=6)
        self.texte_message_parent = tk.Text(cadre_parent, width=26, height=4, font=("Segoe UI", 10))
        self.texte_message_parent.grid(row=4, column=1, pady=6, sticky="w")

        conteneur.grid_columnconfigure(0, weight=1)
        conteneur.grid_columnconfigure(1, weight=1)

        # ---------------- Bouton Enregistrer ----------------
        bouton_valider = tk.Button(
            self.onglet_inscription, text="✔  Valider l'inscription (Enregistrer vers Excel)",
            font=("Segoe UI", 11, "bold"), bg="#457b9d", fg="white",
            relief="flat", padx=15, pady=10,
            command=self.enregistrer_inscription
        )
        bouton_valider.pack(pady=15)

    def enregistrer_inscription(self):
        """Enregistre l'étudiant + le parent dans la feuille 'Etudiants' du fichier Excel."""
        numero = self.var_insc_numero.get().strip()
        nom = self.var_insc_nom.get().strip()
        prenom = self.var_insc_prenom.get().strip()
        classe = self.var_insc_classe.get().strip()

        # --- Vérification des champs obligatoires ---
        if not numero or not nom or not prenom or not classe:
            messagebox.showwarning(
                "Champs incomplets",
                "Veuillez remplir au minimum le N°, le Nom, le Prénom et la Classe de l'étudiant."
            )
            return

        try:
            df_etudiants = lire_feuille(FEUILLE_ETUDIANTS, COLONNES_ETUDIANTS)
            df_notes = lire_feuille(FEUILLE_NOTES, COLONNES_NOTES)
            df_compta = lire_feuille(FEUILLE_COMPTABILITE, COLONNES_COMPTABILITE)

            # Empêcher les doublons de N°
            if not df_etudiants.empty and (df_etudiants["N°"].astype(str) == numero).any():
                messagebox.showwarning("Doublon", f"Un étudiant avec le N° '{numero}' existe déjà.")
                return

            # Vérification de la capacité maximale de la classe (25 par classe)
            if not df_etudiants.empty:
                effectif_classe = (df_etudiants["Classe"].astype(str) == classe).sum()
            else:
                effectif_classe = 0

            if effectif_classe >= CAPACITE_MAX_PAR_CLASSE:
                continuer = messagebox.askyesno(
                    "Capacité atteinte",
                    f"La classe '{classe}' a déjà atteint sa capacité maximale "
                    f"({CAPACITE_MAX_PAR_CLASSE} étudiants).\nVoulez-vous tout de même inscrire cet étudiant ?"
                )
                if not continuer:
                    return

            nouvelle_ligne = {
                "N°": numero,
                "Nom Étudiant": nom,
                "Prénom Étudiant": prenom,
                "Classe": classe,
                "Adresse Étudiant": self.var_insc_adresse.get().strip(),
                "Contact Étudiant": self.var_insc_contact.get().strip(),
                "Nom Parent": self.var_parent_nom.get().strip(),
                "Prénom Parent": self.var_parent_prenom.get().strip(),
                "Téléphone Parent": self.var_parent_tel.get().strip(),
                "Adresse Parent": self.var_parent_adresse.get().strip(),
                "Message": self.texte_message_parent.get("1.0", "end").strip(),
            }

            df_etudiants = pd.concat([df_etudiants, pd.DataFrame([nouvelle_ligne])], ignore_index=True)
            ecrire_toutes_les_feuilles(df_etudiants, df_notes, df_compta)

            messagebox.showinfo("Succès", f"L'étudiant {nom} {prenom} a été enregistré avec succès.")
            self._vider_formulaire_inscription()
            self._actualiser_tableau_bord()

        except Exception as erreur:
            messagebox.showerror("Erreur", f"Une erreur est survenue lors de l'enregistrement :\n{erreur}")

    def _vider_formulaire_inscription(self):
        for variable in (self.var_insc_numero, self.var_insc_nom, self.var_insc_prenom,
                          self.var_insc_classe, self.var_insc_adresse, self.var_insc_contact,
                          self.var_parent_nom, self.var_parent_prenom,
                          self.var_parent_tel, self.var_parent_adresse):
            variable.set("")
        self.texte_message_parent.delete("1.0", "end")

    # ------------------------------------------------------------------
    #   ONGLET 2 : NOTES (VERIFICATION + SAISIE DES NOTES)
    # ------------------------------------------------------------------

    def _construire_onglet_notes(self):
        conteneur = tk.Frame(self.onglet_notes, bg="#eef2f7")
        conteneur.pack(fill="both", expand=True, padx=10, pady=10)

        # ---------------- Cadre Vérification ----------------
        cadre_verif = tk.LabelFrame(
            conteneur, text="  VÉRIFICATION DE L'ÉTUDIANT  ", font=("Segoe UI", 11, "bold"),
            bg="#eef2f7", fg="#1d3557", padx=15, pady=15
        )
        cadre_verif.pack(fill="x", pady=(0, 10))

        self.var_verif_numero = tk.StringVar()
        self.var_verif_nom = tk.StringVar()

        tk.Label(cadre_verif, text="N° Étudiant :", bg="#eef2f7", font=("Segoe UI", 10)).grid(
            row=0, column=0, sticky="w", padx=5, pady=5)
        tk.Entry(cadre_verif, textvariable=self.var_verif_numero, width=20, font=("Segoe UI", 10)).grid(
            row=0, column=1, padx=5, pady=5)

        tk.Label(cadre_verif, text="Nom :", bg="#eef2f7", font=("Segoe UI", 10)).grid(
            row=0, column=2, sticky="w", padx=5, pady=5)
        tk.Entry(cadre_verif, textvariable=self.var_verif_nom, width=20, font=("Segoe UI", 10)).grid(
            row=0, column=3, padx=5, pady=5)

        tk.Button(
            cadre_verif, text="🔍  Vérifier", font=("Segoe UI", 10, "bold"),
            bg="#457b9d", fg="white", relief="flat", padx=12, pady=6,
            command=self.verifier_etudiant_notes
        ).grid(row=0, column=4, padx=10)

        self.label_statut_verif_notes = tk.Label(
            cadre_verif, text="Aucun étudiant vérifié.", bg="#eef2f7",
            font=("Segoe UI", 10, "italic"), fg="#e63946"
        )
        self.label_statut_verif_notes.grid(row=1, column=0, columnspan=5, sticky="w", pady=(8, 0))

        # ---------------- Cadre Notes ----------------
        self.cadre_notes = tk.LabelFrame(
            conteneur, text="  SAISIE DES NOTES  ", font=("Segoe UI", 11, "bold"),
            bg="#eef2f7", fg="#1d3557", padx=15, pady=15
        )
        self.cadre_notes.pack(fill="x", pady=10)

        self.variables_notes = {}
        for index, matiere in enumerate(MATIERES):
            ligne, colonne = divmod(index, 2)
            tk.Label(self.cadre_notes, text=f"{matiere} :", bg="#eef2f7",
                     font=("Segoe UI", 10)).grid(row=ligne, column=colonne * 2, sticky="w", padx=5, pady=8)
            variable = tk.StringVar()
            tk.Entry(self.cadre_notes, textvariable=variable, width=10, font=("Segoe UI", 10)).grid(
                row=ligne, column=colonne * 2 + 1, padx=5, pady=8, sticky="w")
            self.variables_notes[matiere] = variable

        # Désactivé tant qu'aucun étudiant n'est vérifié
        self._activer_champs(self.cadre_notes, False)

        tk.Button(
            self.onglet_notes, text="✔  Valider les notes (Enregistrer vers Excel)",
            font=("Segoe UI", 11, "bold"), bg="#2a9d8f", fg="white",
            relief="flat", padx=15, pady=10,
            command=self.enregistrer_notes
        ).pack(pady=15)

    def _activer_champs(self, cadre, actif):
        etat = "normal" if actif else "disabled"
        for enfant in cadre.winfo_children():
            if isinstance(enfant, (tk.Entry, ttk.Combobox)):
                enfant.config(state=etat)

    def verifier_etudiant_notes(self):
        """Vérifie que l'étudiant (N° + Nom) existe bien dans le fichier Excel (onglet Notes)."""
        etudiant = self._verifier_etudiant_generique(self.var_verif_numero.get(), self.var_verif_nom.get())
        if etudiant is None:
            self.etudiant_verifie_notes = None
            self._activer_champs(self.cadre_notes, False)
            self.label_statut_verif_notes.config(
                text="❌ Étudiant introuvable. Vérifiez le N° et le Nom.", fg="#e63946"
            )
        else:
            self.etudiant_verifie_notes = etudiant
            self._activer_champs(self.cadre_notes, True)
            self.label_statut_verif_notes.config(
                text=f"✔ Étudiant vérifié : {etudiant['Nom Étudiant']} {etudiant['Prénom Étudiant']} "
                     f"(Classe {etudiant['Classe']})",
                fg="#2a9d8f"
            )

    def _verifier_etudiant_generique(self, numero, nom):
        """Recherche un étudiant par N° + Nom dans la feuille Etudiants. Retourne un dict ou None."""
        numero = numero.strip()
        nom = nom.strip()

        if not numero or not nom:
            messagebox.showwarning("Champs requis", "Veuillez saisir le N° et le Nom de l'étudiant.")
            return None

        df_etudiants = lire_feuille(FEUILLE_ETUDIANTS, COLONNES_ETUDIANTS)
        if df_etudiants.empty:
            messagebox.showerror("Introuvable", "Aucun étudiant n'est encore enregistré.")
            return None

        resultat = df_etudiants[
            (df_etudiants["N°"].astype(str).str.strip() == numero) &
            (df_etudiants["Nom Étudiant"].astype(str).str.strip().str.lower() == nom.lower())
        ]

        if resultat.empty:
            messagebox.showerror("Introuvable", "Aucun étudiant ne correspond à ce N° et à ce Nom.")
            return None

        ligne = resultat.iloc[0]
        return {
            "N°": str(ligne["N°"]),
            "Nom Étudiant": ligne["Nom Étudiant"],
            "Prénom Étudiant": ligne["Prénom Étudiant"],
            "Classe": ligne["Classe"],
        }

    def enregistrer_notes(self):
        """Enregistre les notes de l'étudiant vérifié dans la feuille 'Notes' du fichier Excel."""
        if not self.etudiant_verifie_notes:
            messagebox.showwarning(
                "Étudiant non vérifié",
                "Veuillez d'abord vérifier l'étudiant avant de saisir ses notes."
            )
            return

        notes_saisies = {}
        for matiere, variable in self.variables_notes.items():
            valeur = variable.get().strip()
            if valeur == "":
                messagebox.showwarning("Champ manquant", f"Veuillez saisir la note de {matiere}.")
                return
            try:
                note = float(valeur.replace(",", "."))
            except ValueError:
                messagebox.showerror("Valeur invalide", f"La note de {matiere} doit être un nombre.")
                return
            if not (0 <= note <= 20):
                messagebox.showerror("Valeur invalide", f"La note de {matiere} doit être comprise entre 0 et 20.")
                return
            notes_saisies[matiere] = note

        nouvelle_ligne = {
            "N°": self.etudiant_verifie_notes["N°"],
            "Nom Étudiant": self.etudiant_verifie_notes["Nom Étudiant"],
            "Prénom Étudiant": self.etudiant_verifie_notes["Prénom Étudiant"],
        }
        nouvelle_ligne.update(notes_saisies)

        try:
            df_etudiants = lire_feuille(FEUILLE_ETUDIANTS, COLONNES_ETUDIANTS)
            df_notes = lire_feuille(FEUILLE_NOTES, COLONNES_NOTES)
            df_compta = lire_feuille(FEUILLE_COMPTABILITE, COLONNES_COMPTABILITE)

            df_notes = pd.concat([df_notes, pd.DataFrame([nouvelle_ligne])], ignore_index=True)
            ecrire_toutes_les_feuilles(df_etudiants, df_notes, df_compta)

            messagebox.showinfo("Succès", "Les notes ont été enregistrées avec succès.")
            self._vider_formulaire_notes()

        except Exception as erreur:
            messagebox.showerror("Erreur", f"Une erreur est survenue lors de l'enregistrement :\n{erreur}")

    def _vider_formulaire_notes(self):
        for variable in self.variables_notes.values():
            variable.set("")
        self.var_verif_numero.set("")
        self.var_verif_nom.set("")
        self.etudiant_verifie_notes = None
        self._activer_champs(self.cadre_notes, False)
        self.label_statut_verif_notes.config(text="Aucun étudiant vérifié.", fg="#e63946")

    # ------------------------------------------------------------------
    #   ONGLET 3 : COMPTABILITE (PAIEMENTS + REÇU)
    # ------------------------------------------------------------------

    def _construire_onglet_comptabilite(self):
        conteneur = tk.Frame(self.onglet_comptabilite, bg="#eef2f7")
        conteneur.pack(fill="both", expand=True, padx=10, pady=10)

        # ---------------- Cadre Vérification ----------------
        cadre_verif = tk.LabelFrame(
            conteneur, text="  VÉRIFICATION DE L'ÉTUDIANT  ", font=("Segoe UI", 11, "bold"),
            bg="#eef2f7", fg="#1d3557", padx=15, pady=15
        )
        cadre_verif.pack(fill="x", pady=(0, 10))

        self.var_compta_numero = tk.StringVar()
        self.var_compta_nom = tk.StringVar()

        tk.Label(cadre_verif, text="N° Étudiant :", bg="#eef2f7", font=("Segoe UI", 10)).grid(
            row=0, column=0, sticky="w", padx=5, pady=5)
        tk.Entry(cadre_verif, textvariable=self.var_compta_numero, width=20, font=("Segoe UI", 10)).grid(
            row=0, column=1, padx=5, pady=5)

        tk.Label(cadre_verif, text="Nom :", bg="#eef2f7", font=("Segoe UI", 10)).grid(
            row=0, column=2, sticky="w", padx=5, pady=5)
        tk.Entry(cadre_verif, textvariable=self.var_compta_nom, width=20, font=("Segoe UI", 10)).grid(
            row=0, column=3, padx=5, pady=5)

        tk.Button(
            cadre_verif, text="🔍  Vérifier", font=("Segoe UI", 10, "bold"),
            bg="#457b9d", fg="white", relief="flat", padx=12, pady=6,
            command=self.verifier_etudiant_comptabilite
        ).grid(row=0, column=4, padx=10)

        self.label_statut_verif_compta = tk.Label(
            cadre_verif, text="Aucun étudiant vérifié.", bg="#eef2f7",
            font=("Segoe UI", 10, "italic"), fg="#e63946"
        )
        self.label_statut_verif_compta.grid(row=1, column=0, columnspan=5, sticky="w", pady=(8, 0))

        # ---------------- Cadre Paiement ----------------
        self.cadre_paiement = tk.LabelFrame(
            conteneur, text="  ENREGISTREMENT DU PAIEMENT  ", font=("Segoe UI", 11, "bold"),
            bg="#eef2f7", fg="#1d3557", padx=15, pady=15
        )
        self.cadre_paiement.pack(fill="x", pady=10)

        self.var_paiement_date = tk.StringVar(value=date.today().strftime("%d/%m/%Y"))
        self.var_paiement_montant = tk.StringVar()
        self.var_paiement_mode = tk.StringVar(value=MODES_PAIEMENT[0])
        self.var_paiement_observation = tk.StringVar()

        tk.Label(self.cadre_paiement, text="Date (JJ/MM/AAAA) :", bg="#eef2f7",
                 font=("Segoe UI", 10)).grid(row=0, column=0, sticky="w", padx=5, pady=8)
        tk.Entry(self.cadre_paiement, textvariable=self.var_paiement_date, width=18,
                 font=("Segoe UI", 10)).grid(row=0, column=1, padx=5, pady=8, sticky="w")

        tk.Label(self.cadre_paiement, text="Montant payé (Ar) :", bg="#eef2f7",
                 font=("Segoe UI", 10)).grid(row=0, column=2, sticky="w", padx=5, pady=8)
        tk.Entry(self.cadre_paiement, textvariable=self.var_paiement_montant, width=18,
                 font=("Segoe UI", 10)).grid(row=0, column=3, padx=5, pady=8, sticky="w")

        tk.Label(self.cadre_paiement, text="Mode de paiement :", bg="#eef2f7",
                 font=("Segoe UI", 10)).grid(row=1, column=0, sticky="w", padx=5, pady=8)
        ttk.Combobox(self.cadre_paiement, textvariable=self.var_paiement_mode,
                     values=MODES_PAIEMENT, width=16, font=("Segoe UI", 10)).grid(
            row=1, column=1, padx=5, pady=8, sticky="w")

        tk.Label(self.cadre_paiement, text="Observation :", bg="#eef2f7",
                 font=("Segoe UI", 10)).grid(row=1, column=2, sticky="w", padx=5, pady=8)
        tk.Entry(self.cadre_paiement, textvariable=self.var_paiement_observation, width=18,
                 font=("Segoe UI", 10)).grid(row=1, column=3, padx=5, pady=8, sticky="w")

        self._activer_champs(self.cadre_paiement, False)

        # ---------------- Boutons ----------------
        cadre_boutons = tk.Frame(self.onglet_comptabilite, bg="#eef2f7")
        cadre_boutons.pack(pady=15)

        tk.Button(
            cadre_boutons, text="✔  Enregistrer le paiement",
            font=("Segoe UI", 11, "bold"), bg="#2a9d8f", fg="white",
            relief="flat", padx=15, pady=10,
            command=self.enregistrer_paiement
        ).pack(side="left", padx=8)

        tk.Button(
            cadre_boutons, text="🧾  Générer le reçu (PDF)",
            font=("Segoe UI", 11, "bold"), bg="#f4a261", fg="white",
            relief="flat", padx=15, pady=10,
            command=self.generer_recu
        ).pack(side="left", padx=8)

    def verifier_etudiant_comptabilite(self):
        """Vérifie que l'étudiant (N° + Nom) existe bien dans le fichier Excel (onglet Comptabilité)."""
        etudiant = self._verifier_etudiant_generique(self.var_compta_numero.get(), self.var_compta_nom.get())
        if etudiant is None:
            self.etudiant_verifie_compta = None
            self._activer_champs(self.cadre_paiement, False)
            self.label_statut_verif_compta.config(
                text="❌ Étudiant introuvable. Vérifiez le N° et le Nom.", fg="#e63946"
            )
        else:
            self.etudiant_verifie_compta = etudiant
            self._activer_champs(self.cadre_paiement, True)
            self.label_statut_verif_compta.config(
                text=f"✔ Étudiant vérifié : {etudiant['Nom Étudiant']} {etudiant['Prénom Étudiant']} "
                     f"(Classe {etudiant['Classe']})",
                fg="#2a9d8f"
            )

    def enregistrer_paiement(self):
        """Enregistre un paiement dans la feuille 'Comptabilite' du fichier Excel."""
        if not self.etudiant_verifie_compta:
            messagebox.showwarning(
                "Étudiant non vérifié",
                "Veuillez d'abord vérifier l'étudiant avant d'enregistrer un paiement."
            )
            return

        montant = self.var_paiement_montant.get().strip()
        date_paiement = self.var_paiement_date.get().strip()

        if not montant or not date_paiement:
            messagebox.showwarning("Champs manquants", "Veuillez renseigner la date et le montant payé.")
            return

        try:
            montant_float = float(montant.replace(",", "."))
            if montant_float <= 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Valeur invalide", "Le montant payé doit être un nombre positif.")
            return

        nouvelle_ligne = {
            "N°": self.etudiant_verifie_compta["N°"],
            "Nom Étudiant": self.etudiant_verifie_compta["Nom Étudiant"],
            "Prénom Étudiant": self.etudiant_verifie_compta["Prénom Étudiant"],
            "Date": date_paiement,
            "Montant Payé (Ar)": montant_float,
            "Mode de Paiement": self.var_paiement_mode.get(),
            "Observation": self.var_paiement_observation.get().strip(),
        }

        try:
            df_etudiants = lire_feuille(FEUILLE_ETUDIANTS, COLONNES_ETUDIANTS)
            df_notes = lire_feuille(FEUILLE_NOTES, COLONNES_NOTES)
            df_compta = lire_feuille(FEUILLE_COMPTABILITE, COLONNES_COMPTABILITE)

            df_compta = pd.concat([df_compta, pd.DataFrame([nouvelle_ligne])], ignore_index=True)
            ecrire_toutes_les_feuilles(df_etudiants, df_notes, df_compta)

            self.dernier_paiement = nouvelle_ligne
            messagebox.showinfo("Succès", "Le paiement a été enregistré avec succès.")
            self._actualiser_tableau_bord()

        except Exception as erreur:
            messagebox.showerror("Erreur", f"Une erreur est survenue lors de l'enregistrement :\n{erreur}")

    def generer_recu(self):
        """Génère un reçu PDF pour le dernier paiement enregistré et l'ouvre."""
        if not self.dernier_paiement:
            messagebox.showwarning(
                "Aucun paiement",
                "Veuillez d'abord enregistrer un paiement avant de générer un reçu."
            )
            return
        try:
            paiement = dict(self.dernier_paiement)
            paiement["Montant Payé (Ar)"] = f"{paiement['Montant Payé (Ar)']:,.0f}".replace(",", " ")
            chemin_pdf = generer_recu_pdf(paiement)
            messagebox.showinfo("Reçu généré", f"Le reçu a été généré avec succès :\n{chemin_pdf}")
            ouvrir_fichier(chemin_pdf)
        except Exception as erreur:
            messagebox.showerror("Erreur", f"Impossible de générer le reçu :\n{erreur}")

    # ------------------------------------------------------------------
    #   ONGLET 4 : TABLEAU DE BORD (STATISTIQUES + LISTE PAR CLASSE)
    # ------------------------------------------------------------------

    def _construire_onglet_tableau_bord(self):
        conteneur = tk.Frame(self.onglet_tableau_bord, bg="#eef2f7")
        conteneur.pack(fill="both", expand=True, padx=10, pady=10)

        # ---------------- Cadre Statistiques ----------------
        cadre_stats = tk.LabelFrame(
            conteneur, text="  STATISTIQUES GÉNÉRALES  ", font=("Segoe UI", 11, "bold"),
            bg="#eef2f7", fg="#1d3557", padx=15, pady=15
        )
        cadre_stats.pack(fill="x", pady=(0, 10))

        self.label_total_etudiants = tk.Label(
            cadre_stats, text="Nombre total d'étudiants : 0", bg="#eef2f7", font=("Segoe UI", 11, "bold")
        )
        self.label_total_etudiants.grid(row=0, column=0, sticky="w", padx=5, pady=3)

        self.label_total_recettes = tk.Label(
            cadre_stats, text="Total des recettes : 0 Ar", bg="#eef2f7", font=("Segoe UI", 11, "bold")
        )
        self.label_total_recettes.grid(row=0, column=1, sticky="w", padx=20, pady=3)

        tk.Button(
            cadre_stats, text="🔄  Actualiser", font=("Segoe UI", 10, "bold"),
            bg="#457b9d", fg="white", relief="flat", padx=10, pady=5,
            command=self._actualiser_tableau_bord
        ).grid(row=0, column=2, padx=20)

        # Répartition par classe (Treeview)
        self.arbre_classes = ttk.Treeview(
            cadre_stats, columns=("classe", "effectif"), show="headings", height=6
        )
        self.arbre_classes.heading("classe", text="Classe")
        self.arbre_classes.heading("effectif", text=f"Effectif (max {CAPACITE_MAX_PAR_CLASSE})")
        self.arbre_classes.column("classe", width=200, anchor="center")
        self.arbre_classes.column("effectif", width=200, anchor="center")
        self.arbre_classes.grid(row=1, column=0, columnspan=3, pady=10, sticky="ew")

        # ---------------- Cadre Génération de liste par classe (PDF) ----------------
        cadre_liste = tk.LabelFrame(
            conteneur, text="  GÉNÉRATION DE LISTE PAR CLASSE (PDF)  ", font=("Segoe UI", 11, "bold"),
            bg="#eef2f7", fg="#1d3557", padx=15, pady=15
        )
        cadre_liste.pack(fill="x", pady=10)

        tk.Label(cadre_liste, text="Choisir une classe :", bg="#eef2f7", font=("Segoe UI", 10)).grid(
            row=0, column=0, sticky="w", padx=5, pady=5)

        self.var_classe_liste = tk.StringVar()
        self.combo_classe_liste = ttk.Combobox(
            cadre_liste, textvariable=self.var_classe_liste,
            values=CLASSES_DISPONIBLES, width=20, font=("Segoe UI", 10)
        )
        self.combo_classe_liste.grid(row=0, column=1, padx=5, pady=5)

        tk.Button(
            cadre_liste, text="📄  Générer la liste (PDF)", font=("Segoe UI", 10, "bold"),
            bg="#2a9d8f", fg="white", relief="flat", padx=12, pady=6,
            command=self.generer_liste_classe
        ).grid(row=0, column=2, padx=10)

    def _actualiser_tableau_bord(self):
        """Recalcule les statistiques à partir du fichier Excel."""
        df_etudiants = lire_feuille(FEUILLE_ETUDIANTS, COLONNES_ETUDIANTS)
        df_compta = lire_feuille(FEUILLE_COMPTABILITE, COLONNES_COMPTABILITE)

        self.label_total_etudiants.config(text=f"Nombre total d'étudiants : {len(df_etudiants)}")

        if not df_compta.empty and "Montant Payé (Ar)" in df_compta.columns:
            total_recettes = pd.to_numeric(df_compta["Montant Payé (Ar)"], errors="coerce").sum()
        else:
            total_recettes = 0
        self.label_total_recettes.config(text=f"Total des recettes : {total_recettes:,.0f} Ar".replace(",", " "))

        for ligne in self.arbre_classes.get_children():
            self.arbre_classes.delete(ligne)

        if not df_etudiants.empty:
            repartition = df_etudiants["Classe"].astype(str).value_counts()
            for classe in CLASSES_DISPONIBLES:
                effectif = int(repartition.get(classe, 0))
                self.arbre_classes.insert("", "end", values=(classe, effectif))
            # Classes non listées dans CLASSES_DISPONIBLES (saisies manuellement)
            for classe, effectif in repartition.items():
                if classe not in CLASSES_DISPONIBLES:
                    self.arbre_classes.insert("", "end", values=(classe, int(effectif)))
        else:
            for classe in CLASSES_DISPONIBLES:
                self.arbre_classes.insert("", "end", values=(classe, 0))

    def generer_liste_classe(self):
        """Génère la liste PDF des étudiants d'une classe choisie et l'ouvre."""
        classe = self.var_classe_liste.get().strip()
        if not classe:
            messagebox.showwarning("Classe manquante", "Veuillez choisir ou saisir une classe.")
            return
        try:
            df_etudiants = lire_feuille(FEUILLE_ETUDIANTS, COLONNES_ETUDIANTS)
            chemin_pdf = generer_liste_classe_pdf(classe, df_etudiants)
            messagebox.showinfo("Liste générée", f"La liste de la classe '{classe}' a été générée :\n{chemin_pdf}")
            ouvrir_fichier(chemin_pdf)
        except Exception as erreur:
            messagebox.showerror("Erreur", f"Impossible de générer la liste :\n{erreur}")

    # ------------------------------------------------------------------
    #   QUITTER LE PROGRAMME
    # ------------------------------------------------------------------

    def quitter_programme(self):
        confirmation = messagebox.askyesno(
            "Confirmer",
            "Êtes-vous sûr de vouloir quitter le programme ?"
        )
        if confirmation:
            self.destroy()


# --------------------------------------------------------------------------
#   POINT D'ENTREE
# --------------------------------------------------------------------------

if __name__ == "__main__":
    application = ApplicationGestionUniversitaire()
    application.mainloop()
