import pdfplumber
import pandas as pd
import re
from pathlib import Path

# Fonction magique pour transformer le texte du PDF en vrai nombre mathématique
def texte_vers_nombre(valeur_texte):
    if not valeur_texte:
        return None # Laisse la case Excel vraiment vide
    
    # On enlève les espaces (séparateurs de milliers) et on met un point pour les décimales
    valeur_propre = valeur_texte.replace(" ", "").replace(",", ".")
    try:
        return float(valeur_propre) # Convertit en nombre à virgule (float)
    except ValueError:
        return valeur_texte # Sécurité : si ce n'est pas un nombre, on laisse le texte tel quel

# 1. Configuration des chemins dynamiques
dossier_courant = Path(__file__).parent
chemin_excel = dossier_courant / "export_paies_complet.xlsx"

print(f"Lancement de l'analyse dans le dossier : {dossier_courant}\n")

with pd.ExcelWriter(chemin_excel, engine='openpyxl') as writer:
    
    fichiers_pdf = list(dossier_courant.glob("*.pdf"))
    
    if not fichiers_pdf:
        print("Aucun fichier PDF trouvé dans ce dossier !")
    
    for chemin_pdf in fichiers_pdf:
        nom_fichier = chemin_pdf.stem
        print(f"Traitement de : {nom_fichier}...")
        
        lignes_donnees = []
        
        with pdfplumber.open(chemin_pdf) as pdf:
            for page in pdf.pages:
                mots = page.extract_words()
                if not mots:
                    continue
                    
                x1_payer = 430
                x1_deduire = 500
                x1_info = 570
                
                for mot in mots:
                    texte = mot['text'].upper()
                    if texte == "PAYER": x1_payer = mot['x1']
                    elif "DÉDUIRE" in texte or "DEDUIRE" in texte: x1_deduire = mot['x1']
                    elif "INFORMATION" in texte: x1_info = mot['x1']
                        
                frontiere_payer_deduire = (x1_payer + x1_deduire) / 2
                frontiere_deduire_info = (x1_deduire + x1_info) / 2
                frontiere_gauche = x1_payer - 80 
                
                lignes_visuelles = []
                for mot in mots:
                    ajoute = False
                    for ligne in lignes_visuelles:
                        if abs(mot['top'] - ligne[0]['top']) < 4:
                            ligne.append(mot)
                            ajoute = True
                            break
                    if not ajoute:
                        lignes_visuelles.append([mot])
                        
                lignes_visuelles.sort(key=lambda ligne: ligne[0]['top'])
                for ligne in lignes_visuelles:
                    ligne.sort(key=lambda mot: mot['x0'])
                    
                code_en_cours = None
                elements, payer, deduire, info = "", "", "", ""
                
                for ligne in lignes_visuelles:
                    texte_complet_ligne = " ".join([m['text'] for m in ligne]).upper()
                    
                    mots_fin_tableau = ["VOIR EXPLICATIONS", "TOTAUX DU MOIS", "TOTAL CHARGES", "COUT TOTAL"]
                    if any(mot in texte_complet_ligne for mot in mots_fin_tableau):
                        break
                        
                    premier_mot = ligne[0]['text']
                    
                    if re.match(r"^\d{4,6}$", premier_mot):
                        if code_en_cours:
                            # --- LA CONVERSION S'OPÈRE ICI ---
                            lignes_donnees.append([
                                code_en_cours, 
                                elements.strip(), 
                                texte_vers_nombre(payer), 
                                texte_vers_nombre(deduire), 
                                texte_vers_nombre(info)
                            ])
                            
                        code_en_cours = premier_mot
                        elements, payer, deduire, info = "", "", "", ""
                        
                        for mot in ligne[1:]:
                            texte = mot['text']
                            
                            if re.match(r"^-?\d+,\d{2}$", texte):
                                fin_montant = mot['x1']
                                
                                if fin_montant < frontiere_gauche: pass
                                elif fin_montant < frontiere_payer_deduire: payer = texte
                                elif fin_montant < frontiere_deduire_info: deduire = texte
                                else: info = texte
                                    
                            elif texte != "€":
                                elements += " " + texte
                                
                    elif code_en_cours:
                        for mot in ligne:
                            texte = mot['text']
                            
                            if re.match(r"^-?\d+,\d{2}$", texte):
                                fin_montant = mot['x1']
                                
                                if fin_montant < frontiere_gauche: pass
                                elif fin_montant < frontiere_payer_deduire:
                                    if not payer: payer = texte
                                elif fin_montant < frontiere_deduire_info:
                                    if not deduire: deduire = texte
                                else:
                                    if not info: info = texte
                                    
                            elif texte != "€":
                                elements += " " + texte

                if code_en_cours:
                    # --- ET ICI POUR LA DERNIÈRE LIGNE ---
                    lignes_donnees.append([
                        code_en_cours, 
                        elements.strip(), 
                        texte_vers_nombre(payer), 
                        texte_vers_nombre(deduire), 
                        texte_vers_nombre(info)
                    ])

        if lignes_donnees:
            colonnes = ["CODE", "ÉLÉMENTS", "A PAYER", "A DÉDUIRE", "POUR INFORMATION"]
            df = pd.DataFrame(lignes_donnees, columns=colonnes)
            
            nom_onglet = nom_fichier[:31] 
            
            df.to_excel(writer, index=False, sheet_name=nom_onglet)
            print(f"  -> OK : {len(df)} lignes extraites vers l'onglet '{nom_onglet}'")
        else:
            print(f"  -> ÉCHEC : Aucune donnée lue pour ce fichier.")

print(f"\nTerminé ! Les nombres sont maintenant exploitables dans :\n{chemin_excel}")