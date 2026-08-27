import streamlit as st
from pptx import Presentation
from google import genai
import io
import json

st.set_page_config(page_title="Traducteur de Fiches Produits PPT", page_icon="🌐")

st.title("🌐 Traducteur de Fiches Produits PowerPoint")
st.write("Conserve la charte graphique, le design et les zones de texte modifiables !")

api_key = st.text_input("Entre ta clé d'API Google AI Studio :", type="password")
langue_cible = st.selectbox("Choisir la langue de traduction :", ["Anglais", "Espagnol"])

uploaded_file = st.file_uploader("Dépose ton fichier PowerPoint (.pptx) ici", type=["pptx"])

if uploaded_file and api_key:
    if st.button("🚀 Lancer la traduction"):
        try:
            client = genai.Client(api_key=api_key.strip())
            prs = Presentation(uploaded_file)
            
            st.info("Extraction et traduction globale en cours... Veuillez patienter.")
            
            # 1. Collecter tous les morceaux de texte
            text_runs = []
            for slide in prs.slides:
                for shape in slide.shapes:
                    if shape.has_text_frame:
                        for paragraph in shape.text_frame.paragraphs:
                            for run in paragraph.runs:
                                if run.text.strip():
                                    text_runs.append(run)
                    if shape.has_table:
                        for row in shape.table.rows:
                            for cell in row.cells:
                                for paragraph in cell.text_frame.paragraphs:
                                    for run in paragraph.runs:
                                        if run.text.strip():
                                            text_runs.append(run)

            if not text_runs:
                st.warning("⚠️ Aucun texte à traduire n'a été trouvé dans ce fichier.")
            else:
                # 2. Préparer la liste des textes
                original_texts = [r.text for r in text_runs]
                
                prompt = f"""Tu es un traducteur expert en matériel industriel, signalisation et BTP.
Traduis la liste de textes suivante en {langue_cible}.

CONSIGNES STRICTES :
- Conserve le ton technique, concis et professionnel.
- Ne traduis PAS les normes (CE, NF, WL9, PL3, etc.), les dimensions (mm, kg, °C), ni les noms de marque (KELIAS, Px3 Plus, etc.).
- Renvoie UNIQUEMENT un tableau JSON contenant les traductions dans le même ordre exact sous la forme : ["texte1", "texte2", ...]

Textes à traduire :
{json.dumps(original_texts, ensure_ascii=False)}"""

                # 3. Envoi à l'API
                response = client.models.generate_content(
                    model='gemini-3.6-flash',
                    contents=prompt
                )
                
                # Lecture de la réponse JSON
                clean_response = response.text.strip().removeprefix("```json").removeprefix("```").removesuffix("```").strip()
                translated_texts = json.loads(clean_response)

                # 4. Réinjection des textes traduits
                for run, trad in zip(text_runs, translated_texts):
                    run.text = trad

                output = io.BytesIO()
                prs.save(output)
                output.seek(0)
                
                st.success("✅ Traduction terminée avec succès !")
                
                st.download_button(
                    label="📥 Télécharger le PowerPoint Traduit",
                    data=output,
                    file_name=f"Fiche_Produit_{langue_cible}.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )

        except Exception as e:
            error_msg = str(e)
            
            # Personnalisation des messages d'erreur pour les utilisateurs
            if "429" in error_msg or "RESOURCE_EXHAUSTED" in error_msg:
                st.error("⏳ **Quota journalier dépassé.** La limite gratuite de Google a été atteinte pour aujourd'hui. Réessaie demain ou utilise une autre clé API.")
            elif "API_KEY_INVALID" in error_msg or "400" in error_msg:
                st.error("🔑 **Clé API invalide.** Vérifie que tu as bien copié-collé ta clé Google AI Studio sans espace en trop.")
            elif "503" in error_msg or "UNAVAILABLE" in error_msg:
                st.error("🌐 **Serveur temporairement indisponible.** Les serveurs de Google sont en surcharge. Patiente 1 minute et relance.")
            else:
                st.error(f"❌ **Une erreur est survenue :** {error_msg}")
