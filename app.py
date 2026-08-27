import streamlit as st
from pptx import Presentation
from google import genai
import io
import time

st.set_page_config(page_title="Traducteur de Fiches Produits PPT", page_icon="🌐")

st.title("🌐 Traducteur de Fiches Produits PowerPoint")
st.write("Conserve la charte graphique, le design et les zones de texte modifiables !")

# Clé API et Langue
api_key = st.text_input("Entre ta clé d'API Google AI Studio :", type="password")
langue_cible = st.selectbox("Choisir la langue de traduction :", ["Anglais", "Espagnol"])

uploaded_file = st.file_uploader("Dépose ton fichier PowerPoint (.pptx) ici", type=["pptx"])

def traduire_texte(client, texte, langue):
    if not texte.strip() or len(texte.strip()) <= 1:
        return texte
        
    prompt = f"""Tu es un traducteur expert en matériel industriel, signalisation et BTP.
Traduis le texte suivant en {langue}. 
CONSIGNES STRICTES :
- Conserve le ton technique, concis et professionnel.
- Ne traduis PAS les normes (CE, NF, WL9, PL3, etc.), les dimensions (mm, kg, °C), ni les noms de marque (KELIAS, Px3 Plus, etc.).
- Renvoie UNIQUEMENT le texte traduit, sans aucun commentaire ni guillemets.

Texte à traduire :
{texte}"""

    # Gestion des réessais en cas de surcharge serveur (Error 503)
    max_retries = 3
    for attempt in range(max_retries):
        try:
            response = client.models.generate_content(
                model='gemini-3.6-flash',
                contents=prompt
            )
            return response.text.strip()
        except Exception as e:
            if "503" in str(e) and attempt < max_retries - 1:
                time.sleep(2)  # Pause de 2 secondes avant de réessayer
                continue
            else:
                raise e

if uploaded_file and api_key:
    if st.button("🚀 Lancer la traduction"):
        try:
            client = genai.Client(api_key=api_key.strip())
            prs = Presentation(uploaded_file)
            
            progress_bar = st.progress(0)
            st.info("Traduction en cours... Veuillez patienter.")
            
            total_shapes = sum(len(slide.shapes) for slide in prs.slides)
            count = 0

            for slide in prs.slides:
                for shape in slide.shapes:
                    count += 1
                    if shape.has_text_frame:
                        for paragraph in shape.text_frame.paragraphs:
                            for run in paragraph.runs:
                                if run.text.strip():
                                    run.text = traduire_texte(client, run.text, langue_cible)
                    
                    if shape.has_table:
                        for row in shape.table.rows:
                            for cell in row.cells:
                                for paragraph in cell.text_frame.paragraphs:
                                    for run in paragraph.runs:
                                        if run.text.strip():
                                            run.text = traduire_texte(client, run.text, langue_cible)
                                            
                    progress_bar.progress(min(count / total_shapes, 1.0))

            output = io.BytesIO()
            prs.save(output)
            output.seek(0)
            
            st.success("✅ Traduction terminée !")
            
            st.download_button(
                label="📥 Télécharger le PowerPoint Traduit",
                data=output,
                file_name=f"Fiche_Produit_{langue_cible}.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
            )

        except Exception as e:
            st.error(f"Une erreur est survenue : {e}")
