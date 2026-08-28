import fitz  # PyMuPDF για ανάγνωση PDF
import pandas as pd
import json
import io
import streamlit as st
from google import genai
from google.genai import types

st.title("🧪 AI-Powered SDS & TDS Parser (Gemini)")
st.write(
    "Ανέβασε τα αρχεία PDF σου για να εξαχθούν αυτόματα τα δεδομένα τους μέσω"
    " AI σε Excel!"
)

# 1. Ανάκτηση του API Key από τα Streamlit Secrets
# (Θα το ρυθμίσουμε σε 1 λεπτό στις ρυθμίσεις του Streamlit Cloud)
try:
  api_key = st.secrets["GOOGLE_API_KEY"]
except Exception:
  st.error(
      "Δεν βρέθηκε το GOOGLE_API_KEY στα Secrets του Streamlit. Παρακαλώ"
      " πρόσθεσέ το στις ρυθμίσεις της εφαρμογής σου!"
  )
  st.stop()

# Αρχικοποίηση του Gemini Client
client = genai.Client(api_key=api_key)


def analyze_pdf_with_ai(bytes_data, filename):
  try:
    doc = fitz.open(stream=bytes_data, filetype="pdf")
    full_text = ""
    for page in doc:
      full_text += page.get_text()
  except Exception as e:
    st.error(f"Σφάλμα ανάγνωσης PDF {filename}: {e}")
    return []

  prompt = f"""
    Διάβασε προσεκτικά το παρακάτω κείμενο από έγγραφο πρώτης ύλης καλλυντικών.
    Εξάγε τις πληροφορίες σε δομημένη μορφή JSON με ένα κλειδί "materials", το οποίο περιέχει λίστα από αντικείμενα με τα εξής πεδία:
    - trade_name (Εμπορική ονομασία)
    - breakdown_composition (Σύνθεση / INCI)
    - content_percentage (Ποσοστό % ή περιεκτικότητα)
    - concentration (Συγκέντρωση / δόση χρήσης)
    - impurity_1 (1η ακαθαρσία αν υπάρχει, αλλιώς "—")
    - impurity_2 (2η ακαθαρσία αν υπάρχει, αλλιώς "")
    - impurity_3 (3η ακαθαρσία αν υπάρχει, αλλιώς "")

    Κείμενο:
    {full_text[:8000]}
    """

  try:
    response = client.models.generate_content(
        model="gemini-2.5-flash",  # Διόρθωση μοντέλου σε σταθερό διαθέσιμο
        contents=prompt,
        config=types.GenerateContentConfig(
            system_instruction=(
                "Εσύ είσαι ένας έμπειρος χημικός και αναλυτής δεδομένων"
                " καλλυντικών. Απάντησε ΜΟΝΟ σε έγκυρη μορφή JSON."
            ),
            response_mime_type="application/json",
        ),
    )

    result_json = json.loads(response.text)
    if isinstance(result_json, list):
      return result_json
    return result_json.get("materials", [])

  except Exception as e:
    st.error(f"Σφάλμα κατά την επεξεργασία του {filename}: {e}")
    return []


# 2. Widget για πολλαπλή μεταφόρτωση PDF
uploaded_files = st.file_uploader(
    "Επέλεξε αρχεία PDF", type=["pdf"], accept_multiple_files=True
)

if uploaded_files:
  if st.button("🚀 Έναρξη Ανάλυσης με AI"):
    all_extracted_rows = []

    with st.spinner("Το Gemini αναλύει τα έγγραφα... Παρακαλώ περιμένετε."):
      for uploaded_file in uploaded_files:
        bytes_data = uploaded_file.read()
        items = analyze_pdf_with_ai(bytes_data, uploaded_file.name)
        for item in items:
          all_extracted_rows.append(item)

    if all_extracted_rows:
      df = pd.DataFrame(all_extracted_rows)
      st.success("✅ Επιτυχία! Η ανάλυση ολοκληρώθηκε.")
      st.dataframe(df)

      # Δημιουργία αρχείου Excel στη μνήμη για λήψη
      output = io.BytesIO()
      with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
      excel_data = output.getvalue()

      st.download_button(
          label="📥 Λήψη Excel (ai_extracted_raw_materials.xlsx)",
          data=excel_data,
          file_name="ai_extracted_raw_materials.xlsx",
          mime=(
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          ),
      )
    else:
      st.warning("⚠️ Δεν βρέθηκαν δεδομένα προς εξαγωγή.")
