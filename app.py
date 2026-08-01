import streamlit as st
from google import genai
import PIL.Image
import io

# Set up the look of the web app
st.set_page_config(page_title="Court Daily Report Generator", page_icon="⚖️", layout="centered")

st.title("📄 Court Daily Report Generator")
st.write("Upload a raw document or photo, and AI will format it into your court report template.")

# Securely get API key from Streamlit secrets
try:
    api_key = st.secrets["GEMINI_API_KEY"]
except KeyError:
    st.error("API Key not found. Please set GEMINI_API_KEY in your Streamlit secrets.")
    st.stop()

# Initialize the Gemini AI client
client = genai.Client(api_key=api_key)

# Create a file uploader button
uploaded_file = st.file_uploader("Upload Document Photo (JPG/PNG)", type=["jpg", "jpeg", "png"])

if uploaded_file is not None:
    st.image(uploaded_file, caption="Uploaded Document", use_column_width=True)
    
    if st.button("Generate Formatted Report"):
        with st.spinner("Reading document and organizing data... Please wait."):
            
            # Prepare the image for the AI
            image = PIL.Image.open(uploaded_file)
            
            # Tell the AI exactly how to format the table
            prompt = """
            You are a legal data extraction assistant. Extract the information from this image and format it into a single Markdown table with the exact following headers:
            | थाना | अ०सं० | वाद सं० | धारा | वादी का नाम व पता | विवेचक का नाम | निर्णय का दि० | घटना का संक्षिप्त विवरण | न्यायालय के आदेश का विवरण |
            
            Instructions:
            1. If any information is missing from the image, leave the cell blank or write "-".
            2. Output ONLY the Markdown table and nothing else. No introductory text.
            """
            
            try:
                # Send to Gemini
                response = client.models.generate_content(
                    model="gemini-3.6-flash", 
                    contents=[image, prompt]
                )
                
                st.success("Report Generated Successfully!")
                
                # Display the beautifully formatted table
                st.markdown(response.text)
                
                st.info("💡 **Mobile Instructions:** Long-press the table to highlight and copy it, then paste it directly into your printing template or document.")
                
            except Exception as e:
                st.error(f"An error occurred: {e}")
