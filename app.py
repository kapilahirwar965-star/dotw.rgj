import streamlit as st
import google.generativeai as genai

# 1. Page Configuration
st.set_page_config(page_title="राजगढ़ शासकीय प्रारूपक", page_icon="🏛️", layout="wide")

# 2. Original Government Banner 
st.markdown("""
    <div style="background-color:#A32A2A; padding:20px; border-radius:10px; text-align:center; margin-bottom:20px;">
        <h1 style="color:white; margin:0; font-family:'Mangal', sans-serif; font-size:26px;">कार्यालय कलेक्टेरेट (जनजातीय कार्य तथा अनुसूचित जाति कल्याण विभाग)</h1>
        <h2 style="color:yellow; margin:5px 0 0 0; font-family:'Mangal', sans-serif; font-size:20px;">जिला राजगढ़ (ब्यावरा) मध्य प्रदेश</h2>
        <p style="color:white; margin:5px 0 0 0; font-size:14px;">शासकीय कार्यप्रणाली आलेखन सहायक - जिला राजगढ़ (ब्यावरा) मध्य प्रदेश</p>
    </div>
""", unsafe_allow_html=True)

# 3. Sidebar Configuration
st.sidebar.title("⚙️ विन्यास (Configuration)")
gemini_key = st.sidebar.text_input("गूगल जेमिनी API कुंजी (Gemini API Key):", type="password", value="AlzaSy-Key...")
claude_key = st.sidebar.text_input("एंथ्रोपिक क्लॉड API कुंजी (Claude API Key):", type="password", value="sk-ant-...")
openai_key = st.sidebar.text_input("OpenAI API कुंजी (OpenAI API Key):", type="password", value="sk-proj-...")
bytez_key = st.sidebar.text_input("Bytez API कुंजी (Bytez API Key):", type="password", value="bz-...")

provider = st.sidebar.selectbox("मॉडल प्रदाता (Model Provider):", ["Google Gemini", "Anthropic Claude", "OpenAI ChatGPT"])

# 4. Two-Column Side-by-Side Layout
col1, col2 = st.columns(2)

with col1:
    st.markdown("### 📥 इनपुट स्रोत (Input Source)")
    user_input = st.text_area(
        "पत्र का मुख्य विषय या विवरण यहाँ दर्ज करें:", 
        height=150, 
        placeholder="जैसे: ग्राम पंचायत में निर्माण कार्य की पूर्णता प्रमाण पत्र जमा करने बाबत्..."
    )
    uploaded_file = st.file_uploader("दस्तावेज़/चित्र अपलोड करें (Upload Document/Image)", type=["pdf", "png", "jpg", "jpeg"])
    if st.button("Upload"):
        st.info("फ़ाइल अपलोड हो गई है‌।")

with col2:
    st.markdown("### 📑 प्रारूप निर्माण एवं संपादन (Draft Generation)")
    generate_button = st.button("प्रारूप तैयार करें (Generate Draft)", type="primary")
    
    draft_text = ""
    if generate_button:
        if not user_input:
            st.warning("⚠️ कृपया पहले विवरण दर्ज करें!")
        else:
            with st.spinner("🤖 प्रारूप तैयार किया जा रहा है..."):
                try:
                    if provider == "Google Gemini":
                        genai.configure(api_key=gemini_key)
                        model = genai.GenerativeModel('gemini-1.5-flash')
                        system_prompt = f"Write an official government letter in professional administrative Hindi based on: {user_input}"
                        response = model.generate_content(system_prompt)
                        draft_text = response.text
                    else:
                        draft_text = "यह मॉडल अभी बैकएंड कॉन्फ़िगरेशन में है, कृपया Google Gemini चुनें।"
                except Exception as e:
                    draft_text = f"त्रुटि: {e}"

    st.text_area("संपादित करें (Edit Draft Text):", value=draft_text, height=250)
