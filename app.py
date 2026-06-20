import streamlit as st
import google.generativeai as genai
from docx import Document
import io

# 1. पेज की सेटिंग और टाइटल
st.set_page_config(page_title="राजगढ़ शासकीय प्रारूपक", page_icon="🏛️", layout="wide")

# 2. मध्य प्रदेश शासन - शासकीय हेडर और सुंदर डिज़ाइन
st.markdown("""
    <div style="background-color:#A32A2A; padding:18px; border-radius:10px; text-align:center; margin-bottom:25px; border: 3px solid #FF9933;">
        <h1 style="color:white; margin:0; font-family:'Mangal', sans-serif; font-size:28px;">कार्यालय कलेक्टेरेट (जनजातीय कार्य तथा अनुसूचित जाति कल्याण विभाग)</h1>
        <h2 style="color:yellow; margin:8px 0 0 0; font-family:'Mangal', sans-serif; font-size:22px;">जिला राजगढ़ (ब्यावरा) मध्य प्रदेश</h2>
        <p style="color:white; margin:5px 0 0 0; font-size:16px;">✨ शासकीय कार्यप्रणाली आलेखन सहायक (AI Powered) ✨</p>
    </div>
""", unsafe_allow_html=True)

# 3. साइडबार में सेटिंग और परमानेंट API की (Secrets) का जुगाड़
st.sidebar.title("🛠️ विन्यास (Configuration)")

# यहाँ सिस्टम चेक करेगा कि क्या चाबी बैकएंड में लॉक है, नहीं तो हाथ से डालने का डिब्बा देगा
if "GEMINI_API_KEY" in st.secrets:
    api_key = st.secrets["GEMINI_API_KEY"]
    st.sidebar.success("🔑 API कुंजी सुरक्षित रूप से बैकएंड में लॉक है!")
else:
    api_key = st.sidebar.text_input("गूगल जेमिनी API कुंजी दर्ज करें (Gemini API Key):", type="password", help="स्थाई रूप से लॉक करने के लिए Streamlit Secrets का उपयोग करें।")

st.sidebar.markdown("---")
st.sidebar.info("💡 *सुझाव:* मोबाइल पर टाइप करते समय कीबोर्ड के *माइक (Microphone) बटन* को दबाकर शुद्ध हिंदी में बोलकर भी आप पूरा विवरण लिख सकते हैं!")

# 4. मुख्य पेज: रेडी-मेड शासकीय फॉर्मेट्स (Dropdown Template)
st.subheader("📋 पत्र का प्रारूप और विवरण")

template_type = st.selectbox(
    "पत्र का प्रकार चुनें (Select Document Template):",
    [
        "सामान्य शासकीय पत्र (General Official Letter)",
        "कारण बताओ नोटिस का जवाब (Reply to Show-Cause Notice)",
        "ग्राम पंचायत प्रस्ताव / प्रतिवेदन (Gram Panchayat Resolution/Report)",
        "वरिष्ठ कार्यालयों को भेजा जाने वाला प्रतिवेदन (Official Report)"
    ]
)

# विवरण दर्ज करने का बड़ा डिब्बा
user_input = st.text_area("पत्र का मुख्य विषय, संदर्भ या संक्षिप्त विवरण यहाँ दर्ज करें:", height=150, placeholder="जैसे: ग्राम पंचायत में निर्माण कार्य की पूर्णता प्रमाण पत्र जमा करने बाबत्...")

# फाइल अपलोडर
uploaded_file = st.file_uploader("संदर्भ दस्तावेज़ या पुराना पत्र अपलोड करें (यदि आवश्यक हो):", type=["pdf", "png", "jpg", "jpeg"])

# 5. ड्राफ्ट जनरेट करने का बटन और बैकएंड प्रोसेस
if st.button("प्रारूप तैयार करें (Generate Draft)", type="primary"):
    if not api_key:
        st.error("⚠️ कृपया पहले साइडबार में अपनी जेमिनी API कुंजी दर्ज करें या बैकएंड में लॉक करें!")
    elif not user_input:
        st.warning("⚠️ कृपया पत्र का विवरण या विषय दर्ज करें!")
    else:
        with st.spinner("🤖 आपका आधिकारिक शासकीय प्रारूप तैयार किया जा रहा है, कृपया प्रतीक्षा करें..."):
            try:
                # एआई मॉडल कॉन्फ़िगरेशन
                genai.configure(api_key=api_key)
                model = genai.GenerativeModel('gemini-1.5-flash')
                
                # सरकारी भाषा के लिए स्पेशल निर्देश (System Prompt)
                system_prompt = f"""
                You are a senior official drafting expert for the Government of Madhya Pradesh. 
                Write a highly professional and formal official document/letter in pure administrative Hindi (शासकीय हिंदी).
                Use standard government formatting including sections like: 'प्रति', 'विषय', 'संदर्भ', 'महोदय', and appropriate endings like 'भवदीय/आदेशानुसार'.
                
                Document Type to generate: {template_type}
                Context provided by user: {user_input}
                
                Ensure the tone is perfectly legal, administrative, and adheres to MP government protocols.
                """
                
                response = model.generate_content(system_prompt)
                draft_text = response.text
                
                # परिणाम दिखाना
                st.success("✨ शासकीय प्रारूप सफलतापूर्वक तैयार कर दिया गया है!")
                final_draft = st.text_area("संपादित करें (Edit Document Text):", draft_text, height=350)
                
                # 6. एमएस वर्ड (.docx) फाइल बनाने का जादू
                doc = Document()
                doc.add_heading('कार्यालय कलेक्टेरेट राजगढ़ (म.प्र.)', level=1)
                doc.add_paragraph(final_draft)
                
                # मेमोरी में फाइल सेव करना
                bio = io.BytesIO()
                doc.save(bio)
                
                st.markdown("---")
                # डाउनलोड बटन
                st.download_button(
                    label="💾 वर्ड फाइल डाउनलोड करें (.docx)",
                    data=bio.getvalue(),
                    file_name="शासकीय_प्रारूप_राजगढ़.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                
            except Exception as e:
                st.error(f"❌ त्रुटि उत्पन्न हुई: {e}")
