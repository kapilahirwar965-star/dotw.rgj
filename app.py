import streamlit as st
from docx import Document
import io

# सभी एआई लाइब्रेरी को बैकएंड में लोड करना
try:
    import google.generativeai as genai
except ImportError:
    pass

try:
    from openai import OpenAI
except ImportError:
    pass

try:
    import anthropic
except ImportError:
    pass

# 1. पेज की सेटिंग
st.set_page_config(page_title="राजगढ़ शासकीय प्रारूपक", page_icon="🏛️", layout="wide")

# 2. मध्य प्रदेश शासन - शासकीय हेडर और डिज़ाइन
st.markdown("""
    <div style="background-color:#A32A2A; padding:18px; border-radius:10px; text-align:center; margin-bottom:25px; border: 3px solid #FF9933;">
        <h1 style="color:white; margin:0; font-family:'Mangal', sans-serif; font-size:28px;">कार्यालय कलेक्टेरेट (जनजातीय कार्य तथा अनुसूचित जाति कल्याण विभाग)</h1>
        <h2 style="color:yellow; margin:8px 0 0 0; font-family:'Mangal', sans-serif; font-size:22px;">जिला राजगढ़ (ब्यावरा) मध्य प्रदेश</h2>
        <p style="color:white; margin:5px 0 0 0; font-size:16px;">✨ शासकीय कार्यप्रणाली आलेखन सहायक (Multi-Model AI) ✨</p>
    </div>
""", unsafe_allow_html=True)

# 3. साइडबार विन्यास (Configuration)
st.sidebar.title("🛠️ विन्यास (Configuration)")

# मॉडल चुनने का ड्रॉपडाउन ऑप्शन
provider = st.sidebar.selectbox(
    "मॉडल प्रदाता चुनें (Select Model Provider):", 
    ["Google Gemini", "Anthropic Claude", "OpenAI ChatGPT"]
)

# तीनों एआई की चाबियाँ डालने के डिब्बे (यदि Secrets में लॉक हैं तो वहां से उठा लेगा)
gemini_key = st.sidebar.text_input("गूगल जेमिनी API कुंजी (Gemini API Key):", type="password", value=st.secrets.get("GEMINI_API_KEY", ""))
claude_key = st.sidebar.text_input("एंथ्रोपिक क्लॉड API कुंजी (Claude API Key):", type="password", value=st.secrets.get("CLAUDE_API_KEY", ""))
openai_key = st.sidebar.text_input("OpenAI API कुंजी (OpenAI API Key):", type="password", value=st.secrets.get("OPENAI_API_KEY", ""))

st.sidebar.markdown("---")
st.sidebar.info("💡 *सुझाव:* मोबाइल पर टाइप करते समय कीबोर्ड के *माइक (Microphone) बटन* को दबाकर शुद्ध हिंदी में बोलकर भी विवरण दर्ज कर सकते हैं!")

# 4. मुख्य पेज: प्रारूप और इनपुट
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

user_input = st.text_area("पत्र का मुख्य विषय, संदर्भ या संक्षिप्त विवरण यहाँ दर्ज करें:", height=150, placeholder="जैसे: ग्राम पंचायत में निर्माण कार्य की पूर्णता प्रमाण पत्र जमा करने बाबत्...")
uploaded_file = st.file_uploader("संदर्भ दस्तावेज़ या पुराना पत्र अपलोड करें (यदि आवश्यक हो):", type=["pdf", "png", "jpg", "jpeg"])

# 5. बटन दबाने पर एआई प्रोसेस
if st.button("प्रारूप तैयार करें (Generate Draft)", type="primary"):
    system_prompt = f"""
    You are a senior official drafting expert for the Government of Madhya Pradesh. 
    Write a highly professional and formal official document/letter in pure administrative Hindi (शासकीय हिंदी).
    Use standard government formatting including sections like: 'प्रति', 'विषय', 'संदर्भ', 'महोदय', and appropriate endings like 'भवदीय/आदेशानुसार'.
    
    Document Type to generate: {template_type}
    Context provided by user: {user_input}
    
    Ensure the tone is perfectly legal, administrative, and adheres to MP government protocols. Do not include any english explanations, give only the clean Hindi document text.
    """
    
    draft_text = ""
    error_flag = False
    
    with st.spinner(f"🤖 {provider} द्वारा आपका शासकीय प्रारूप तैयार किया जा रहा है..."):
        try:
            if provider == "Google Gemini":
                if not gemini_key:
                    st.error("⚠️ कृपया साइडबार में जेमिनी API कुंजी दर्ज करें!")
                    error_flag = True
                else:
                    genai.configure(api_key=gemini_key)
                    model = genai.GenerativeModel('gemini-1.5-flash')
                    response = model.generate_content(system_prompt)
                    draft_text = response.text
                    
            elif provider == "Anthropic Claude":
                if not claude_key:
                    st.error("⚠️ कृपया साइडबार में क्लॉड API कुंजी दर्ज करें!")
                    error_flag = True
                else:
                    client = anthropic.Anthropic(api_key=claude_key)
                    message = client.messages.create(
                        model="claude-3-5-sonnet-20241022",
                        max_tokens=2000,
                        messages=[{"role": "user", "content": system_prompt}]
                    )
                    draft_text = message.content[0].text
                    
            elif provider == "OpenAI ChatGPT":
                if not openai_key:
                    st.error("⚠️ कृपया साइडबार में OpenAI API कुंजी दर्ज करें!")
                    error_flag = True
                else:
                    client = OpenAI(api_key=openai_key)
                    response = client.chat.completions.create(
                        model="gpt-4o-mini",
                        messages=[{"role": "user", "content": system_prompt}]
                    )
                    draft_text = response.choices[0].message.content
                    
            # परिणाम दिखाना और वर्ड फाइल जनरेट करना
            if not error_flag and draft_text:
                st.success("✨ शासकीय प्रारूप सफलतापूर्वक तैयार कर दिया गया है!")
                final_draft = st.text_area("संपादित करें (Edit Document Text):", draft_text, height=350)
                
                # वर्ड डॉक्यूमेंट क्रिएशन
                doc = Document()
                doc.add_heading('कार्यालय कलेक्टेरेट राजगढ़ (म.प्र.)', level=1)
                doc.add_paragraph(final_draft)
                
                bio = io.BytesIO()
                doc.save(bio)
                
                st.markdown("---")
                st.download_button(
                    label="💾 वर्ड फाइल डाउनलोड करें (.docx)",
                    data=bio.getvalue(),
                    file_name="शासकीय_प्रारूप_राजगढ़.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
        except Exception as e:
            st.error(f"❌ त्रुटि उत्पन्न हुई: {e}")
