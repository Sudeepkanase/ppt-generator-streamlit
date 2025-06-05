import streamlit as st
import os
from ppt_generator import PPTGenerator
from datetime import datetime

# ✅ Streamlit secrets
api_key_default = st.secrets.get("GROQ_API_KEY", "")
model_name_default = st.secrets.get("MODEL_NAME", "provider-4/claude-3.5-haiku")

st.set_page_config(
    page_title="Professional PPT Generator",
    page_icon="📊",
    layout="wide"
)

# Custom CSS
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        text-align: center;
        margin-bottom: 2rem;
        color: #2c3e50;
    }
    .stButton > button {
        background: linear-gradient(90deg, #0056b3 0%, #003366 100%);
        color: white;
        border: none;
        border-radius: 8px;
        padding: 0.75rem 1.5rem;
        font-weight: bold;
        width: 100%;
        font-size: 1.1rem;
    }
    .success-box {
        background: #e6f7ff;
        border: 1px solid #b3e0ff;
        border-radius: 8px;
        padding: 1.5rem;
        margin: 1.5rem 0;
        color: #003366;
    }
    .section-box {
        background: #f8f9fa;
        border-radius: 8px;
        padding: 1.5rem;
        margin-bottom: 1.5rem;
    }
</style>
""", unsafe_allow_html=True)

def initialize_session_state():
    if 'presentation_generated' not in st.session_state:
        st.session_state.presentation_generated = False
    if 'content_sections' not in st.session_state:
        st.session_state.content_sections = ["Overview", "Key Concepts", "Implementation", "Case Studies", "Conclusion"]
    if 'generating' not in st.session_state:
        st.session_state.generating = False
    if 'generated_file_path' not in st.session_state:
        st.session_state.generated_file_path = None

def generate_presentation(topic: str, content_sections: list, api_key: str, model_name: str):
    try:
        st.session_state.generating = True
        
        with st.spinner("🔍 Creating professional presentation..."):
            generator = PPTGenerator(api_key, model_name)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            output_filename = f"professional_{topic.replace(' ', '_')}_{timestamp}.pptx"
            
            output_path = generator.create_presentation(topic, content_sections, output_filename)
            
            st.session_state.presentation_generated = True
            st.session_state.generated_file_path = output_path
            st.session_state.generating = False
            
            return output_path
            
    except Exception as e:
        st.session_state.generating = False
        st.error(f"Error generating presentation: {str(e)}")
        return None

def main():
    initialize_session_state()
    
    st.markdown('<h1 class="main-header">📊 Professional PowerPoint Generator</h1>', unsafe_allow_html=True)
    
    # API Configuration
    with st.expander("🔑 API Configuration", expanded=True):
        col1, col2 = st.columns(2)
        
        with col1:
            api_key = st.text_input(
                "API Key", 
                type="password",
                value=api_key_default,
                help="Enter your Groq API key"
            )
        
        with col2:
            model_name = st.text_input(
                "Model Name",
                value=model_name_default,
                help="AI model for content generation"
            )
    
    # Presentation Details
    with st.expander("📝 Presentation Details", expanded=True):
        topic = st.text_input(
            "Presentation Topic:",
            placeholder="e.g., Digital Transformation Strategy, AI in Healthcare...",
            help="Enter a specific topic for your presentation"
        )
        
        st.markdown("### 📋 Content Sections")
        num_sections = st.slider("Number of sections:", 3, 10, len(st.session_state.content_sections))
        
        # Adjust sections
        if len(st.session_state.content_sections) != num_sections:
            if len(st.session_state.content_sections) < num_sections:
                for i in range(len(st.session_state.content_sections), num_sections):
                    st.session_state.content_sections.append(f"Section {i+1}")
            else:
                st.session_state.content_sections = st.session_state.content_sections[:num_sections]
        
        content_sections = []
        cols = st.columns(2)
        
        for i in range(num_sections):
            with cols[i % 2]:
                section = st.text_input(
                    f"Section {i+1}:",
                    value=st.session_state.content_sections[i] if i < len(st.session_state.content_sections) else "",
                    key=f"section_{i}"
                )
                content_sections.append(section)
                if i < len(st.session_state.content_sections):
                    st.session_state.content_sections[i] = section
    
    # Validation
    is_valid = (topic and len(topic.strip()) >= 3 and 
                api_key and len(api_key.strip()) >= 10 and
                all(section.strip() for section in content_sections))
    
    # Generation
    st.markdown("---")
    st.markdown("## 🚀 Generate Presentation")
    
    if st.button("📊 Generate Professional PowerPoint", 
                disabled=not is_valid or st.session_state.generating,
                type="primary"):
        if is_valid:
            output_path = generate_presentation(topic, content_sections, api_key, model_name)
            
            if output_path and os.path.exists(output_path):
                st.markdown('<div class="success-box">✅ Professional presentation generated successfully!</div>', unsafe_allow_html=True)
    
    # Download Section
    if st.session_state.presentation_generated and st.session_state.generated_file_path:
        st.markdown("---")
        st.markdown("## 📥 Download Presentation")
        
        if os.path.exists(st.session_state.generated_file_path):
            file_size = os.path.getsize(st.session_state.generated_file_path) / (1024 * 1024)
            
            col1, col2 = st.columns(2)
            with col1:
                st.metric("File Size", f"{file_size:.1f} MB")
            with col2:
                st.metric("Total Slides", f"{len(content_sections) + 2}")
            
            with open(st.session_state.generated_file_path, "rb") as file:
                st.download_button(
                    label="⬇️ Download PowerPoint File",
                    data=file.read(),
                    file_name=os.path.basename(st.session_state.generated_file_path),
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
            
            if st.button("🔄 Create New Presentation"):
                st.session_state.presentation_generated = False
                st.session_state.generated_file_path = None
                st.rerun()

if __name__ == "__main__":
    main()
