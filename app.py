import streamlit as st
import json
import zipfile
import io
from datetime import datetime
from collections import defaultdict, Counter
import re
from docx import Document
from docx.shared import RGBColor, Pt
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.units import inch

# Page config
st.set_page_config(
    page_title="Conversation Archive Browser",
    page_icon="🌙",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for pastel goth aesthetic
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Georgia&display=swap');

    .stApp {
        background: linear-gradient(135deg, #1a0a1f 0%, #2d1b3d 50%, #1a0a1f 100%);
        font-family: 'Georgia', serif;
    }

    .main {
        background: rgba(255, 255, 255, 0.03);
        border-radius: 15px;
        padding: 20px;
    }

    h1, h2, h3 {
        color: #e0b3ff !important;
        font-family: 'Georgia', serif !important;
        text-shadow: 0 0 10px rgba(224, 179, 255, 0.5);
    }

    .stButton>button {
        background: linear-gradient(135deg, #b388ff 0%, #ff80ab 100%);
        color: #1a0a1f;
        border: none;
        border-radius: 10px;
        padding: 10px 24px;
        font-weight: bold;
        transition: all 0.3s;
        box-shadow: 0 4px 15px rgba(179, 136, 255, 0.4);
    }

    .stButton>button:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(179, 136, 255, 0.6);
    }

    .divider {
        border-bottom: 2px solid #b388ff;
        margin: 20px 0;
        opacity: 0.3;
    }

    .conversation-card {
        background: rgba(179, 136, 255, 0.1);
        border: 1px solid #b388ff;
        border-radius: 10px;
        padding: 15px;
        margin: 10px 0;
        box-shadow: 0 4px 10px rgba(0, 0, 0, 0.3);
    }

    .user-message {
        background: rgba(179, 229, 252, 0.15);
        border-left: 3px solid #b3e5fc;
        padding: 10px;
        margin: 5px 0;
        border-radius: 5px;
    }

    .assistant-message {
        background: rgba(178, 255, 193, 0.15);
        border-left: 3px solid #b2ffc1;
        padding: 10px;
        margin: 5px 0;
        border-radius: 5px;
    }

    .theme-badge {
        display: inline-block;
        padding: 3px 10px;
        margin: 3px;
        background: rgba(255, 128, 171, 0.2);
        border: 1px solid #ff80ab;
        border-radius: 12px;
        font-size: 0.85em;
    }

    .character-tag {
        display: inline-block;
        padding: 3px 8px;
        margin: 3px;
        background: rgba(179, 136, 255, 0.2);
        border: 1px solid #b388ff;
        border-radius: 8px;
        font-size: 0.8em;
    }

    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #2d1b3d 0%, #1a0a1f 100%);
    }

    [data-testid="stSidebar"] .stRadio > label {
        color: #e0b3ff !important;
    }

    .stTextInput>div>div>input {
        background: rgba(179, 136, 255, 0.1);
        color: #e0b3ff;
        border: 1px solid #b388ff;
        border-radius: 8px;
    }

    .stSelectbox>div>div>div {
        background: rgba(179, 136, 255, 0.1);
        color: #e0b3ff;
    }

    .stMultiSelect>div>div>div {
        background: rgba(179, 136, 255, 0.1);
        color: #e0b3ff;
    }
</style>
""", unsafe_allow_html=True)

# Theme detection keywords
THEMES = {
    "💕 Romance": ["love", "kiss", "romantic", "date", "relationship", "affection", "heart", "crush", "romance", "loving"],
    "✨ Fluff": ["cute", "sweet", "adorable", "wholesome", "soft", "tender", "gentle", "cozy", "warm"],
    "🔥 Explicit": ["nsfw", "explicit", "mature", "adult", "intimate", "sexual"],
    "😢 Angst": ["pain", "hurt", "sad", "crying", "angst", "suffer", "grief", "heartbreak", "tears", "sorrow"],
    "😱 Horror/Fear": ["horror", "terror", "scary", "fear", "nightmare", "scream", "blood", "dark", "creepy"],
    "⚔️ Action": ["fight", "battle", "action", "combat", "attack", "war", "weapon", "sword", "gun"],
    "😂 Humor": ["funny", "hilarious", "joke", "laugh", "comedy", "humor", "silly", "amusing"],
    "🎭 Drama": ["drama", "conflict", "tension", "argument", "betrayal", "secret", "reveal"],
    "🔮 Fantasy": ["magic", "fantasy", "wizard", "spell", "dragon", "mythical", "enchant", "sorcery"],
    "🚀 Sci-Fi": ["space", "robot", "alien", "future", "technology", "sci-fi", "cybernetic", "android"],
    "🌸 Slice of Life": ["daily", "routine", "ordinary", "everyday", "mundane", "normal", "casual"],
    "🎨 Creative Writing": ["story", "character", "plot", "narrative", "writing", "fiction", "creative"]
}

def parse_chatgpt_export(file_content):
    """Parse ChatGPT export JSON and extract all conversations"""
    try:
        data = json.loads(file_content)
        conversations = []

        for conv in data:
            if "mapping" not in conv:
                continue

            messages = []
            mapping = conv["mapping"]

            # Traverse the tree structure
            for node_id, node in mapping.items():
                if node.get("message"):
                    msg = node["message"]
                    if msg.get("content") and msg["content"].get("parts"):
                        messages.append({
                            "role": msg.get("author", {}).get("role", "unknown"),
                            "content": " ".join(msg["content"]["parts"]) if isinstance(msg["content"]["parts"], list) else str(msg["content"]["parts"]),
                            "timestamp": msg.get("create_time", 0)
                        })

            # Sort messages by timestamp
            messages.sort(key=lambda x: x.get("timestamp", 0))

            if messages:
                conversations.append({
                    "title": conv.get("title", "Untitled Conversation"),
                    "create_time": conv.get("create_time", 0),
                    "update_time": conv.get("update_time", 0),
                    "messages": messages
                })

        return conversations
    except Exception as e:
        st.error(f"Error parsing file: {str(e)}")
        return []

def detect_themes(conversation):
    """Detect themes in a conversation based on keywords"""
    text = " ".join([msg["content"].lower() for msg in conversation["messages"]])
    detected = []

    for theme, keywords in THEMES.items():
        if any(keyword in text for keyword in keywords):
            detected.append(theme)

    return detected

def extract_characters(conversation):
    """Extract character names from conversation"""
    text = " ".join([msg["content"] for msg in conversation["messages"]])

    # Find capitalized words (potential character names)
    words = re.findall(r'\b[A-Z][a-z]+(?:\s+[A-Z][a-z]+)*\b', text)

    # Filter out common words
    common_words = {"I", "The", "A", "An", "This", "That", "Here", "There", "What", "When",
                   "Where", "Why", "How", "Yes", "No", "Please", "Thank", "You", "It", "He",
                   "She", "They", "We", "My", "Your", "His", "Her", "Their", "Our"}

    characters = Counter([word for word in words if word not in common_words])

    return characters

def generate_rogues_gallery(all_conversations):
    """Generate character analytics for Rogue's Gallery"""
    character_data = defaultdict(lambda: {
        "thread_count": 0,
        "total_mentions": 0,
        "themes": Counter()
    })

    for conv in all_conversations:
        characters = extract_characters(conv)
        themes = detect_themes(conv)

        for char, count in characters.items():
            character_data[char]["thread_count"] += 1
            character_data[char]["total_mentions"] += count
            for theme in themes:
                character_data[char]["themes"][theme] += 1

    # Sort by thread count
    sorted_characters = sorted(
        character_data.items(),
        key=lambda x: (x[1]["thread_count"], x[1]["total_mentions"]),
        reverse=True
    )

    return sorted_characters

def export_to_word(conversation):
    """Export conversation to Word document"""
    doc = Document()

    # Title
    title = doc.add_heading(conversation["title"], 0)
    title.runs[0].font.color.rgb = RGBColor(179, 136, 255)

    # Metadata
    date_str = datetime.fromtimestamp(conversation["create_time"]).strftime("%Y-%m-%d %H:%M:%S")
    meta = doc.add_paragraph(f"Created: {date_str} | Messages: {len(conversation['messages'])}")
    meta.runs[0].font.size = Pt(10)
    meta.runs[0].font.color.rgb = RGBColor(128, 128, 128)

    doc.add_paragraph("")

    # Messages
    for msg in conversation["messages"]:
        role_color = RGBColor(100, 149, 237) if msg["role"] == "user" else RGBColor(144, 238, 144)

        p = doc.add_paragraph()
        role_run = p.add_run(f"{msg['role'].upper()}: ")
        role_run.font.color.rgb = role_color
        role_run.font.bold = True

        content_run = p.add_run(msg["content"])

        if msg.get("timestamp"):
            ts = datetime.fromtimestamp(msg["timestamp"]).strftime("%H:%M:%S")
            time_run = p.add_run(f" ({ts})")
            time_run.font.size = Pt(8)
            time_run.font.color.rgb = RGBColor(128, 128, 128)

        doc.add_paragraph("")

    # Save to bytes
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

def export_to_pdf(conversation):
    """Export conversation to PDF document"""
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter)
    story = []
    styles = getSampleStyleSheet()

    # Custom styles
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=24,
        textColor=colors.HexColor('#b388ff'),
        spaceAfter=12
    )

    user_style = ParagraphStyle(
        'UserMessage',
        parent=styles['Normal'],
        fontSize=11,
        textColor=colors.HexColor('#6495ed'),
        leftIndent=20,
        spaceAfter=10
    )

    assistant_style = ParagraphStyle(
        'AssistantMessage',
        parent=styles['Normal'],
        fontSize=11,
        textColor=colors.HexColor('#90ee90'),
        leftIndent=20,
        spaceAfter=10
    )

    # Title
    story.append(Paragraph(conversation["title"], title_style))

    # Metadata
    date_str = datetime.fromtimestamp(conversation["create_time"]).strftime("%Y-%m-%d %H:%M:%S")
    meta_text = f"Created: {date_str} | Messages: {len(conversation['messages'])}"
    story.append(Paragraph(meta_text, styles['Normal']))
    story.append(Spacer(1, 0.2*inch))

    # Messages
    for msg in conversation["messages"]:
        style = user_style if msg["role"] == "user" else assistant_style
        role_text = f"<b>{msg['role'].upper()}:</b> {msg['content']}"
        story.append(Paragraph(role_text, style))

    doc.build(story)
    buffer.seek(0)
    return buffer

# Initialize session state
if "conversations" not in st.session_state:
    st.session_state.conversations = []
if "gallery" not in st.session_state:
    st.session_state.gallery = []
if "selected_character" not in st.session_state:
    st.session_state.selected_character = None

# Sidebar navigation
with st.sidebar:
    st.markdown("# 🌙 Navigation")
    page = st.radio("", ["📤 Upload & Browse", "ℹ️ About"], label_visibility="collapsed")

    st.markdown("<div class='divider'></div>", unsafe_allow_html=True)

    # Rogue's Gallery
    if st.session_state.gallery:
        st.markdown("## 🦇 Rogue's Gallery")
        st.markdown("*Character Analytics*")

        if st.session_state.selected_character:
            if st.button("❌ Clear Filter"):
                st.session_state.selected_character = None
                st.rerun()

        st.markdown("### Top Characters")
        for char, data in st.session_state.gallery[:10]:
            with st.expander(f"🦇 {char}"):
                st.write(f"**Threads:** {data['thread_count']}")
                st.write(f"**Total Mentions:** {data['total_mentions']}")

                if data['themes']:
                    st.write("**Themes:**")
                    for theme, count in data['themes'].most_common(3):
                        st.write(f"  - {theme} ({count})")

                if st.button(f"📖 Show {char}'s threads", key=f"char_{char}"):
                    st.session_state.selected_character = char
                    st.rerun()

# Main content
if page == "📤 Upload & Browse":
    st.markdown("# 🌙 Conversation Archive Browser")
    st.markdown("<div class='divider'></div>", unsafe_allow_html=True)

    # File upload
    st.markdown("## 📦 Upload Archive")
    uploaded_file = st.file_uploader(
        "Upload ChatGPT export (ZIP or JSON)",
        type=["zip", "json"],
        help="Upload your ChatGPT conversation archive"
    )

    if uploaded_file:
        try:
            if uploaded_file.name.endswith(".zip"):
                with zipfile.ZipFile(uploaded_file, 'r') as zip_ref:
                    # Look for conversations.json in the zip
                    json_file = None
                    for name in zip_ref.namelist():
                        if "conversations.json" in name.lower():
                            json_file = name
                            break

                    if json_file:
                        with zip_ref.open(json_file) as f:
                            content = f.read().decode('utf-8')
                            conversations = parse_chatgpt_export(content)
                    else:
                        st.error("No conversations.json found in ZIP file")
                        conversations = []
            else:
                content = uploaded_file.read().decode('utf-8')
                conversations = parse_chatgpt_export(content)

            if conversations:
                st.session_state.conversations = conversations
                st.session_state.gallery = generate_rogues_gallery(conversations)
                st.success(f"✨ Found {len(conversations)} conversations!")
                st.info(f"🦇 Generated Rogue's Gallery with {len(st.session_state.gallery)} characters! Check the sidebar!")
        except Exception as e:
            st.error(f"Error processing file: {str(e)}")

    # Display conversations
    if st.session_state.conversations:
        st.markdown("<div class='divider'></div>", unsafe_allow_html=True)
        st.markdown("## 🔍 Browse & Search")

        # Search and filters
        col1, col2 = st.columns([2, 1])
        with col1:
            search_text = st.text_input("🔍 Search", placeholder="Search conversations...")
            search_in_messages = st.checkbox("Search in message content (not just titles)", value=False)

        with col2:
            sort_option = st.selectbox(
                "Sort by",
                ["Recent first", "Oldest first", "Title A-Z", "Most messages"]
            )

        # Theme filter
        theme_filter = st.multiselect(
            "🎨 Filter by Theme",
            list(THEMES.keys()),
            default=[]
        )

        # Character search
        character_search = st.text_input(
            "👤 Search Characters (comma-separated)",
            placeholder="e.g., Lyra, Lucius, Hook"
        )

        # Filter conversations
        filtered_convs = st.session_state.conversations

        # Apply gallery character filter
        if st.session_state.selected_character:
            filtered_convs = [
                conv for conv in filtered_convs
                if st.session_state.selected_character in extract_characters(conv)
            ]
            st.info(f"🔍 Showing threads featuring: **{st.session_state.selected_character}**")

        # Apply text search
        if search_text:
            if search_in_messages:
                filtered_convs = [
                    conv for conv in filtered_convs
                    if search_text.lower() in conv["title"].lower() or
                    any(search_text.lower() in msg["content"].lower() for msg in conv["messages"])
                ]
            else:
                filtered_convs = [
                    conv for conv in filtered_convs
                    if search_text.lower() in conv["title"].lower()
                ]

        # Apply theme filter
        if theme_filter:
            filtered_convs = [
                conv for conv in filtered_convs
                if any(theme in detect_themes(conv) for theme in theme_filter)
            ]

        # Apply character search
        if character_search:
            search_chars = [c.strip() for c in character_search.split(",")]
            filtered_convs = [
                conv for conv in filtered_convs
                if any(char in extract_characters(conv) for char in search_chars)
            ]

        # Sort conversations
        if sort_option == "Recent first":
            filtered_convs.sort(key=lambda x: x["create_time"], reverse=True)
        elif sort_option == "Oldest first":
            filtered_convs.sort(key=lambda x: x["create_time"])
        elif sort_option == "Title A-Z":
            filtered_convs.sort(key=lambda x: x["title"].lower())
        elif sort_option == "Most messages":
            filtered_convs.sort(key=lambda x: len(x["messages"]), reverse=True)

        st.markdown(f"### 📖 Conversations ({len(filtered_convs)})")

        # Display conversations
        for i, conv in enumerate(filtered_convs):
            with st.container():
                st.markdown("<div class='conversation-card'>", unsafe_allow_html=True)

                col1, col2 = st.columns([3, 1])

                with col1:
                    st.markdown(f"### {conv['title']}")
                    date_str = datetime.fromtimestamp(conv["create_time"]).strftime("%Y-%m-%d %H:%M")
                    st.markdown(f"📅 {date_str} | 💬 {len(conv['messages'])} messages")

                    # Themes
                    themes = detect_themes(conv)
                    if themes:
                        theme_html = " ".join([f"<span class='theme-badge'>{t}</span>" for t in themes])
                        st.markdown(theme_html, unsafe_allow_html=True)

                    # Characters
                    characters = extract_characters(conv)
                    if characters:
                        top_chars = characters.most_common(5)
                        char_html = " ".join([
                            f"<span class='character-tag'>{char} ({count}×)</span>"
                            for char, count in top_chars
                        ])
                        st.markdown("**Characters:** " + char_html, unsafe_allow_html=True)

                with col2:
                    view_key = f"view_{i}"
                    if st.button("👁️ View", key=view_key):
                        st.session_state[f"expanded_{i}"] = not st.session_state.get(f"expanded_{i}", False)

                # Expanded view
                if st.session_state.get(f"expanded_{i}", False):
                    st.markdown("---")
                    st.markdown("### 💬 Messages")

                    for msg in conv["messages"]:
                        role_class = "user-message" if msg["role"] == "user" else "assistant-message"
                        timestamp = datetime.fromtimestamp(msg["timestamp"]).strftime("%H:%M:%S") if msg.get("timestamp") else ""

                        st.markdown(f"""
                        <div class='{role_class}'>
                            <strong>{msg["role"].upper()}</strong> {timestamp}<br>
                            {msg["content"]}
                        </div>
                        """, unsafe_allow_html=True)

                    # Export buttons
                    st.markdown("---")
                    col1, col2 = st.columns(2)

                    with col1:
                        word_doc = export_to_word(conv)
                        st.download_button(
                            label="📄 Download as Word",
                            data=word_doc,
                            file_name=f"{conv['title'][:50]}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key=f"word_{i}"
                        )

                    with col2:
                        pdf_doc = export_to_pdf(conv)
                        st.download_button(
                            label="📑 Download as PDF",
                            data=pdf_doc,
                            file_name=f"{conv['title'][:50]}.pdf",
                            mime="application/pdf",
                            key=f"pdf_{i}"
                        )

                st.markdown("</div>", unsafe_allow_html=True)
                st.markdown("")

elif page == "ℹ️ About":
    st.markdown("# 🌙 About")
    st.markdown("<div class='divider'></div>", unsafe_allow_html=True)

    st.markdown("""
    ## Conversation Archive Browser

    A tool for browsing and analyzing exported conversation archives.

    ### Features
    - 📦 Upload and parse archive files
    - 🔍 Advanced search and filtering
    - 🎨 Theme detection and categorization
    - 👤 Character mention tracking
    - 📊 Analytics and insights
    - 📄 Export to Word and PDF

    ### How to Use
    1. Export your conversations from the source platform
    2. Upload the ZIP or JSON file
    3. Browse, search, and filter conversations
    4. View detailed message threads
    5. Export conversations as needed

    ### Privacy
    All processing happens locally in your browser. No data is uploaded to external servers.
    """)
