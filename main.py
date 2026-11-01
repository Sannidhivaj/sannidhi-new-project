import streamlit as st
import os
import json
import re
import requests
from io import BytesIO
from dotenv import load_dotenv
from langchain_google_genai import ChatGoogleGenerativeAI
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor

# --------------------------------------------------
# 1. CONFIGURATION
# --------------------------------------------------

load_dotenv()

API_KEY = os.getenv("GEM") or st.secrets.get("GEM")

st.set_page_config(
    page_title="Sannidhi's PPT Generator",
    page_icon="🎀",
    layout="centered"
)

# --------------------------------------------------
# 2. COLOR THEMES
# --------------------------------------------------

THEMES = {
    "Pink": {
        "background": "FFF0F5",
        "primary": "C2185B",
        "secondary": "F8BBD0",
        "accent": "E91E63",
        "text": "5A1735",
        "card": "FFFFFF"
    },
    "Blue": {
        "background": "EFF6FF",
        "primary": "1D4ED8",
        "secondary": "BFDBFE",
        "accent": "38BDF8",
        "text": "172554",
        "card": "FFFFFF"
    },
    "Purple": {
        "background": "F5F0FF",
        "primary": "6D28D9",
        "secondary": "DDD6FE",
        "accent": "A78BFA",
        "text": "3B0764",
        "card": "FFFFFF"
    },
    "Professional": {
        "background": "F3F4F6",
        "primary": "26364A",
        "secondary": "CBD5E1",
        "accent": "0F766E",
        "text": "1F2937",
        "card": "FFFFFF"
    }
}


def rgb(hex_color):
    """Convert a hexadecimal color into a PowerPoint RGB color."""
    hex_color = hex_color.lstrip("#")
    return RGBColor(
        int(hex_color[0:2], 16),
        int(hex_color[2:4], 16),
        int(hex_color[4:6], 16)
    )


# --------------------------------------------------
# 3. CUSTOM STREAMLIT DESIGN
# --------------------------------------------------

st.markdown(
    """
    <style>
    .stApp {
        background-color: #FFF0F5;
    }

    .main .block-container {
        max-width: 1050px;
        padding-top: 3rem;
        padding-bottom: 3rem;
    }

    h1 {
        color: #C2185B !important;
        font-family: "Lucida Bright", Georgia, serif !important;
        font-weight: bold !important;
        text-align: center;
        font-size: 2.8rem !important;
    }

    h2, h3 {
        color: #C2185B !important;
        font-family: "Lucida Bright", Georgia, serif !important;
    }

    .stTextArea label,
    .stTextArea label p,
    .stSelectbox label,
    .stSelectbox label p {
        color: #C2185B !important;
        font-family: "Lucida Bright", Georgia, serif !important;
        font-weight: bold !important;
        font-size: 1.1rem !important;
    }

    .stTextArea textarea {
        background-color: #FFFFFF !important;
        color: #5A1735 !important;
        border: 2px solid #F2B6CE !important;
        border-radius: 14px !important;
        font-family: Georgia, serif !important;
    }

    .stTextArea textarea:focus {
        border-color: #C2185B !important;
        box-shadow: 0 0 0 2px #F8BBD0 !important;
    }

    .stSelectbox div[data-baseweb="select"] > div {
        background-color: #FFFFFF;
        border-color: #F2B6CE;
        border-radius: 10px;
    }

    .stButton button,
    .stDownloadButton button {
        background-color: #C2185B !important;
        color: #FFFFFF !important;
        border: 1px solid #C2185B !important;
        border-radius: 10px !important;
        padding: 0.6rem 1.3rem !important;
        font-family: "Lucida Bright", Georgia, serif !important;
        font-weight: bold !important;
    }

    .stButton button:hover,
    .stDownloadButton button:hover {
        background-color: #A3154B !important;
        border-color: #A3154B !important;
    }

    div[data-testid="stAlert"] {
        border-radius: 12px;
    }

    header[data-testid="stHeader"] {
        background-color: transparent;
    }
    </style>
    """,
    unsafe_allow_html=True
)


# --------------------------------------------------
# 4. SEARCH FOR RELEVANT IMAGES
# --------------------------------------------------

@st.cache_data(show_spinner=False, ttl=3600)
def search_wikimedia_image(query):
    """
    Search Wikimedia Commons for an image related to the topic.
    Returns image bytes, or None if no usable image is found.
    """

    headers = {
        "User-Agent": "SannidhiPPTGenerator/1.0 (educational project)"
    }

    try:
        search_response = requests.get(
            "https://commons.wikimedia.org/w/api.php",
            params={
                "action": "query",
                "list": "search",
                "srsearch": f"filetype:bitmap {query}",
                "srnamespace": 6,
                "srlimit": 5,
                "format": "json"
            },
            headers=headers,
            timeout=15
        )

        search_response.raise_for_status()

        search_results = search_response.json().get(
            "query", {}
        ).get("search", [])

        for result in search_results:
            file_title = result.get("title")

            if not file_title:
                continue

            image_response = requests.get(
                "https://commons.wikimedia.org/w/api.php",
                params={
                    "action": "query",
                    "titles": file_title,
                    "prop": "imageinfo",
                    "iiprop": "url",
                    "iiurlwidth": 1200,
                    "format": "json"
                },
                headers=headers,
                timeout=15
            )

            image_response.raise_for_status()

            pages = image_response.json().get(
                "query", {}
            ).get("pages", {})

            for page in pages.values():
                image_info = page.get("imageinfo", [])

                if not image_info:
                    continue

                image_url = image_info[0].get(
                    "thumburl",
                    image_info[0].get("url")
                )

                if not image_url:
                    continue

                download_response = requests.get(
                    image_url,
                    headers=headers,
                    timeout=20
                )

                download_response.raise_for_status()

                image_bytes = download_response.content

                if (
                    download_response.headers.get(
                        "Content-Type", ""
                    ).startswith("image/")
                    and image_bytes
                ):
                    return image_bytes

    except (requests.RequestException, ValueError, KeyError):
        return None

    return None


# --------------------------------------------------
# 5. POWERPOINT HELPER FUNCTIONS
# --------------------------------------------------

def add_text(
    slide,
    text,
    x,
    y,
    width,
    height,
    font_size,
    color,
    bold=False,
    font_name="Georgia",
    alignment=PP_ALIGN.LEFT
):
    """Add a formatted text box to a slide."""

    box = slide.shapes.add_textbox(
        Inches(x),
        Inches(y),
        Inches(width),
        Inches(height)
    )

    text_frame = box.text_frame
    text_frame.clear()
    text_frame.word_wrap = True
    text_frame.margin_left = Inches(0.08)
    text_frame.margin_right = Inches(0.08)
    text_frame.margin_top = Inches(0.05)
    text_frame.margin_bottom = Inches(0.05)

    paragraph = text_frame.paragraphs[0]
    paragraph.alignment = alignment
    paragraph.space_after = Pt(6)

    run = paragraph.add_run()
    run.text = str(text)
    run.font.name = font_name
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = rgb(color)

    return box


def set_slide_background(slide, color):
    """Set the slide background color."""

    fill = slide.background.fill
    fill.solid()
    fill.fore_color.rgb = rgb(color)


def add_top_accent(slide, theme):
    """Add a decorative strip at the top of the slide."""

    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(0),
        Inches(0),
        Inches(13.333),
        Inches(0.13)
    )

    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb(theme["accent"])
    shape.line.fill.background()


def add_slide_number(slide, number, theme):
    """Add a slide number in the footer."""

    add_text(
        slide,
        f"{number:02d}",
        12.1,
        7.02,
        0.6,
        0.25,
        10,
        theme["primary"],
        bold=True,
        alignment=PP_ALIGN.RIGHT
    )


def add_picture_if_available(slide, image_bytes, x, y, w, h):
    """Insert an image when valid image bytes are available."""

    if not image_bytes:
        return False

    try:
        picture_stream = BytesIO(image_bytes)

        slide.shapes.add_picture(
            picture_stream,
            Inches(x),
            Inches(y),
            width=Inches(w),
            height=Inches(h)
        )

        return True

    except (ValueError, TypeError, OSError):
        return False


# --------------------------------------------------
# 6. CREATE THE PRESENTATION
# --------------------------------------------------

def create_presentation(data, theme_name):
    """Build and return a themed PowerPoint presentation."""

    theme = THEMES[theme_name]

    prs = Presentation()

    # Widescreen 16:9 presentation
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    blank_layout = prs.slide_layouts[6]

    title = str(data.get("presentation_title", "My Presentation"))
    subtitle = str(data.get("subtitle", "An overview of the topic"))
    slides_data = data.get("slides", [])

    if not isinstance(slides_data, list):
        raise ValueError("Gemini returned an invalid slides structure.")

    # ---------------- TITLE SLIDE ----------------

    slide = prs.slides.add_slide(blank_layout)
    set_slide_background(slide, theme["background"])
    add_top_accent(slide, theme)

    # Decorative panel
    panel = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE,
        Inches(0.55),
        Inches(1.0),
        Inches(7.0),
        Inches(5.7)
    )

    panel.fill.solid()
    panel.fill.fore_color.rgb = rgb(theme["card"])
    panel.line.color.rgb = rgb(theme["secondary"])

    add_text(
        slide,
        "YOUR PRESENTATION",
        1.0,
        1.65,
        5.9,
        0.4,
        13,
        theme["accent"],
        bold=True
    )

    add_text(
        slide,
        title,
        1.0,
        2.2,
        5.9,
        2.0,
        30,
        theme["primary"],
        bold=True
    )

    add_text(
        slide,
        subtitle,
        1.0,
        4.55,
        5.9,
        1.1,
        17,
        theme["text"]
    )

    add_text(
        slide,
        "Created with Sannidhi's PPT Generator",
        1.0,
        6.1,
        5.9,
        0.35,
        11,
        theme["primary"]
    )

    cover_query = data.get(
        "cover_image_query",
        title
    )

    cover_image = search_wikimedia_image(cover_query)

    add_picture_if_available(
        slide,
        cover_image,
        8.0,
        1.35,
        4.65,
        4.9
    )

    # ---------------- CONTENT SLIDES ----------------

    for index, item in enumerate(slides_data, start=2):

        if not isinstance(item, dict):
            continue

        slide = prs.slides.add_slide(blank_layout)
        set_slide_background(slide, theme["background"])
        add_top_accent(slide, theme)

        slide_title = str(
            item.get("title", f"Section {index - 1}")
        )

        bullets = item.get("bullets", [])

        if not isinstance(bullets, list):
            bullets = [str(bullets)]

        image_query = str(
            item.get("image_search_query", slide_title)
        )

        # Download a relevant image for this slide
        image_bytes = search_wikimedia_image(image_query)

        # Text uses more space if the image is unavailable
        text_width = 7.1 if image_bytes else 11.5

        add_text(
            slide,
            slide_title,
            0.65,
            0.55,
            12.0,
            0.8,
            25,
            theme["primary"],
            bold=True
        )

        # Divider under the slide title
        divider = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE,
            Inches(0.75),
            Inches(1.45),
            Inches(1.3),
            Inches(0.055)
        )

        divider.fill.solid()
        divider.fill.fore_color.rgb = rgb(theme["accent"])
        divider.line.fill.background()

        # Bullet content
        content_box = slide.shapes.add_textbox(
            Inches(0.7),
            Inches(1.8),
            Inches(text_width),
            Inches(4.9)
        )

        text_frame = content_box.text_frame
        text_frame.clear()
        text_frame.word_wrap = True
        text_frame.margin_left = Inches(0.12)
        text_frame.margin_right = Inches(0.12)
        text_frame.margin_top = Inches(0.12)
        text_frame.margin_bottom = Inches(0.12)

        if not bullets:
            bullets = [
                "Explore the main concepts and key takeaways.",
                "Understand the practical importance of this topic."
            ]

        for bullet_index, bullet in enumerate(bullets):
            paragraph = (
                text_frame.paragraphs[0]
                if bullet_index == 0
                else text_frame.add_paragraph()
            )

            paragraph.text = f"•  {str(bullet)}"
            paragraph.font.name = "Georgia"
            paragraph.font.size = Pt(17)
            paragraph.font.color.rgb = rgb(theme["text"])
            paragraph.space_after = Pt(17)
            paragraph.level = 0

        # Image panel on the right
        if image_bytes:
            image_panel = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE,
                Inches(8.15),
                Inches(1.8),
                Inches(4.55),
                Inches(4.65)
            )

            image_panel.fill.solid()
            image_panel.fill.fore_color.rgb = rgb(theme["card"])
            image_panel.line.color.rgb = rgb(theme["secondary"])

            add_picture_if_available(
                slide,
                image_bytes,
                8.28,
                1.93,
                4.29,
                4.39
            )

        add_slide_number(slide, index, theme)

    # Return the presentation as bytes
    output = BytesIO()
    prs.save(output)
    output.seek(0)

    return output.getvalue()


# --------------------------------------------------
# 7. STREAMLIT USER INTERFACE
# --------------------------------------------------

st.title("Sannidhi's PPT Generator")

st.markdown(
    """
    <p style="
        text-align:center;
        color:#8E174A;
        font-family:Georgia,serif;
        font-size:17px;
        margin-top:-10px;
        margin-bottom:30px;">
        Turn your ideas into beautiful presentations. you can make one🎀
    </p>
    """,
    unsafe_allow_html=True
)

st.text_area(
    "Enter your PPT topic",
    key="ppt_topic",
    placeholder=(
        "Example: Create a 3 or more slides presentation about any app or product or a topic "
        "including its introduction, uses, and features."
    ),
    height=150
)

theme_name = st.selectbox(
    "Choose your slide theme",
    options=list(THEMES.keys()),
    index=0,
    help="Choose the colors used throughout your presentation."
)

# Show a preview of the selected theme
selected_theme = THEMES[theme_name]

st.markdown(
    f"""
    <div style="
        display:flex;
        align-items:center;
        gap:12px;
        padding:14px;
        background:#{selected_theme['card']};
        border:1px solid #{selected_theme['secondary']};
        border-radius:12px;
        margin:8px 0 20px 0;">
        <div style="
            width:35px;height:35px;border-radius:8px;
            background:#{selected_theme['background']};
            border:1px solid #{selected_theme['secondary']};">
        </div>
        <div style="
            width:35px;height:35px;border-radius:8px;
            background:#{selected_theme['primary']};">
        </div>
        <div style="
            width:35px;height:35px;border-radius:8px;
            background:#{selected_theme['accent']};">
        </div>
        <span style="
            color:#{selected_theme['text']};
            font-family:Georgia,serif;">
            Preview: {theme_name} theme
        </span>
    </div>
    """,
    unsafe_allow_html=True
)

if st.button("🎀 Generate my PPT", use_container_width=True):

    topic = st.session_state.ppt_topic.strip()

    if not topic:
        st.warning("Please enter a topic first.")

    elif not API_KEY:
        st.error(
            "Gemini API key not found. Add GEM=your_api_key "
            "to your .env file."
        )

    else:
        try:
            with st.spinner(
                "Creating your slides and finding relevant images... 💗"
            ):

                model = ChatGoogleGenerativeAI(
                    model="gemini-2.5-flash",
                    google_api_key=API_KEY,
                    temperature=0.4
                )

                system_prompt = f"""
You are a professional presentation writer.

Create detailed, accurate, well-organized PowerPoint slide content
for the user's requested topic.

The chosen visual theme is: {theme_name}.
The presentation will be designed in Python using python-pptx.

Return ONLY valid JSON. Do not use Markdown fences or introductory text.

Use exactly this JSON structure:

{{
  "presentation_title": "A clear presentation title",
  "subtitle": "A short descriptive subtitle",
  "cover_image_query": "A specific visual search phrase for the topic",
  "slides": [
    {{
      "title": "Slide heading",
      "bullets": [
        "A specific, informative point",
        "Another useful and accurate point",
        "A practical example or explanation"
      ],
      "image_search_query": "A specific phrase describing an image relevant to this slide"
    }}
  ]
}}

Requirements:
1. Create a title and approximately 5 content slides unless the user
   requests a different number.
2. Cover the topic in a logical sequence, with an introduction,
   essential concepts, practical uses or examples, and a conclusion.
3. Write meaningful, accurate, topic-specific content.
4. Each content slide should contain 3 to 5 concise but informative
   bullet points.
5. Give each slide a distinct, relevant image_search_query.
6. Image queries must describe the actual subject of the slide,
   not generic phrases such as "business presentation".
7. Do not invent facts, statistics, quotations, or sources.
8. Return valid JSON with double-quoted keys and string values.
9. Do not include Python code in the response.
"""

                response = model.invoke([
                    ("system", system_prompt),
                    ("user", topic)
                ])

                response_content = response.content

                if isinstance(response_content, list):
                    response_content = "".join(
                        block.get("text", "")
                        if isinstance(block, dict)
                        else str(block)
                        for block in response_content
                    )

                response_content = response_content.strip()

                # Remove accidental Markdown code fences
                response_content = re.sub(
                    r"^```(?:json)?\s*",
                    "",
                    response_content
                )

                response_content = re.sub(
                    r"\s*```$",
                    "",
                    response_content
                )

                data = json.loads(response_content)

                if not isinstance(data, dict) or not isinstance(
                    data.get("slides"), list
                ):
                    raise ValueError(
                        "Gemini returned an unexpected presentation format. "
                        "Please try again."
                    )

                if not data["slides"]:
                    raise ValueError(
                        "No content slides were returned. Please try again."
                    )

                # Build the actual PPTX file
                pptx_bytes = create_presentation(
                    data,
                    theme_name
                )

                st.session_state["generated_ppt_bytes"] = pptx_bytes
                st.session_state["generated_ppt_name"] = (
                    re.sub(
                        r"[^a-zA-Z0-9_-]+",
                        "_",
                        str(data.get("presentation_title", "presentation"))
                    ).strip("_")[:70] or "presentation"
                ) + ".pptx"

                st.session_state["generated_ppt_theme"] = theme_name

        except json.JSONDecodeError:
            st.error(
                "Gemini did not return valid JSON. Please click Generate "
                "my PPT again."
            )

        except Exception as error:
            st.error(f"Something went wrong: {error}")

# Display download button outside the generation block so it remains
# available when Streamlit reruns after a download or other interaction.
if st.session_state.get("generated_ppt_bytes"):

    st.success(
        f"Your {st.session_state.get('generated_ppt_theme', '')} "
        "presentation is ready!"
    )

    st.download_button(
        label="Download your PowerPoint",
        data=st.session_state["generated_ppt_bytes"],
        file_name=st.session_state["generated_ppt_name"],
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "presentationml.presentation"
        ),
        use_container_width=True
    )

st.caption(
    "Images are retrieved from Wikimedia Commons. "
    "Internet access is required, and image availability may vary."
)
