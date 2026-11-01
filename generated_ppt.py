import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE

def create_hair_beauty_presentation():
    """
    Generates a 3-slide PowerPoint presentation about hair and beauty,
    covering its definition, importance, and maintenance.
    """
    prs = Presentation()

    # Define a consistent color palette
    DARK_BLUE = RGBColor(31, 73, 125)
    MEDIUM_GREY = RGBColor(89, 89, 89)
    LIGHT_GREY = RGBColor(127, 127, 127)
    ACCENT_COLOR = RGBColor(192, 80, 77) # A subtle reddish accent

    # --- Slide 1: Title Slide ---
    # Layout 0 is typically the Title Slide
    slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(slide_layout)

    title = slide.shapes.title
    subtitle = slide.placeholders[1]

    title.text = "The Essence of Hair & Beauty"
    title.text_frame.paragraphs[0].font.name = 'Arial'
    title.text_frame.paragraphs[0].font.size = Pt(48)
    title.text_frame.paragraphs[0].font.color.rgb = DARK_BLUE
    title.text_frame.paragraphs[0].font.bold = True

    subtitle.text = "A Guide to Understanding, Care, and Maintenance"
    subtitle.text_frame.paragraphs[0].font.name = 'Arial'
    subtitle.text_frame.paragraphs[0].font.size = Pt(28)
    subtitle.text_frame.paragraphs[0].font.color.rgb = MEDIUM_GREY
    subtitle.top = Inches(4.5) # Adjust position for better spacing

    # Add a subtle footer
    left = Inches(0.5)
    top = Inches(6.5)
    width = Inches(9)
    height = Inches(0.5)
    footer = slide.shapes.add_textbox(left, top, width, height)
    footer.text_frame.text = "Empowering Your Natural Radiance"
    footer.text_frame.paragraphs[0].font.name = 'Arial'
    footer.text_frame.paragraphs[0].font.size = Pt(14)
    footer.text_frame.paragraphs[0].font.color.rgb = LIGHT_GREY
    footer.text_frame.paragraphs[0].alignment = MSO_ANCHOR.BOTTOM


    # --- Slide 2: What is Hair & Beauty? Why it's Important ---
    # Layout 1 is typically Title and Content
    slide_layout = prs.slide_layouts[1]
    slide = prs.slides.add_slide(slide_layout)

    title = slide.shapes.title
    content_placeholder = slide.placeholders[1]

    title.text = "Understanding Hair & Beauty: Beyond the Surface"
    title.text_frame.paragraphs[0].font.name = 'Arial'
    title.text_frame.paragraphs[0].font.size = Pt(36)
    title.text_frame.paragraphs[0].font.color.rgb = DARK_BLUE
    title.text_frame.paragraphs[0].font.bold = True

    tf = content_placeholder.text_frame
    tf.clear() # Clear existing text

    # What is Hair?
    p = tf.add_paragraph()
    p.text = "What is Hair?"
    p.font.name = 'Arial'
    p.font.size = Pt(24)
    p.font.color.rgb = ACCENT_COLOR
    p.font.bold = True

    p = tf.add_paragraph()
    p.text = "Biological structure (keratin, follicles) with protective and sensory functions."
    p.level = 1
    p.font.name = 'Arial'
    p.font.size = Pt(20)
    p.font.color.rgb = MEDIUM_GREY

    # What is Beauty?
    p = tf.add_paragraph()
    p.text = "What is Beauty?"
    p.font.name = 'Arial'
    p.font.size = Pt(24)
    p.font.color.rgb = ACCENT_COLOR
    p.font.bold = True

    p = tf.add_paragraph()
    p.text = "A holistic concept encompassing healthy skin, vibrant hair, and overall well-being."
    p.level = 1
    p.font.name = 'Arial'
    p.font.size = Pt(20)
    p.font.color.rgb = MEDIUM_GREY

    p = tf.add_paragraph()
    p.text = "It's about self-expression, confidence, and personal care."
    p.level = 1
    p.font.name = 'Arial'
    p.font.size = Pt(20)
    p.font.color.rgb = MEDIUM_GREY

    # Why is it Important?
    p = tf.add_paragraph()
    p.text = "Why is it Important?"
    p.font.name = 'Arial'
    p.font.size = Pt(24)
    p.font.color.rgb = ACCENT_COLOR
    p.font.bold = True

    p = tf.add_paragraph()
    p.text = "Health Indicator: Reflects internal health and nutrition."
    p.level = 1
    p.font.name = 'Arial'
    p.font.size = Pt(20)
    p.font.color.rgb = MEDIUM_GREY

    p = tf.add_paragraph()
    p.text = "Self-Esteem & Confidence: Boosts mood and self-perception."
    p.level = 1
    p.font.name = 'Arial'
    p.font.size = Pt(20)
    p.font.color.rgb = MEDIUM_GREY

    p = tf.add_paragraph()
    p.text = "Protection: Hair and skin shield against environmental factors."
    p.level = 1
    p.font.name = 'Arial'
    p.font.size = Pt(20)
    p.font.color.rgb = MEDIUM_GREY


    # --- Slide 3: How to Maintain Hair & Beauty ---
    slide_layout = prs.slide_layouts[1]
    slide = prs.slides.add_slide(slide_layout)

    title = slide.shapes.title
    content_placeholder = slide.placeholders[1]

    title.text = "Nurturing Your Radiance: Essential Maintenance Tips"
    title.text_frame.paragraphs[0].font.name = 'Arial'
    title.text_frame.paragraphs[0].font.size = Pt(36)
    title.text_frame.paragraphs[0].font.color.rgb = DARK_BLUE
    title.text_frame.paragraphs[0].font.bold = True

    tf = content_placeholder.text_frame
    tf.clear() # Clear existing text

    # Hair Care Fundamentals
    p = tf.add_paragraph()
    p.text = "Hair Care Fundamentals:"
    p.font.name = 'Arial'
    p.font.size = Pt(24)
    p.font.color.rgb = ACCENT_COLOR
    p.font.bold = True

    p = tf.add_paragraph()
    p.text = "Gentle cleansing, conditioning, and minimizing heat styling."
    p.level = 1
    p.font.name = 'Arial'
    p.font.size = Pt(20)
    p.font.color.rgb = MEDIUM_GREY

    p = tf.add_paragraph()
    p.text = "Regular trims to prevent split ends and maintain health."
    p.level = 1
    p.font.name = 'Arial'
    p.font.size = Pt(20)
    p.font.color.rgb = MEDIUM_GREY

    # Skin Care Essentials
    p = tf.add_paragraph()
    p.text = "Skin Care Essentials:"
    p.font.name = 'Arial'
    p.font.size = Pt(24)
    p.font.color.rgb = ACCENT_COLOR
    p.font.bold = True

    p = tf.add_paragraph()
    p.text = "Daily cleansing, moisturizing, and crucial sun protection (SPF)."
    p.level = 1
    p.font.name = 'Arial'
    p.font.size = Pt(20)
    p.font.color.rgb = MEDIUM_GREY

    p = tf.add_paragraph()
    p.text = "Stay hydrated and maintain a balanced, nutrient-rich diet."
    p.level = 1
    p.font.name = 'Arial'
    p.font.size = Pt(20)
    p.font.color.rgb = MEDIUM_GREY

    # Holistic Well-being
    p = tf.add_paragraph()
    p.text = "Holistic Well-being:"
    p.font.name = 'Arial'
    p.font.size = Pt(24)
    p.font.color.rgb = ACCENT_COLOR
    p.font.bold = True

    p = tf.add_paragraph()
    p.text = "Adequate sleep, stress management, and regular exercise."
    p.level = 1
    p.font.name = 'Arial'
    p.font.size = Pt(20)
    p.font.color.rgb = MEDIUM_GREY

    p = tf.add_paragraph()
    p.text = "Consult professionals (stylists, dermatologists) for personalized advice."
    p.level = 1
    p.font.name = 'Arial'
    p.font.size = Pt(20)
    p.font.color.rgb = MEDIUM_GREY

    # Save the presentation
    output_filename = "presentation.pptx"
    prs.save(output_filename)
    print(f"Presentation '{output_filename}' created successfully!")

if __name__ == "__main__":
    create_hair_beauty_presentation()