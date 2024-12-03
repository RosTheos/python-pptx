from pptx import Presentation

# Load the uploaded presentation
input_presentation_path = "/mnt/data/Presentation.pptx"
input_presentation = Presentation(input_presentation_path)

# Create a new PowerPoint presentation to outline improvements
output_presentation = Presentation()

# Title slide for the feedback presentation
title_slide_layout = output_presentation.slide_layouts[0]
slide = output_presentation.slides.add_slide(title_slide_layout)
title = slide.shapes.title
subtitle = slide.placeholders[1]

title.text = "Feedback for Keyence BT-A700 Presentation"
subtitle.text = "Suggested Improvements and Additions"

# Slide 1: Strengths
bullet_slide_layout = output_presentation.slide_layouts[1]
slide = output_presentation.slides.add_slide(bullet_slide_layout)
shapes = slide.shapes
title = shapes.title
content = shapes.placeholders[1]

title.text = "Strengths"
content.text = (
    "- Clear problem-solution structure.\n"
    "- Detailed benefits of BT-A700 highlighted.\n"
    "- Effective competitive comparison.\n"
    "- Customer testimonials add credibility."
)

# Slide 2: Improvements Overview
slide = output_presentation.slides.add_slide(bullet_slide_layout)
title = slide.shapes.title
content = slide.placeholders[1]

title.text = "Overview of Improvements"
content.text = (
    "- Refine introduction to highlight Keyence's unique value.\n"
    "- Add more visuals and reduce text density.\n"
    "- Provide specific ROI metrics and integration details.\n"
    "- Enhance call-to-action clarity."
)

# Slide 3: Detailed Suggestions - Visuals and Content
slide = output_presentation.slides.add_slide(bullet_slide_layout)
title = slide.shapes.title
content = slide.placeholders[1]

title.text = "Visuals and Content Suggestions"
content.text = (
    "- Use diagrams showing warehouse integration.\n"
    "- Include photos/videos of the BT-A700 in action.\n"
    "- Add graphs or charts demonstrating performance improvements.\n"
    "- Show examples of durability (IP67, drop resistance)."
)

# Slide 4: Technical Details Simplification
slide = output_presentation.slides.add_slide(bullet_slide_layout)
title = slide.shapes.title
content = slide.placeholders[1]

title.text = "Simplify Technical Details"
content.text = (
    "- Break down technical specs into bullet points.\n"
    "- Use simpler language for complex features.\n"
    "- Highlight practical benefits for Blue Wagon specifically."
)

# Slide 5: Call-to-Action Enhancements
slide = output_presentation.slides.add_slide(bullet_slide_layout)
title = slide.shapes.title
content = slide.placeholders[1]

title.text = "Call-to-Action Improvements"
content.text = (
    "- Add QR codes or links for catalog/demo requests.\n"
    "- Include clear contact details (phone, email, website).\n"
    "- Provide pricing or cost-benefit analysis options."
)

# Slide 6: Anticipated Questions
slide = output_presentation.slides.add_slide(bullet_slide_layout)
title = slide.shapes.title
content = slide.placeholders[1]

title.text = "Anticipated Questions"
content.text = (
    "- Pricing details and cost-effectiveness vs. competitors.\n"
    "- Integration process with existing systems.\n"
    "- Scalability for future growth."
)

# Save the feedback presentation
output_presentation_path = "/mnt/data/Keyence_BT-A700_Feedback.pptx"
output_presentation.save(output_presentation_path)

output_presentation_path
