from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# Create a presentation object
prs = Presentation()

# Slide layouts
title_slide_layout = prs.slide_layouts[0]  # Title Slide
content_slide_layout = prs.slide_layouts[1]  # Title and Content

# Function to add a slide with title, content, and notes
def add_slide(title, content_bullets, background_color=None, notes_text=""):
    slide = prs.slides.add_slide(content_slide_layout)
    shapes = slide.shapes

    # Set background color if provided
    if background_color:
        fill = slide.background.fill
        fill.solid()
        fill.fore_color.rgb = background_color

    # Title
    title_placeholder = shapes.title
    title_placeholder.text = title
    title_placeholder.text_frame.paragraphs[0].font.size = Pt(44)
    title_placeholder.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 0, 0)  # Red for festive

    # Content
    content_placeholder = shapes.placeholders[1]
    tf = content_placeholder.text_frame
    tf.text = ""
    for bullet in content_bullets:
        p = tf.add_paragraph()
        p.text = bullet
        p.font.size = Pt(24)
        p.font.color.rgb = RGBColor(0, 128, 0)  # Green
        p.level = 0

    # Add notes
    notes_slide = slide.notes_slide
    notes_slide.notes_text_frame.text = notes_text

# Title Slide
title_slide = prs.slides.add_slide(title_slide_layout)
title = title_slide.shapes.title
title.text = "Magical Christmas Party at Church"
subtitle = title_slide.placeholders[1]
subtitle.text = "A Festive Adventure for Kids and Families!\nGet ready for fun, crafts, and holiday magic!"
title.text_frame.paragraphs[0].font.size = Pt(48)
title.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 215, 0)  # Gold
subtitle.text_frame.paragraphs[0].font.size = Pt(32)
subtitle.text_frame.paragraphs[0].font.color.rgb = RGBColor(0, 128, 0)

# Notes for title slide
title_slide.notes_slide.notes_text_frame.text = "Animation: Fade in title with snowflake effects. Background: Red and green gradient."

# Introduction Slide
add_slide(
    "Welcome to the Party!",
    [
        "Ho ho ho! Welcome to our magical Christmas party!",
        "We'll have crafts, stories, films, and gifts galore.",
        "Parents: Join the WhatsApp group for updates and fun ideas.",
        "Kids: Get your passports ready for adventure!"
    ],
    background_color=RGBColor(0, 128, 0),  # Green
    notes_text="Animation: Bounce in bullets. Image idea: Santa waving. Interactivity: Click to start."
)

# Slide 1: Registration for Kids
add_slide(
    "Step 1: Registration for Kids",
    [
        "While buying tickets, register your little elves!",
        "Prepare their 'passports' – custom ID cards for the event.",
        "Fun pun: It's their passport to fun and frolic!",
        "Volunteers: Collect names, ages, and favorite colors."
    ],
    background_color=RGBColor(255, 0, 0),  # Red
    notes_text="Animation: Slide in from left. Image idea: Kids at a ticket booth with Santa hats."
)

# Slide 2: WhatsApp Group for Parents
add_slide(
    "Step 2: WhatsApp Group for Parents",
    [
        "Create a group to promote the event and share ideas.",
        "Remind about gift-sharing and invite friends!",
        "Share holiday recipes or craft tips.",
        "Parents: Stay connected for last-minute updates."
    ],
    background_color=RGBColor(255, 215, 0),  # Gold
    notes_text="Animation: Fade in text. Image idea: Phones with holiday emojis. Interactivity: Link to join group."
)

# Slide 3: Passport Distribution
add_slide(
    "Step 3: Passport Distribution",
    [
        "On party day, kids get their magical passports!",
        "Hand them out upon arrival – no elf left behind.",
        "Fun pun: 'Passport to Christmas Cheer!'",
        "Volunteers: Stamp them with holiday stickers."
    ],
    background_color=RGBColor(0, 128, 0),  # Green
    notes_text="Animation: Zoom in. Image idea: Kids receiving cards with reindeer stamps."
)

# Slide 4: Slot Division and Story Narration
add_slide(
    "Step 4: Slot Division and Story Narration",
    [
        "Divide kids into slots of 10 for organized fun.",
        "Narrate the Saint Nicolas story to each group.",
        "Bring the tale to life with voices and gestures!",
        "Volunteers: Be the storytellers – wear Santa hats!"
    ],
    background_color=RGBColor(255, 0, 0),  # Red
    notes_text="Animation: Story text appears word by word. Image idea: Group of kids listening to Santa."
)

# Slide 5: Art and Crafts Stations
add_slide(
    "Step 5: Art and Crafts Stations",
    [
        "Guide kids to 4 stations, each 10-15 minutes.",
        "Creative activities: Make ornaments, cards, and more!",
        "Fun pun: 'Crafty elves at work!'",
        "Volunteers: Supervise and provide supplies."
    ],
    background_color=RGBColor(255, 215, 0),  # Gold
    notes_text="Animation: Rotate in bullets. Image idea: Kids painting snowflakes. Interactivity: Clickable station previews."
)

# Slide 6: Tote Bag and Film Viewing
add_slide(
    "Step 6: Tote Bag and Film Viewing",
    [
        "Kids get a tote bag to hold their crafts.",
        "Then, watch the end of 'Christmas Factory' film.",
        "Relax and enjoy the holiday magic on screen!",
        "Volunteers: Pass out bags and start the movie."
    ],
    background_color=RGBColor(0, 128, 0),  # Green
    notes_text="Animation: Fade in. Image idea: Kids with tote bags watching TV. Embed video if possible."
)

# Slide 7: Gift Wrapping
add_slide(
    "Step 7: Gift Wrapping",
    [
        "Kids wrap the gifts they brought.",
        "Stored safely in the 'kasha room' (gift area).",
        "Fun pun: 'Wrap it up with love!'",
        "Volunteers: Help with wrapping paper and tape."
    ],
    background_color=RGBColor(255, 0, 0),  # Red
    notes_text="Animation: Slide up. Image idea: Kids wrapping presents under a tree."
)

# Conclusion Slide
add_slide(
    "Thank You and Merry Christmas!",
    [
        "Thanks for joining our festive celebration!",
        "We hope you had a magical time.",
        "Stay tuned for more church events.",
        "Merry Christmas to all, and to all a good night!"
    ],
    background_color=RGBColor(255, 215, 0),  # Gold
    notes_text="Animation: Sparkle effects. Image idea: Group photo with Santa. Interactivity: Link to feedback form."
)

# Save the presentation
prs.save('christmas_party_presentation.pptx')