"""
MEP Proposal Generator - Complete Template Version
Kimley-Horn Engineering Services
Generates professional .docx proposals matching exact template format
Version 3.0 - Complete Template Implementation
"""

import streamlit as st
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from datetime import datetime
from io import BytesIO
import re
import os
import base64

# Hardcoded Kimley-Horn logo (base64 encoded)
# This will be embedded directly in every document
KIMLEY_HORN_LOGO_BASE64 = """
iVBORw0KGgoAAAANSUhEUgAAAhwAAACDCAYAAADGUcLDAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAALiMAAC4jAXilP3YAANaTSURBVHhe7P13vCZlff+PP68yM3c5bXfZpYOACNIsQRRUFIlYgoktUWOiJmqSj4nGfPz+YiwxsUSjSUw0zRosgKKoiL1XbChI7yC9bDt7zrnLzFzl98d1XXPmHBZ2ZSHqx/u1j3mcc+6de+aa67ruer+udxXOOc8EE0wwwQQTTDDBfQi5+oMJJphgggkmmGCCexsTwjHBBBNMMMEEE9znmBCOCSaYYIIJJpjgPseEcEwwwQQTTDDBBPc5JoRjggkmmGCCCSa4zzEhHBNMMMEEE0wwwX2OCeGYYIIJJphgggnuc0wIxwQTTDDBBBNMcJ9jQjgmmGCCCSaYYIL7HBPCMcEEE0wwwQQT3OeYEI4JJphgggkmmOA+x4RwTDDBBBNMMMEE9zkmhGOCCSaYYIIJJrjPMSEcE0wwwQQTTDDBfQ4xKU8/wQQTTHDP4QWI1ira/p34/+nn6v9b/beLW8DVn6drrP58V5GuCyC8W/H3Xe1HQxscAF6Ac1a3675q7wS/2pBCCIQIs8N7j/crZ4j3nrqum3OEEBhjALDWrjj3nsA5hxCiuYcxpmlDuufdQQiBtbb5bvtZjDFUtQUhsA4QAmM9CNF85pzDe09VVVRVheqP0Wh0p76Y4B5ACJwX1MY1/e8J/Z8+G5d183+XXX4lnz77s3znu98N4zTBBC14sfIIgi8IP6KA2/4hEF4Aq49dg0fgmnaltjhwDiyAoC4NHoH1gsrUWO+w3uC9DV+0gPXgPQ6HwYK3CBx1XeKweDzjeoxzDlsbjLNYv/zc9wReQI1jaEtqV2O9wXiDw2HjP4/AIzAIDB7rPcbVeCwIj8VSprW7qrE2tMtiqV3d9ImIxwS/3hDGGE8U3EKIhnSkI8synHMYY1BKAZDnOc45rLXkeb5rReHWP0ajEXmeN9eU69Z1jTGGTqfT3FvKwDiXlpaoqug30+l0UCrzY58ShljLsmIymbBlyxaqumQ8GpPnOZ1Oh/F4zGQyodPpMBgM6PV6AAwGA4qiQGsNkbhMcM8xLms6nRznQEooyxppNZs2bmTLli1cccUVfvPmzdx2221cccUVlGUJwMzMDKeccgpFHsZhgglYtSOHsCsP2P5OexlxI3InjnGXX9gpeAQu3lQmweoAG7UDxmOUwGcKBAg
"""

def get_logo_from_base64():
    """Convert base64 logo to BytesIO object for embedding"""
    try:
        logo_data = base64.b64decode(KIMLEY_HORN_LOGO_BASE64.strip())
        return BytesIO(logo_data)
    except:
        return None

# Page config
st.set_page_config(
    page_title="MEP Proposal Generator - Kimley-Horn",
    page_icon="📄",
    layout="wide"
)

# Custom CSS
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(90deg, #8B0000, #C8102E);
        padding: 20px;
        border-radius: 10px;
        margin-bottom: 20px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    .main-header h1 {
        color: white;
        margin: 0;
    }
    .main-header p {
        color: #ffcccc;
        margin: 5px 0 0 0;
    }
</style>
""", unsafe_allow_html=True)

# Header
st.markdown("""
<div class="main-header">
    <h1>📄 MEP Proposal Generator v3.0</h1>
    <p>Kimley-Horn Engineering Services - Complete Template Match</p>
</div>
""", unsafe_allow_html=True)


# ============== HELPER FUNCTIONS ==============

def validate_currency(value):
    """Validate and format currency input"""
    if not value:
        return True, ""
    cleaned = str(value).replace('$', '').replace(',', '').strip()
    try:
        float(cleaned)
        return True, cleaned
    except ValueError:
        return False, None


def validate_email(email):
    """Validate email format"""
    if not email:
        return True
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return re.match(pattern, email) is not None


def format_currency(value):
    """Format number as currency"""
    try:
        num = float(str(value).replace(',', '').replace('$', ''))
        return f"{num:,.0f}"
    except:
        return value


def setup_styles(doc):
    """Setup proper Word styles for the document"""
    # Normal style - 10pt Arial to match template
    normal = doc.styles['Normal']
    normal.font.name = 'Arial'
    normal.font.size = Pt(10)
    normal.paragraph_format.space_after = Pt(0)
    normal.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    
    # Try to get or create List Bullet style
    try:
        bullet_style = doc.styles['List Bullet']
    except KeyError:
        bullet_style = doc.styles.add_style('List Bullet', WD_STYLE_TYPE.PARAGRAPH)
    
    bullet_style.base_style = doc.styles['Normal']
    bullet_style.font.name = 'Arial'
    bullet_style.font.size = Pt(10)
    bullet_style.paragraph_format.left_indent = Inches(0.25)
    bullet_style.paragraph_format.space_after = Pt(0)
    bullet_style.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE


def add_header_with_logo(section, page_num=1):
    """Add header with hardcoded Kimley-Horn logo and page number"""
    header = section.header
    header.is_linked_to_previous = False
    
    p = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    
    # Try to use hardcoded logo image
    logo_stream = get_logo_from_base64()
    if logo_stream:
        try:
            run = p.add_run()
            run.add_picture(logo_stream, width=Inches(2.5))
        except:
            # Fallback to text if image fails
            add_text_logo(p)
    else:
        # Fallback to text logo
        add_text_logo(p)
    
    # Add page number on right side - 11pt Arial, right justified
    if page_num:
        # Add tab stop for right alignment at right margin
        tab_stops = p.paragraph_format.tab_stops
        tab_stops.add_tab_stop(Inches(6.5), WD_TAB_ALIGNMENT.RIGHT)
        p.add_run("\t")
        run_page = p.add_run(f"Page {page_num}")
        run_page.font.size = Pt(11)
        run_page.font.name = 'Arial'
        run_page.font.bold = False


def add_text_logo(paragraph):
    """Add text-based logo as fallback - Arial Narrow"""
    run1 = paragraph.add_run("Kimley")
    run1.font.size = Pt(28)
    run1.font.bold = False
    run1.font.color.rgb = RGBColor(88, 89, 91)
    run1.font.name = 'Arial Narrow'
    
    run2 = paragraph.add_run("»")
    run2.font.size = Pt(28)
    run2.font.bold = False
    run2.font.color.rgb = RGBColor(88, 89, 91)
    run2.font.name = 'Arial Narrow'
    
    run3 = paragraph.add_run("Horn")
    run3.font.size = Pt(28)
    run3.font.bold = False
    run3.font.color.rgb = RGBColor(166, 25, 46)
    run3.font.name = 'Arial Narrow'


def add_footer(section, text_left, text_center, text_right):
    """Add colored footer with exact specs: Col1=1.1", Col2=4.23", Col3=0.96", Height=0.22", with 0.01" white gaps"""
    footer = section.footer
    footer.is_linked_to_previous = False
    
    # Create table with 5 columns (3 content + 2 white gaps of 0.01" each)
    gap_width = Inches(0.01)  # Tiny white space between colored sections
    total_width = Inches(1.1) + gap_width + Inches(4.23) + gap_width + Inches(0.96)
    
    table = footer.add_table(rows=1, cols=5, width=total_width)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    
    # Set exact column widths
    table.columns[0].width = Inches(1.1)      # Grey section
    table.columns[1].width = gap_width        # White gap (0.01")
    table.columns[2].width = Inches(4.23)     # Red/mauve section (address)
    table.columns[3].width = gap_width        # White gap (0.01")
    table.columns[4].width = Inches(0.96)     # Red/mauve section (phone)
    
    cells = table.rows[0].cells
    
    # Set row height to exactly 0.22"
    row = table.rows[0]
    row.height = Inches(0.22)
    
    # Colors
    grey_fill = 'ABABAB'
    red_fill = 'BF8F96'
    
    def set_cell_margins(cell, tight=True):
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        tcMar = OxmlElement('w:tcMar')
        if tight:
            for margin_name in ['top', 'bottom']:
                margin = OxmlElement(f'w:{margin_name}')
                margin.set(qn('w:w'), '20')
                margin.set(qn('w:type'), 'dxa')
                tcMar.append(margin)
            for margin_name in ['left', 'right']:
                margin = OxmlElement(f'w:{margin_name}')
                margin.set(qn('w:w'), '40')
                margin.set(qn('w:type'), 'dxa')
                tcMar.append(margin)
        else:
            for margin_name in ['top', 'bottom', 'left', 'right']:
                margin = OxmlElement(f'w:{margin_name}')
                margin.set(qn('w:w'), '0')
                margin.set(qn('w:type'), 'dxa')
                tcMar.append(margin)
        tcPr.append(tcMar)
    
    def set_cell_color_and_text(cell, text, fill_color):
        """Set cell background color and text"""
        cell.text = text
        
        cell_shading = OxmlElement('w:shd')
        cell_shading.set(qn('w:fill'), fill_color)
        cell._tc.get_or_add_tcPr().append(cell_shading)
        
        tc = cell._tc
        tcPr = tc.get_or_add_tcPr()
        vAlign = OxmlElement('w:vAlign')
        vAlign.set(qn('w:val'), 'center')
        tcPr.append(vAlign)
        
        for paragraph in cell.paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(0)
            paragraph.paragraph_format.line_spacing = 1.0
            for run in paragraph.runs:
                run.font.size = Pt(8)
                run.font.color.rgb = RGBColor(255, 255, 255)
                run.font.name = 'Arial'
                run.font.bold = False
    
    # Column 0: Grey section
    set_cell_margins(cells[0], tight=True)
    set_cell_color_and_text(cells[0], text_left, grey_fill)
    
    # Column 1: White gap
    set_cell_margins(cells[1], tight=False)
    cells[1].text = ""
    
    # Column 2: Red/mauve section
    set_cell_margins(cells[2], tight=True)
    set_cell_color_and_text(cells[2], text_center, red_fill)
    
    # Column 3: White gap
    set_cell_margins(cells[3], tight=False)
    cells[3].text = ""
    
    # Column 4: Red/mauve section
    set_cell_margins(cells[4], tight=True)
    set_cell_color_and_text(cells[4], text_right, red_fill)
    
    # Remove all table borders
    tbl = table._tbl
    tblPr = tbl.tblPr if tbl.tblPr is not None else OxmlElement('w:tblPr')
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'nil')
        tblBorders.append(border)
    tblPr.append(tblBorders)
    if tbl.tblPr is None:
        tbl.insert(0, tblPr)


def add_section_header(doc, text):
    """Add bold and underlined section header - 10pt to match body"""
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(6)
    run = p.add_run(text)
    run.bold = True
    run.underline = True
    run.font.name = 'Arial'
    run.font.size = Pt(10)


def add_paragraph(doc, text, justify=False):
    """Add a normal paragraph with proper spacing"""
    p = doc.add_paragraph(text)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.line_spacing = 1.0
    if justify:
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    for run in p.runs:
        run.font.name = 'Arial'
        run.font.size = Pt(11)
    return p


def add_bullet(doc, text):
    """Add a standard bullet point with proper Word styling"""
    p = doc.add_paragraph(text, style='List Bullet')
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.line_spacing = 1.0
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    return p


def add_sub_bullet(doc, text):
    """Add a circle sub-bullet with proper alignment and hanging indent"""
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.5)
    p.paragraph_format.first_line_indent = Inches(-0.15)  # Hanging indent
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.line_spacing = 1.0
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run = p.add_run("○  " + text)
    run.font.name = 'Arial'
    run.font.size = Pt(10)
    return p


def add_sub_sub_bullet(doc, text):
    """Add a square sub-sub-bullet"""
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(1.0)
    p.paragraph_format.space_after = Pt(0)
    p.add_run("▪  " + text)
    return p


def calculate_total(data):
    """Calculate total fee"""
    fees = []
    for key in ['fee_sd', 'fee_dd', 'fee_cd', 'fee_bidding', 'fee_construction']:
        fees.append(data.get(key, '0'))
    
    if data.get('include_record_drawings'):
        fees.append(data.get('fee_record_drawings', '0'))
    
    total = 0
    for fee in fees:
        try:
            total += float(str(fee).replace(',', '').replace('$', '') or 0)
        except:
            pass
    
    return f"{total:,.0f}" if total > 0 else "___________"


def create_proposal_document(data):
    """Generate complete proposal matching template exactly"""
    doc = Document()
    
    # Setup styles
    setup_styles(doc)
    
    # Set margins
    for section in doc.sections:
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
    
    # Add header with hardcoded logo and footer
    section = doc.sections[0]
    add_header_with_logo(section)
    add_footer(section, "kimley-horn.com", 
               "200 Central Avenue Suite 600 St. Petersburg, FL 33701", 
               "727-547-3999")
    
    # === DOCUMENT CONTENT ===
    
    # Date
    add_paragraph(doc, data['date'])
    doc.add_paragraph()  # Blank line
    
    # Recipient
    add_paragraph(doc, f"{data['client_title']} {data['client_contact']}")
    add_paragraph(doc, data['company_name'])
    if data.get('address1'):
        add_paragraph(doc, data['address1'])
    if data.get('address2'):
        add_paragraph(doc, data['address2'])
    doc.add_paragraph()  # Blank line
    
    # Re: line
    add_paragraph(doc, "Re:\tLetter Agreement for Professional Services for")
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.5)
    p.add_run(data['project_name'])
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.5)
    p.add_run(f"{data['project_address']}, {data['project_city']}, {data['project_state']}")
    doc.add_paragraph()  # Blank line
    
    # Salutation
    last_name = data['client_contact'].split()[-1] if data['client_contact'] else "XXX"
    add_paragraph(doc, f"Dear {data['client_title']} {last_name}:")
    doc.add_paragraph()  # Blank line
    
    # Opening paragraph
    opening_text = f"Kimley-Horn and Associates, Inc. (\"Kimley-Horn\" or \"Consultant\") is pleased to submit this Letter Agreement (the \"Agreement\") to {data['company_name'] or '___________'} (\"Client\") for providing mechanical, electrical, plumbing, and fire protection consulting engineering services for the proposed {data['project_name'] or 'XX'} development located on {data['project_address'] or 'XXX Avenue'} in {data['project_city'] or 'XXX'}, {data['project_state'] or 'XX'} (\"Project\")."
    add_paragraph(doc, opening_text, justify=True)
    doc.add_paragraph()  # Blank line
    
    # === PROJECT UNDERSTANDING AND ASSUMPTIONS ===
    add_section_header(doc, "Project Understanding and Assumptions")
    
    intro_text = "Kimley-Horn's scope and fee are based on the following project understanding and assumptions. If any of these assumptions are not correct, then the scope and fee provided below may change:"
    add_paragraph(doc, intro_text, justify=True)
    doc.add_paragraph()  # Blank line
    
    # Add assumptions based on selections
    if data.get('is_new_building'):
        add_bullet(doc, f"{data['company_name'] or 'Client'} will be building a new building located at {data['project_address'] or 'XXX'}, {data['project_state'] or 'XX'}.")
    
    if data.get('is_renovation'):
        add_bullet(doc, f"The project will be a renovation to an existing tenant space located at {data['project_address'] or 'XXX'}.")
    
    if data.get('building_stories'):
        add_bullet(doc, f"The {data['project_name'] or 'XXX'} building is estimated to be roughly {data['building_stories']} stories with a total area of {data.get('total_area', 'XXX,000')} sf.")
    
    if data.get('construction_phases'):
        add_bullet(doc, f"The project will be constructed in {data['construction_phases']} phases of design, permitting and construction.")
    
    if data.get('separate_buildings'):
        add_bullet(doc, "The office and the parking garage will be two separate buildings connected by ground floor retail and common outdoor spaces.")
    
    if data.get('core_and_shell'):
        add_bullet(doc, f"The {data['project_name'] or 'XXX'} building will be provided as a core and shell building.")
    
    if data.get('leed_rating') and data['leed_rating'] != 'Not Applicable':
        add_bullet(doc, f"The Project will be designed to {data['leed_rating']}. Kimley-Horn will work with the LEED consultant on their assigned credits and provide the required calculations and documentation needed throughout the design phase.")
    
    if data.get('construction_budget'):
        add_bullet(doc, f"Kimley-Horn understands that the project is based on a ${data['construction_budget']} estimated construction budget.")
    
    if data.get('unit_types'):
        add_bullet(doc, f"Kimley-Horn will provide MEP design scope of services below for up to {data['unit_types']} unit types.")
    
    if data.get('typical_floors'):
        add_bullet(doc, f"Kimley-Horn will provide MEP design scope of services below for up to {data['typical_floors']} typical floors.")
    
    # Retail Core & Shell section
    if data.get('retail_core_shell'):
        doc.add_paragraph()  # Blank line before section
        add_bullet(doc, "All retail will be provided as core and shell. All retail core and shell spaces will be designed based on the following understanding:")
        
        if data.get('retail_electrical'):
            add_sub_bullet(doc, "Electrical systems will be designed as a meter center and empty conduits as needed for future tenant connection.")
            add_sub_sub_bullet(doc, "Design and engineering of tenant panel and transformers for tenant are not included in this scope of services.")
        
        if data.get('retail_plumbing'):
            add_sub_bullet(doc, "Plumbing systems will be provided with sanitary, vent, water, grease waste and gas stub-ins to the space and capped for future tenant connection. No plumbing connections and distribution piping will be designed for tenant spaces as part of this scope of services.")
        
        if data.get('retail_food_beverage'):
            add_sub_bullet(doc, "Retail spaces are to be food and beverage retail with cooking within the retail space. Occupancy loads provided by the Client or Client's architect and / or owner will be the basis for grease trap sizing.")
        
        if data.get('retail_mechanical'):
            add_sub_bullet(doc, "Mechanical systems will be provided as condenser water systems with piping stub-ins for future tenant provided water source heat pumps.")
    
    # HVAC Design Basis
    doc.add_paragraph()  # Blank line
    add_bullet(doc, "The HVAC design shall be based on the following:")
    
    hvac_descriptions = {
        'Centralized Chilled Water': 'The system will be designed as a centralized chilled water system with individual air handler per floor.',
        'Condenser Water': 'The systems will be designed as a condenser water system with centralized cooling tower and compressor driven air handlers per floor.',
        'Rooftop Units with VAV': 'The systems will be designed as central roof top units located on the roof and ducted down to the associated spaces with Variable Air Volume Units provided on each floor.',
        'Rooftop Units without VAV': 'The systems will be designed as central roof top units located on the roof and ducted down to the associated spaces without Variable Air Volume Units.',
        'VRF': 'The systems will be designed as Variable Refrigerant Flow units with heat recovery units located throughout the space.',
        'Split DX': 'The system will be designed as a split DX system with indoor air handlers located throughout the building.'
    }
    
    if data.get('hvac_system'):
        add_sub_bullet(doc, hvac_descriptions.get(data['hvac_system'], ''))
    
    if data.get('hvac_residential_highrise'):
        add_sub_bullet(doc, "Kimley-Horn will work with the Client's architect and Client to select the mechanical system during the conceptual and schematic design phase. Variable refrigerant flow, DX, and condenser water systems will be evaluated for this project.")
    
    if data.get('hvac_existing_reuse'):
        add_sub_bullet(doc, "The existing mechanical system will be reused in its current condition. Ductwork will be demolished back to the existing unit and replaced with all new ductwork and air distribution to accommodate the new architectural layout.")
    
    if data.get('outside_air_unit'):
        add_sub_bullet(doc, "Outside air will be provided by a dedicated 100% outside air unit located on the roof serving each of the units and the corridors.")
    
    exhaust_descriptions = {
        'Dedicated Roof Fan': 'Exhaust will be provided as a dedicated exhaust fan located on the roof.',
        'Individual Fans': 'Exhaust will be provided as individual exhaust fans discharging out of the side through louvers.',
        'Through OA Unit': 'Exhaust systems will be collected and routed back through the dedicated outside air unit.'
    }
    
    if data.get('exhaust_system'):
        add_sub_bullet(doc, exhaust_descriptions.get(data['exhaust_system'], ''))
    
    if data.get('parking_garage') == 'Open-Air':
        add_sub_bullet(doc, "The Parking garage will be designed as an open-air parking garage with no mechanical ventilation to be provided.")
    elif data.get('parking_garage') == 'Enclosed':
        add_sub_bullet(doc, "The Parking garage will be designed as an enclosed parking garage with mechanical ventilation.")
    
    if data.get('smoke_control'):
        add_sub_bullet(doc, "Smoke control system design is included in the below scope of work and will be designed per the rational analysis as provided from the life safety consultant.")
    
    if data.get('elevator_hoistway'):
        add_sub_bullet(doc, "Elevator hoist ways are enclosed lobbies and no hoist way pressurization will be designed.")
    
    # Plumbing Design Basis
    doc.add_paragraph()  # Blank line
    add_bullet(doc, "The plumbing design shall be based on the following:")
    
    if data.get('water_service') == 'Single Meter':
        add_sub_bullet(doc, "Domestic water design is included and for the purposes of this letter agreement is assumed that domestic water service will be provided as a single meter to the building from the public water main.")
    elif data.get('water_service') == 'Multiple Meters':
        add_sub_bullet(doc, "Domestic water design is included and for the purposes of this letter agreement is assumed that domestic water service will be provided to the building as multiple meters for each space from the public water main.")
    
    if data.get('roof_storm_drain'):
        add_sub_bullet(doc, "All roof storm drain fixture locations and roof sloping layouts shall be provided by the Client's architect.")
    
    if data.get('parking_garage_drain'):
        add_sub_bullet(doc, "All parking garage drain fixture locations and roof sloping layouts shall be provided by the Client's architect.")
    
    if data.get('water_oil_separator'):
        add_sub_bullet(doc, "The plumbing system for parking garage will include the design of a water oil separator system.")
    
    if data.get('sump_pump'):
        add_sub_bullet(doc, "Below grade parking includes the design of sump pump systems for drainage of the parking system.")
    
    if data.get('booster_pump'):
        add_sub_bullet(doc, "The domestic water system will be designed to include a booster pump system.")
    
    if data.get('sanitary_vent'):
        add_sub_bullet(doc, "Sanitary and vent system design is included in this scope of services.")
    
    if data.get('grease_waste'):
        add_sub_bullet(doc, "Grease waste system will be designed for cooking and restaurant spaces.")
    
    if data.get('natural_gas'):
        add_sub_bullet(doc, "Coordinate and design of the natural gas system to be used for domestic hot water heating or cooking.")
    
    if data.get('fuel_delivery'):
        add_sub_bullet(doc, "Coordinate and design of the fuel delivery system to be used for emergency power.")
    
    if data.get('roof_drainage') == 'Internal Drains':
        add_sub_bullet(doc, "Roof drainage system will be designed as internal roof drains with secondary overflows.")
    elif data.get('roof_drainage') == 'Gutters/Downspouts':
        add_sub_bullet(doc, "Roof drainage system will be designed as gutter and downspouts exterior to the building.")
    
    if data.get('civil_coordination'):
        add_sub_bullet(doc, "Coordination with the civil engineer is anticipated in the scope of services.")
    
    # Electrical Design Basis
    doc.add_paragraph()  # Blank line
    add_bullet(doc, "The electrical design shall be based on the following:")
    
    if data.get('existing_electrical_renovation'):
        add_bullet(doc, "The existing electrical system is being renovated and anticipated to exceed the loads currently in the space and therefore a 30-day load study, provided by the Client, will be required prior to issuing final construction documents.")
    
    if data.get('power_receptacles'):
        add_bullet(doc, "Power receptacle layout and design is included in this scope of services.")
    
    if data.get('core_shell_electrical'):
        add_bullet(doc, "Power receptacle layout and design consists of the front and back of house areas described above. All core and shell areas shall be provided with only the anticipated panel sizing and conduits for future tenants to route power through.")
    
    if data.get('lighting_coordination'):
        add_bullet(doc, "Lighting design for all front of house areas will be coordinated with the Client's architect and / or their lighting designer. The Client's lighting designer shall provide Kimley-Horn with all front of house lighting fixture layouts, schedules, control diagrams and switching layouts, along with CAD plans showing lighting photometrics to be included in the electrical engineering plans for building permit.")
    
    if data.get('lightning_protection') == 'Included':
        add_bullet(doc, "Building lightning protection design is included as a performance-based design.")
    else:
        add_bullet(doc, "Building lightning protection design is excluded in this scope of work.")
    
    if data.get('emergency_generator') == 'Included':
        add_bullet(doc, "Emergency generator design is included for code required life safety systems only.")
    else:
        add_bullet(doc, "Emergency generator design is excluded from this scope of services.")
    
    if data.get('ev_charging') == 'Included':
        add_bullet(doc, f"Electrical vehicle charging design is included for up to {data.get('ev_ready_spaces', 'XX')} electrical vehicle ready spaces, and {data.get('ev_capable_spaces', 'XX')} electrical vehicle capable spaces.")
        add_sub_bullet(doc, "EV Ready spaces are provided with dedicated EV charging equipment, feeders, and raceways.")
        add_sub_bullet(doc, "EV Capable spaces are provided with future capacity in electrical switchgear and spare conduits routed from the electrical room to five feet outside the building.")
    else:
        add_bullet(doc, "Electrical vehicle charging design is excluded from this scope of work.")
    
    if data.get('fire_alarm'):
        add_bullet(doc, "Fire Alarm design to consist of schematic plans and \"preliminary based design\" (FAC 61G15) specifications. Detailed fire sprinkler drawings shall be provided by the Client's sprinkler contractor.")
    
    if data.get('technology_design'):
        add_bullet(doc, "Technology design services provided in the MEP design scope of services below will have the design for the pathway and backboxes only.")
    
    # Fire Protection Design Basis
    doc.add_paragraph()  # Blank line
    add_bullet(doc, "The Fire protection design shall be based on the following:")
    
    add_sub_bullet(doc, "Fire protection design to consist of schematic plans and \"performance-based\" (FAC 61G15) specifications. Detailed fire sprinkler drawings shall be provided by the Client's fire sprinkler contractor.")
    add_sub_bullet(doc, "The Client's fire sprinkler contractor shall be responsible for fire sprinkler permit documents.")
    
    if data.get('fire_pump') == 'Included':
        add_sub_bullet(doc, "The design of a fire pump is included in the scope of services.")
    else:
        add_sub_bullet(doc, "The design of a fire pump is not included in this scope of services.")
    
    # Weekly Meetings
    if data.get('weekly_meetings'):
        doc.add_paragraph()  # Blank line
        meeting_text = "For budgeting purposes, Kimley-Horn assumes that weekly meetings will occur throughout each design phase task provided below beginning with the kickoff meeting and in accordance with the duration of each design phase task outlined below in the scope of services. Should the design schedule be extended beyond its initially established timeframe, attendance at any additional meetings may be considered an additional service and subject to additional charges."
        add_bullet(doc, meeting_text)
    
    # Revit
    doc.add_paragraph()  # Blank line
    revit_intro = f"Revit: Kimley-Horn utilizes Revit as the basis for Kimley-Horn's design software. Kimley-Horn's Revit model will be prepared to a Level of Development (LOD) {data.get('revit_lod', '300')} standard which will consist of the following:"
    add_bullet(doc, revit_intro)
    
    add_sub_bullet(doc, "Model elements represented for all ductwork, piping, conduits (6\" or greater), duct banks, panel boards, mechanical and plumbing equipment, fire protection equipment. Refrigerant piping will be modeled for the intent of routing purposes only.")
    add_sub_bullet(doc, "Interdisciplinary coordination with MEPFP systems in the building to address major coordination items such as chases, above ceiling heights, coordination with structural members and foundations. Though major coordination will be done, the model will not be clash free. Any elements under 6\" in diameter or depth will be modeled for design intent only and will be left up to the Client's contractor for all final coordination.")
    add_sub_bullet(doc, "All lighting and plumbing fixtures will be placed and controlled by the Client's architect in their model and referenced and modeled in the Kimley-Horn Revit model.")
    
    if data.get('revit_coordination_hours'):
        add_sub_bullet(doc, f"Meeting with Client's architect and Client's other subconsultants will be for coordination only (Kimley-Horn will attend up to {data['revit_coordination_hours']} hrs for meetings). Clash detection meetings are not part of this scope of services and can be provided as an additional service.")
    
    # === SCOPE OF SERVICES ===
    doc.add_paragraph()  # Blank line
    add_section_header(doc, "Scope of Services")
    
    # Task 110 - Schematic Design
    p = doc.add_paragraph()
    run = p.add_run("Task 110 – Schematic Design")
    run.bold = True
    run.font.name = 'Arial'
    run.font.size = Pt(11)
    doc.add_paragraph()  # Blank line
    
    add_bullet(doc, "Attend one (1) Client and / or architect kickoff meeting for project initiation.")
    
    if data.get('sd_existing_survey'):
        add_bullet(doc, "Attend one (1) existing building site survey for review of existing building systems.")
        add_bullet(doc, "Prepare site visit observation report outlining field observations and meeting notes from the existing building site survey.")
        if data.get('sd_site_visit_hours'):
            add_bullet(doc, f"Kimley-Horn will attend a site visit to observe the existing conditions of the mechanical and electrical systems serving the {data['project_name'] or 'XXX'}. The site visit will include up to two Kimley-Horn representatives for up to {data['sd_site_visit_hours']} hours on site, plus travel time.")
    
    add_bullet(doc, f"Schematic Design phase is anticipated to last up to {data.get('sd_weeks', '3')} weeks.")
    
    if data.get('sd_meeting_hours'):
        add_bullet(doc, f"Kimley-Horn will attend up to one (1) weekly coordination meeting per week, for up to {data['sd_meeting_hours']} hours, as requested by the Client for the duration of the Schematic Design phase.")
        if data.get('sd_total_meetings'):
            add_sub_bullet(doc, f"For the purposes of this letter agreement, Kimley-Horn assumes there will be {data['sd_total_meetings']} total weekly design meetings for this task.")
    
    add_bullet(doc, "Prepare preliminary load estimates for coordination of MEP equipment for space requirements.")
    add_bullet(doc, "Coordinate locations of incoming services to buildings with the civil engineer.")
    add_bullet(doc, "Coordinate the approximate location of the existing utility infrastructure.")
    add_bullet(doc, "Coordinate the availability and requirements of storm, water, sewer, and reclaim water for points of connection and discharge with the civil engineer.")
    add_bullet(doc, "Prepare preliminary sanitary sizing for civil engineering coordination and connection.")
    add_bullet(doc, "Prepare Schematic Design narratives describing the proposed MEP/FP systems.")
    add_bullet(doc, "Respond to up to two (2) rounds of schematic design narrative comments from Client.")
    
    # Task 120 - Design Development
    doc.add_paragraph()  # Blank line
    p = doc.add_paragraph()
    run = p.add_run("Task 120 – Design Development")
    run.bold = True
    run.font.name = 'Arial'
    run.font.size = Pt(11)
    doc.add_paragraph()  # Blank line
    
    add_bullet(doc, "Upon written approval of the Schematic Design narrative by the Client, Kimley-Horn will proceed into the Design Development phase.")
    
    if data.get('dd_weeks'):
        add_bullet(doc, f"Kimley-Horn anticipates the Design Development phase is anticipated to last up to {data['dd_weeks']} weeks.")
    
    if data.get('dd_meeting_hours'):
        add_bullet(doc, f"Kimley-Horn will attend up to one (1) weekly coordination meeting per week, for up to {data['dd_meeting_hours']} hours, as requested by the Client for the duration of the Design Development phase.")
        if data.get('dd_total_meetings'):
            add_sub_bullet(doc, f"For the purposes of this letter agreement, Kimley-Horn assumes there will be {data['dd_total_meetings']} total weekly design meetings for this task.")
    
    add_bullet(doc, "Kimley-Horn will provide update design calculations, equipment selections, and fixture selections.")
    add_bullet(doc, "The Revit model will be updated to show major system equipment locations, routing, and coordinate with Client's architect and their sub consultants.")
    add_bullet(doc, "Prepare and deliver Design Development drawings in PDF format.")
    add_bullet(doc, f"Respond to up to {data.get('dd_rounds', '2')} rounds of owner Design Development (DD) review comments.")
    
    # Task 130 - Construction Documents
    doc.add_paragraph()  # Blank line
    p = doc.add_paragraph()
    run = p.add_run("Task 130 – Construction Documents")
    run.bold = True
    run.font.name = 'Arial'
    run.font.size = Pt(11)
    doc.add_paragraph()  # Blank line
    
    add_bullet(doc, "Upon written approval of the Design Development deliverables by the Client, Kimley-Horn will proceed into the Construction Document phase.")
    
    if data.get('cd_weeks'):
        add_bullet(doc, f"Kimley-Horn anticipates the Construction Document phase to last up to {data['cd_weeks']} weeks.")
    
    if data.get('cd_meeting_hours'):
        add_bullet(doc, f"Kimley-Horn will attend up to one (1) weekly coordination meeting per week, for up to {data['cd_meeting_hours']} hours, as requested by the Client for the duration of the Construction Document phase.")
        if data.get('cd_total_meetings'):
            add_sub_bullet(doc, f"For the purposes of this letter agreement, Kimley-Horn assumes there will be {data['cd_total_meetings']} total weekly design meetings for this task.")
    
    add_bullet(doc, "Finalized equipment, calculations, and fixture selections.")
    add_bullet(doc, f"Prepare one (1) Construction Document progress drawing PDF submittal at approximately {data.get('cd_percentages', '25%, 50%, 75%, and 90%')} CDs.")
    add_bullet(doc, "Kimley-Horn will respond to up to two (2) rounds of comments, from the Client, for each submittal.")
    add_bullet(doc, "Kimley-Horn will respond to up to two (2) rounds of 90% Construction Documents review comments.")
    add_bullet(doc, "Provide and submit response narrative addressing permit comments provided by the AHJ and permit reviewers.")
    add_bullet(doc, "Coordinate with the Client's architect and Client's other consultants addressing permit comments.")
    add_bullet(doc, "Prepare final Construction Documents and specifications for bidding and final submission to the building department.")
    add_bullet(doc, "Specifications will be prepared as standard book specs or sheet specs.")
    add_bullet(doc, "Submit stamped and signed PDF drawings and specifications for building permit application and final building permit coordination. All municipal permit coordination is to be handled by the Client's project architect.")
    
    # Task 140 - Bidding
    doc.add_paragraph()  # Blank line
    p = doc.add_paragraph()
    run = p.add_run("Task 140 – Bidding and Negotiations")
    run.bold = True
    run.font.name = 'Arial'
    run.font.size = Pt(11)
    doc.add_paragraph()  # Blank line
    
    add_bullet(doc, "Kimley-Horn will attend up to one (1) pre-bid meeting with potential bidders online or in person as requested by the Client.")
    add_bullet(doc, "Consultant will review up to two (2) rounds of sub-contractor bids and provide written feedback to Client on received bids.")
    
    # Task 150 - Construction Phase
    doc.add_paragraph()  # Blank line
    p = doc.add_paragraph()
    run = p.add_run("Task 150 – Limited Construction Phase Services")
    run.bold = True
    run.font.name = 'Arial'
    run.font.size = Pt(11)
    doc.add_paragraph()  # Blank line
    
    if data.get('site_visits'):
        site_visit_text = f"Site Visits and Construction Observation. Kimley-Horn will make up to {data['site_visits']} site visits to observe the progress of the work. Observations will not be exhaustive or extend to every aspect of Contractor's work, but will be limited to spot checking, and similar methods of general observation."
        add_bullet(doc, site_visit_text)
    
    add_bullet(doc, "Kimley-Horn will not supervise, direct, or control Contractor's work, and will not have authority to stop the Work or responsibility for the means, methods, techniques, equipment choice and use, schedules, or procedures of construction selected by Contractor.")
    add_bullet(doc, "Kimley-Horn is not responsible for any duties assigned to it in the construction contract that are not expressly provided for in this Agreement.")
    add_bullet(doc, "Shop Drawings and Samples. Kimley-Horn will review Shop Drawings and Samples and other data which Contractor is required to submit, but only for general conformance with the Contract Documents.")
    add_bullet(doc, "Substitutes and \"or-equal/equivalent.\" Kimley-Horn will evaluate the acceptability of substitute or \"or-equal/equivalent\" materials and equipment proposed by Contractor in accordance with the Contract Documents.")
    add_bullet(doc, "Kimley-Horn will respond to RFIs and Submittals within a reasonable amount of time, but not more than five (5) business days for RFIs and ten (10) business days for submittals.")
    
    # Task 160 - Record Drawings (optional)
    if data.get('include_record_drawings'):
        doc.add_paragraph()  # Blank line
        p = doc.add_paragraph()
        run = p.add_run("Task 160 – Record Drawings")
        run.bold = True
        run.font.name = 'Arial'
        run.font.size = Pt(11)
        doc.add_paragraph()  # Blank line
        
        record_text = "Kimley-Horn will prepare a record drawing showing significant changes reported by the Contractor or made to the design by Kimley-Horn. Record drawings are not guaranteed to be as-built but will be based on information made available."
        add_paragraph(doc, record_text, justify=True)
        doc.add_paragraph()  # Blank line
        
        if data.get('record_drawings_hours'):
            add_bullet(doc, f"Given the unknown quantity of revisions, Kimley-Horn has allocated {data['record_drawings_hours']} hours for coordination and responses in this task. Additional responses may require additional fee.")
    
    # === ADDITIONAL SERVICES ===
    doc.add_paragraph()  # Blank line
    add_section_header(doc, "Additional Services")
    
    additional_intro = "Any services not specifically provided for in the above scope of services will be billed as additional services and performed at our then current hourly rates. Additional services we can provide include, but are not limited to, the following:"
    add_paragraph(doc, additional_intro, justify=True)
    doc.add_paragraph()  # Blank line
    
    additional_services = [
        "Commissioning Services.",
        "Technology Design (detailed design of access control, telecommunication, audio visual)",
        "Sustainable Certification",
        "LEED Design or Administration.",
        "Life Cycle Cost Analysis",
        "Cost Estimating",
        "Solar Photovoltaic Design",
        "Record Drawings",
        "Value engineering request, design changes, and meetings.",
        "Revit Modeling beyond standard Level of Development (LOD) 300.",
        "Project phasing or fast track construction bid / documentation.",
        "Construction administration visits beyond what is listed in the scope of services above.",
        "As-built drawings or record drawings",
        "Civil Engineering Services",
        "Structural Engineering"
    ]
    
    for service in additional_services:
        add_bullet(doc, service)
    
    # === INFORMATION PROVIDED BY CLIENT ===
    doc.add_paragraph()  # Blank line
    add_section_header(doc, "Information Provided by Client")
    
    client_intro = "Kimley-Horn shall be entitled to rely on the completeness and accuracy of all information provided by the Client or the Client's consultants or representatives. The Client shall provide all information requested by Kimley-Horn during the project, including but not limited to the following:"
    add_paragraph(doc, client_intro, justify=True)
    doc.add_paragraph()  # Blank line
    
    client_info_items = [
        "Architectural floor plan, site plans, life safety plans, elevations, building sections, reflected ceiling plans and architectural floor plan backgrounds, complete with room names, numbers and rated or special wall construction, will be provided by the Client's architect during the course of the design (Kimley-Horn standard is Revit).",
        "Room and equipment cut sheet information for each area, indicating equipment and furniture locations, quantity of each type of outlet, receptacle, special lighting and plumbing equipment, and connection for services as part of the Kimley-Horn design.",
        "Project contacts for all consultants and or sub consultants on the project.",
        "Client will provide all plumbing fixture cut sheets and locations to be incorporated into the plumbing scope of services above.",
        "Client will provide all lighting fixture cut sheets and locations to be incorporated into the electrical scope of services above.",
        "Civil, site drawings and surveys, indicating all underground and overhead mechanical, plumbing and electrical site utilities, which may affect design.",
        "Fire hydrant flow test data, performed at the hydrants required by the design as coordinated with the MEP, civil engineer and agency having jurisdiction."
    ]
    
    for item in client_info_items:
        add_bullet(doc, item)
    
    # === SCHEDULE ===
    doc.add_paragraph()  # Blank line
    add_section_header(doc, "Schedule")
    
    schedule_text = "Kimley-Horn will perform the services as expeditiously as practicable with the goal of meeting a mutually agreed upon schedule."
    add_paragraph(doc, schedule_text, justify=True)
    
    # === FEE AND EXPENSES ===
    doc.add_paragraph()  # Blank line
    add_section_header(doc, "Fee and Expenses")
    
    fee_intro = "Kimley-Horn will perform the services in Tasks 110 – 150 for the total lump sum labor fee below. Individual task amounts are informational only. In addition to the lump sum labor fee, direct reimbursable expenses such as express delivery services, fees, air travel, and other direct expenses will be billed at 1.15 times cost. All permitting, application, and similar project fees will be paid directly by the Client."
    add_paragraph(doc, fee_intro, justify=True)
    doc.add_paragraph()  # Blank line
    
    # Fee table
    num_rows = 7 if not data.get('include_record_drawings') else 8
    fee_table = doc.add_table(rows=num_rows, cols=3)
    fee_table.style = 'Table Grid'
    fee_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    
    # Header row
    header_cells = fee_table.rows[0].cells
    header_cells[0].text = "Task Number and Name"
    header_cells[1].text = "Fee"
    header_cells[2].text = "Type"
    
    for cell in header_cells:
        cell_shading = OxmlElement('w:shd')
        cell_shading.set(qn('w:fill'), '8B0000')
        cell._tc.get_or_add_tcPr().append(cell_shading)
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.bold = True
                run.font.color.rgb = RGBColor(255, 255, 255)
                run.font.size = Pt(10)
                run.font.name = 'Arial'
    
    # Data rows
    fee_data = [
        ("110 Schematic Design Phase", data.get('fee_sd', 'XXX'), "Lump Sum"),
        ("120 Design Development", data.get('fee_dd', 'XXX'), "Lump Sum"),
        ("130 Construction Documents", data.get('fee_cd', 'XXX'), "Lump Sum"),
        ("140 Bidding and Negotiations", data.get('fee_bidding', 'XXX'), "Lump Sum"),
        ("150 Limited Construction Phase Services", data.get('fee_construction', 'XXX'), "Lump Sum"),
    ]
    
    if data.get('include_record_drawings'):
        fee_data.append(("160 Record Drawings", data.get('fee_record_drawings', 'XXX'), "Lump Sum"))
    
    for i, (task, fee, fee_type) in enumerate(fee_data):
        row = fee_table.rows[i + 1]
        row.cells[0].text = task
        row.cells[1].text = f"${fee}"
        row.cells[2].text = fee_type
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.name = 'Arial'
                    run.font.size = Pt(11)
    
    # Total row
    total = calculate_total(data)
    total_row = fee_table.rows[num_rows - 1]
    total_row.cells[0].text = "Total"
    total_row.cells[1].text = f"${total}"
    total_row.cells[2].text = ""
    for cell in total_row.cells:
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.bold = True
                run.font.name = 'Arial'
                run.font.size = Pt(11)
    
    doc.add_paragraph()  # Blank line
    
    fee_text1 = "Lump sum fees will be invoiced monthly based upon the overall percentage of services performed. Reimbursable expenses will be invoiced based upon expenses incurred."
    add_paragraph(doc, fee_text1, justify=True)
    doc.add_paragraph()  # Blank line
    
    fee_text2 = "Payment will be due within 25 days of your receipt of the invoice and should include the invoice number and Kimley-Horn project number."
    add_paragraph(doc, fee_text2, justify=True)
    doc.add_paragraph()  # Blank line
    
    fee_text3 = "This scope of services and associated fee are predicated on the assumption that no significant architectural design changes will occur following the Final Design Development (DD) stage. Should any substantial architectural design modifications be requested after the Final DD deliverable, additional design fees will be required to address and incorporate such changes."
    add_paragraph(doc, fee_text3, justify=True)
    
    # === CLOSURE ===
    doc.add_paragraph()  # Blank line
    add_section_header(doc, "Closure")
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    run1 = p.add_run("In addition to the matters set forth herein, our Agreement shall include and be subject to, and only to, the attached Standard Provisions, which are incorporated by reference. As used in the Standard Provisions, \"Kimley-Horn\" shall refer to Kimley-Horn and Associates, Inc., and \"Client\" shall refer to ")
    run2 = p.add_run(data['company_name'] or "___Insert Client's Legal Entity Name___")
    run2.font.highlight_color = 7
    p.add_run(".")
    doc.add_paragraph()  # Blank line
    
    invoice_text = "Kimley-Horn, in an effort to expedite invoices and reduce paper waste, submits invoices via email in a PDF. We can also provide a paper copy via regular mail if requested. Please include the invoice number and Kimley-Horn project number with all payments. Please provide the following information:"
    add_paragraph(doc, invoice_text, justify=True)
    doc.add_paragraph()  # Blank line
    
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.5)
    p.add_run(f"____ Please email all invoices to {data.get('invoice_email', '___________________________')}")
    
    p = doc.add_paragraph()
    p.paragraph_format.left_indent = Inches(0.5)
    p.add_run(f"____ Please copy {data.get('invoice_copy', '_______________________________________')}")
    doc.add_paragraph()  # Blank line
    
    proceed_text = "To proceed with the services, please have an authorized person sign this Agreement below and return to us. We will commence services only after we have received a fully-executed agreement. Fees and times stated in this Agreement are valid for sixty (60) days after the date of this letter."
    add_paragraph(doc, proceed_text, justify=True)
    doc.add_paragraph()  # Blank line
    
    rfi_text = "To ensure proper set up of your projects so that we can get started, please complete and return with the signed copy of this Agreement the attached Request for Information. Failure to supply this information could result in delay in starting work on this project."
    add_paragraph(doc, rfi_text, justify=True)
    doc.add_paragraph()  # Blank line
    
    add_paragraph(doc, "We appreciate the opportunity to provide these services. Please contact me if you have any questions.")
    doc.add_paragraph()  # Blank line
    
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    p.add_run("Sincerely,")
    doc.add_paragraph()  # Blank line
    doc.add_paragraph()  # Blank line
    
    p = doc.add_paragraph()
    run = p.add_run("KIMLEY-HORN AND ASSOCIATES, INC.")
    run.bold = True
    run.font.name = 'Arial'
    doc.add_paragraph()  # Blank line
    doc.add_paragraph()  # Blank line
    
    # Signature table
    sig_table = doc.add_table(rows=1, cols=2)
    sig_table.autofit = False
    sig_table.columns[0].width = Inches(3.0)
    sig_table.columns[1].width = Inches(3.0)
    
    sig_table.rows[0].cells[0].text = f"{data.get('project_manager', 'Clayton Scelzi')}\nProject Manager"
    sig_table.rows[0].cells[1].text = f"{data.get('senior_vp', 'Scott W. Gilner, PE')}\nSenior Vice President"
    
    for cell in sig_table.rows[0].cells:
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.name = 'Arial'
                run.font.size = Pt(11)
    
    # === CLIENT SIGNATURE PAGE ===
    doc.add_page_break()
    
    signature_instructions = "If the recipient changes the legal entity name or signs as any name other than the client named in the opening address block, do not accept this and prepare a new Letter Agreement with the appropriate client identified after discussion with the client."
    add_paragraph(doc, signature_instructions)
    doc.add_paragraph()  # Blank line
    
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(12)
    run = p.add_run("CORRECT CLIENT ENTITY – ALL CAPS – CHECK SUNBIZ.")
    run.bold = True
    run.font.name = 'Arial'
    doc.add_paragraph()  # Blank line
    doc.add_paragraph()  # Blank line
    
    add_paragraph(doc, "SIGNED: _________________________________")
    doc.add_paragraph()  # Blank line
    add_paragraph(doc, "PRINTED NAME: _________________________________")
    doc.add_paragraph()  # Blank line
    add_paragraph(doc, "TITLE: _________________________________")
    doc.add_paragraph()  # Blank line
    add_paragraph(doc, "DATE: _________________________________")
    doc.add_paragraph()  # Blank line
    doc.add_paragraph()  # Blank line
    
    add_paragraph(doc, "Client's Federal Tax ID: _________________________________")
    doc.add_paragraph()  # Blank line
    add_paragraph(doc, "Client's Business License No.: _________________________________")
    doc.add_paragraph()  # Blank line
    add_paragraph(doc, "Client's Street Address: _________________________________")
    doc.add_paragraph()  # Blank line
    doc.add_paragraph()  # Blank line
    
    add_paragraph(doc, "Attachment – Request for Information")
    add_paragraph(doc, "Attachment – Standard Provisions")
    
    return doc


# ============== STREAMLIT FORM ==============

if 'form_data' not in st.session_state:
    st.session_state.form_data = {}

tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
    "📋 Client & Project", 
    "🔧 Assumptions", 
    "⚙️ MEP Systems", 
    "📅 Schedule", 
    "💰 Fees",
    "📄 Generate"
])

with tab1:
    st.subheader("Client Information")
    with st.expander("ℹ️ Why we need this", expanded=False):
        st.info("Client information appears in the proposal header and throughout the document.")
    
    col1, col2 = st.columns(2)
    with col1:
        client_title = st.selectbox("Client Title", ["Mr.", "Mrs.", "Ms.", "Dr."],
                                    help="Professional title of the client contact")
        client_contact = st.text_input("Client Contact Name", 
                                      placeholder="John Smith",
                                      help="Full name of primary contact")
        company_name = st.text_input("Company Name (Legal Entity)", 
                                    placeholder="Company Name, LLC",
                                    help="⚠️ Must match legal entity exactly")
    with col2:
        address1 = st.text_input("Address Line 1", placeholder="123 Main Street")
        address2 = st.text_input("Address Line 2 (City, State ZIP)", 
                                placeholder="Tampa, FL 33601")
        proposal_date = st.date_input("Proposal Date", datetime.now())
    
    st.markdown("---")
    st.subheader("Project Information")
    col1, col2 = st.columns(2)
    with col1:
        project_name = st.text_input("Project Name", 
                                    placeholder="Downtown Office Complex",
                                    help="Official project name")
        project_address = st.text_input("Project Address", 
                                       placeholder="456 Business Avenue")
    with col2:
        project_city = st.text_input("City", placeholder="St. Petersburg")
        project_state = st.text_input("State", value="FL", max_chars=2)

with tab2:
    st.subheader("Project Understanding & Assumptions")
    st.info("✅ Select all applicable items for the Project Understanding section.")
    
    col1, col2 = st.columns(2)
    with col1:
        is_new_building = st.checkbox("New Building Construction", value=True)
        is_renovation = st.checkbox("Renovation to Existing Space")
        separate_buildings = st.checkbox("Separate Office & Parking Buildings")
        core_and_shell = st.checkbox("Core and Shell Building")
    with col2:
        building_stories = st.text_input("Building Stories", placeholder="10")
        total_area = st.text_input("Total Area (SF)", placeholder="150,000")
        construction_phases = st.text_input("Construction Phases", placeholder="2")
        construction_budget = st.text_input("Construction Budget ($)", 
                                           placeholder="25,000,000")
    
    col1, col2 = st.columns(2)
    with col1:
        leed_rating = st.selectbox("LEED Rating", 
                                  ["Not Applicable", "LEED Certified", "LEED Silver", 
                                   "LEED Gold", "LEED Platinum"])
        unit_types = st.text_input("Unit Types", placeholder="8")
    with col2:
        typical_floors = st.text_input("Typical Floors", placeholder="5")
    
    st.markdown("---")
    st.subheader("Retail Core & Shell")
    retail_core_shell = st.checkbox("Include Retail Core & Shell Provisions")
    if retail_core_shell:
        col1, col2 = st.columns(2)
        with col1:
            retail_electrical = st.checkbox("Electrical: Meter center and conduits", value=True)
            retail_plumbing = st.checkbox("Plumbing: Stub-ins for future tenant", value=True)
        with col2:
            retail_food_beverage = st.checkbox("Food & Beverage Retail")
            retail_mechanical = st.checkbox("Mechanical: Condenser water stub-ins")

with tab3:
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("🌡️ HVAC System")
        st.info("Select the primary HVAC system type")
        hvac_centralized_chilled = st.checkbox("Centralized Chilled Water")
        hvac_condenser = st.checkbox("Condenser Water")
        hvac_rooftop_vav = st.checkbox("Rooftop Units with VAV")
        hvac_rooftop_no_vav = st.checkbox("Rooftop Units without VAV")
        hvac_vrf = st.checkbox("VRF")
        hvac_split_dx = st.checkbox("Split DX")
        
        # Determine which system is selected
        hvac_systems = {
            'Centralized Chilled Water': hvac_centralized_chilled,
            'Condenser Water': hvac_condenser,
            'Rooftop Units with VAV': hvac_rooftop_vav,
            'Rooftop Units without VAV': hvac_rooftop_no_vav,
            'VRF': hvac_vrf,
            'Split DX': hvac_split_dx
        }
        hvac_system = next((k for k, v in hvac_systems.items() if v), 'Centralized Chilled Water')
        
        st.markdown("---")
        hvac_residential_highrise = st.checkbox("Residential Highrise (system TBD)")
        hvac_existing_reuse = st.checkbox("Reuse Existing Mechanical System")
        outside_air_unit = st.checkbox("Dedicated Outside Air Unit", value=True)
        
        st.markdown("**Exhaust System**")
        exhaust_dedicated_roof = st.checkbox("Dedicated Roof Fan", value=True)
        exhaust_individual = st.checkbox("Individual Fans")
        exhaust_through_oa = st.checkbox("Through OA Unit")
        
        # Determine exhaust system
        exhaust_systems = {
            'Dedicated Roof Fan': exhaust_dedicated_roof,
            'Individual Fans': exhaust_individual,
            'Through OA Unit': exhaust_through_oa
        }
        exhaust_system = next((k for k, v in exhaust_systems.items() if v), 'Dedicated Roof Fan')
        
        st.markdown("**Parking Garage**")
        parking_open_air = st.checkbox("Open-Air", value=True)
        parking_enclosed = st.checkbox("Enclosed")
        parking_garage = 'Open-Air' if parking_open_air else 'Enclosed'
        
        smoke_control = st.checkbox("Smoke Control System")
        elevator_hoistway = st.checkbox("Elevator Hoistway (no pressurization)")
        
        st.markdown("---")
        st.subheader("🚿 Plumbing System")
        st.markdown("**Water Service**")
        water_single = st.checkbox("Single Meter", value=True)
        water_multiple = st.checkbox("Multiple Meters")
        water_service = 'Single Meter' if water_single else 'Multiple Meters'
        
        st.markdown("**Roof Drainage**")
        roof_internal = st.checkbox("Internal Drains", value=True)
        roof_gutters = st.checkbox("Gutters/Downspouts")
        roof_drainage = 'Internal Drains' if roof_internal else 'Gutters/Downspouts'
        
        roof_storm_drain = st.checkbox("Roof Storm Drain (by Architect)", value=True)
        parking_garage_drain = st.checkbox("Parking Garage Drain")
        water_oil_separator = st.checkbox("Water Oil Separator")
        sump_pump = st.checkbox("Sump Pump (Below Grade)")
        booster_pump = st.checkbox("Booster Pump System")
        sanitary_vent = st.checkbox("Sanitary and Vent System", value=True)
        grease_waste = st.checkbox("Grease Waste System")
        natural_gas = st.checkbox("Natural Gas System")
        fuel_delivery = st.checkbox("Fuel Delivery System")
        civil_coordination = st.checkbox("Civil Engineer Coordination", value=True)
    
    with col2:
        st.subheader("⚡ Electrical System")
        existing_electrical_renovation = st.checkbox("Existing Electrical Renovation (load study)")
        power_receptacles = st.checkbox("Power Receptacle Design", value=True)
        core_shell_electrical = st.checkbox("Core & Shell Electrical Only")
        lighting_coordination = st.checkbox("Lighting Design Coordination", value=True)
        
        st.markdown("**Lightning Protection**")
        lightning_excluded = st.checkbox("Excluded", value=True, key="lightning_excl")
        lightning_included = st.checkbox("Included", key="lightning_incl")
        lightning_protection = 'Included' if lightning_included else 'Excluded'
        
        st.markdown("**Emergency Generator**")
        generator_excluded = st.checkbox("Excluded", value=True, key="gen_excl")
        generator_included = st.checkbox("Included", key="gen_incl")
        emergency_generator = 'Included' if generator_included else 'Excluded'
        
        st.markdown("**EV Charging**")
        ev_excluded = st.checkbox("Excluded", value=True, key="ev_excl")
        ev_included = st.checkbox("Included", key="ev_incl")
        ev_charging = 'Included' if ev_included else 'Excluded'
        
        if ev_included:
            ev_ready_spaces = st.text_input("EV Ready Spaces", placeholder="10")
            ev_capable_spaces = st.text_input("EV Capable Spaces", placeholder="20")
        else:
            ev_ready_spaces = ""
            ev_capable_spaces = ""
        
        fire_alarm = st.checkbox("Fire Alarm Design", value=True)
        technology_design = st.checkbox("Technology Design (pathway only)", value=True)
        
        st.markdown("---")
        st.subheader("🔥 Fire Protection")
        fire_pump_excluded = st.checkbox("Fire Pump Excluded", value=True, key="fp_excl")
        fire_pump_included = st.checkbox("Fire Pump Included", key="fp_incl")
        fire_pump = 'Included' if fire_pump_included else 'Excluded'
        
        st.markdown("---")
        st.subheader("🏗️ Revit Standards")
        weekly_meetings = st.checkbox("Weekly Meetings", value=True)
        
        st.markdown("**Revit LOD**")
        lod_200 = st.checkbox("200", key="lod_200")
        lod_300 = st.checkbox("300", value=True, key="lod_300")
        lod_350 = st.checkbox("350", key="lod_350")
        lod_400 = st.checkbox("400", key="lod_400")
        
        lod_options = {'200': lod_200, '300': lod_300, '350': lod_350, '400': lod_400}
        revit_lod = next((k for k, v in lod_options.items() if v), '300')
        
        revit_coordination_hours = st.text_input("Revit Coordination Hours", placeholder="Optional")

with tab4:
    st.subheader("Design Phase Schedule")
    st.info("📅 Enter duration and meeting details for each design phase")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("**Task 110 - Schematic Design**")
        sd_existing_survey = st.checkbox("Include Existing Building Survey")
        if sd_existing_survey:
            sd_site_visit_hours = st.text_input("Site Visit Hours", placeholder="4", key="sd_hours")
        else:
            sd_site_visit_hours = ""
        sd_weeks = st.text_input("SD Duration (weeks)", value="3")
        sd_meeting_hours = st.text_input("SD Meeting Hours/Week", placeholder="1")
        sd_total_meetings = st.text_input("Total SD Meetings", placeholder="3")
    
    with col2:
        st.markdown("**Task 120 - Design Development**")
        dd_weeks = st.text_input("DD Duration (weeks)", placeholder="6")
        dd_meeting_hours = st.text_input("DD Meeting Hours/Week", placeholder="1")
        dd_total_meetings = st.text_input("Total DD Meetings", placeholder="6")
        dd_rounds = st.text_input("DD Review Rounds", value="2")
    
    with col3:
        st.markdown("**Task 130 - Construction Documents**")
        cd_weeks = st.text_input("CD Duration (weeks)", placeholder="12")
        cd_meeting_hours = st.text_input("CD Meeting Hours/Week", placeholder="1")
        cd_total_meetings = st.text_input("Total CD Meetings", placeholder="12")
        cd_percentages = st.text_input("CD Submittal %", value="25%, 50%, 75%, and 90%")
    
    st.markdown("---")
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**Task 150 - Construction Phase**")
        site_visits = st.text_input("Number of Site Visits", placeholder="6")
    
    with col2:
        st.markdown("**Task 160 - Record Drawings (Optional)**")
        include_record_drawings = st.checkbox("Include Record Drawings Task")
        if include_record_drawings:
            record_drawings_hours = st.text_input("Record Drawings Hours", placeholder="40")
        else:
            record_drawings_hours = ""

with tab5:
    st.subheader("Fee Structure")
    st.info("💰 Enter fees as numbers only (commas and $ symbols will be added automatically)")
    
    col1, col2 = st.columns(2)
    with col1:
        fee_sd = st.text_input("Task 110 - Schematic Design ($)", 
                              placeholder="25000",
                              help="Enter amount without commas")
        fee_dd = st.text_input("Task 120 - Design Development ($)", 
                              placeholder="45000")
        fee_cd = st.text_input("Task 130 - Construction Documents ($)", 
                              placeholder="85000")
    with col2:
        fee_bidding = st.text_input("Task 140 - Bidding ($)", 
                                   placeholder="5000")
        fee_construction = st.text_input("Task 150 - Construction Phase ($)", 
                                        placeholder="25000")
        if include_record_drawings:
            fee_record_drawings = st.text_input("Task 160 - Record Drawings ($)", 
                                               placeholder="10000")
        else:
            fee_record_drawings = ""
    
    # Calculate and display total
    st.markdown("---")
    fees_list = [fee_sd, fee_dd, fee_cd, fee_bidding, fee_construction]
    if include_record_drawings:
        fees_list.append(fee_record_drawings)
    
    total = 0
    valid_fees = True
    for fee in fees_list:
        is_valid, cleaned = validate_currency(fee)
        if not is_valid and fee:
            valid_fees = False
            st.error(f"Invalid fee amount: {fee}")
        elif cleaned:
            total += float(cleaned)
    
    if valid_fees:
        st.success(f"### Total Fee: **${total:,.0f}**" if total > 0 else "### Total Fee: **$___________**")
    
    st.markdown("---")
    st.subheader("Closure Information")
    col1, col2 = st.columns(2)
    with col1:
        invoice_email = st.text_input("Invoice Email", 
                                     placeholder="accounting@company.com",
                                     help="Primary email for invoices")
        if invoice_email and not validate_email(invoice_email):
            st.error("Please enter a valid email address")
            
        invoice_copy = st.text_input("CC Email", 
                                    placeholder="manager@company.com",
                                    help="Additional recipient for invoices")
        if invoice_copy and not validate_email(invoice_copy):
            st.error("Please enter a valid email address")
    with col2:
        project_manager = st.text_input("Project Manager", value="Clayton Scelzi")
        senior_vp = st.text_input("Senior Vice President", value="Scott W. Gilner, PE")

with tab6:
    st.subheader("Generate Proposal Document")
    
    st.info("📄 **Ready to Generate Your Proposal?**\n\nReview your inputs in the other tabs, then click below to generate your professional Word document with proper headers and footers on every page.")
    
    # Validation check
    required_fields = {
        'Client Contact': client_contact,
        'Company Name': company_name,
        'Project Name': project_name,
        'Project Address': project_address,
        'Project City': project_city,
    }
    
    missing_fields = [field for field, value in required_fields.items() if not value]
    
    if missing_fields:
        st.warning(f"⚠️ **Missing Required Information**\n\nPlease complete the following fields before generating:\n\n" + "\n".join([f"- {field}" for field in missing_fields]))
    
    # Generate button
    if st.button("🚀 Generate MEP Proposal", 
                type="primary", 
                use_container_width=True,
                disabled=bool(missing_fields)):
        
        # Collect all form data
        form_data = {
            'date': proposal_date.strftime("%B %d, %Y"),
            'client_title': client_title,
            'client_contact': client_contact,
            'company_name': company_name,
            'address1': address1,
            'address2': address2,
            'project_name': project_name,
            'project_address': project_address,
            'project_city': project_city,
            'project_state': project_state,
            'is_new_building': is_new_building,
            'is_renovation': is_renovation,
            'building_stories': building_stories,
            'total_area': total_area,
            'construction_phases': construction_phases,
            'separate_buildings': separate_buildings,
            'core_and_shell': core_and_shell,
            'leed_rating': leed_rating,
            'construction_budget': construction_budget,
            'unit_types': unit_types,
            'typical_floors': typical_floors,
            'retail_core_shell': retail_core_shell,
            'retail_electrical': retail_electrical if retail_core_shell else False,
            'retail_plumbing': retail_plumbing if retail_core_shell else False,
            'retail_food_beverage': retail_food_beverage if retail_core_shell else False,
            'retail_mechanical': retail_mechanical if retail_core_shell else False,
            'hvac_system': hvac_system,
            'hvac_residential_highrise': hvac_residential_highrise,
            'hvac_existing_reuse': hvac_existing_reuse,
            'outside_air_unit': outside_air_unit,
            'exhaust_system': exhaust_system,
            'parking_garage': parking_garage,
            'smoke_control': smoke_control,
            'elevator_hoistway': elevator_hoistway,
            'water_service': water_service,
            'roof_drainage': roof_drainage,
            'roof_storm_drain': roof_storm_drain,
            'parking_garage_drain': parking_garage_drain,
            'water_oil_separator': water_oil_separator,
            'sump_pump': sump_pump,
            'booster_pump': booster_pump,
            'sanitary_vent': sanitary_vent,
            'grease_waste': grease_waste,
            'natural_gas': natural_gas,
            'fuel_delivery': fuel_delivery,
            'civil_coordination': civil_coordination,
            'existing_electrical_renovation': existing_electrical_renovation,
            'power_receptacles': power_receptacles,
            'core_shell_electrical': core_shell_electrical,
            'lighting_coordination': lighting_coordination,
            'lightning_protection': lightning_protection,
            'emergency_generator': emergency_generator,
            'ev_charging': ev_charging,
            'ev_ready_spaces': ev_ready_spaces if ev_charging == "Included" else "",
            'ev_capable_spaces': ev_capable_spaces if ev_charging == "Included" else "",
            'fire_alarm': fire_alarm,
            'technology_design': technology_design,
            'fire_pump': fire_pump,
            'weekly_meetings': weekly_meetings,
            'revit_lod': revit_lod,
            'revit_coordination_hours': revit_coordination_hours,
            'sd_existing_survey': sd_existing_survey,
            'sd_site_visit_hours': sd_site_visit_hours if sd_existing_survey else "",
            'sd_weeks': sd_weeks,
            'sd_meeting_hours': sd_meeting_hours,
            'sd_total_meetings': sd_total_meetings,
            'dd_weeks': dd_weeks,
            'dd_meeting_hours': dd_meeting_hours,
            'dd_total_meetings': dd_total_meetings,
            'dd_rounds': dd_rounds,
            'cd_weeks': cd_weeks,
            'cd_meeting_hours': cd_meeting_hours,
            'cd_total_meetings': cd_total_meetings,
            'cd_percentages': cd_percentages,
            'site_visits': site_visits,
            'include_record_drawings': include_record_drawings,
            'record_drawings_hours': record_drawings_hours if include_record_drawings else "",
            'fee_sd': format_currency(fee_sd) if fee_sd else "XXX",
            'fee_dd': format_currency(fee_dd) if fee_dd else "XXX",
            'fee_cd': format_currency(fee_cd) if fee_cd else "XXX",
            'fee_bidding': format_currency(fee_bidding) if fee_bidding else "XXX",
            'fee_construction': format_currency(fee_construction) if fee_construction else "XXX",
            'fee_record_drawings': format_currency(fee_record_drawings) if include_record_drawings and fee_record_drawings else "",
            'invoice_email': invoice_email,
            'invoice_copy': invoice_copy,
            'project_manager': project_manager,
            'senior_vp': senior_vp,
        }
        
        progress_bar = st.progress(0, text="Initializing document...")
        
        try:
            progress_bar.progress(20, text="Creating proposal structure...")
            doc = create_proposal_document(form_data)
            
            progress_bar.progress(60, text="Formatting document...")
            buffer = BytesIO()
            doc.save(buffer)
            buffer.seek(0)
            
            progress_bar.progress(90, text="Finalizing...")
            filename = f"MEP_Proposal_{company_name.replace(' ', '_') if company_name else 'Draft'}_{datetime.now().strftime('%Y%m%d')}.docx"
            
            progress_bar.progress(100, text="Complete!")
            
            st.success("✅ **Document Generated Successfully!**\n\nYour professional MEP proposal is ready for download.")
            
            st.download_button(
                label="📥 Download Word Document",
                data=buffer.getvalue(),
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
                use_container_width=True
            )
            
        except Exception as e:
            progress_bar.empty()
            st.error(f"❌ **Error generating document:** {str(e)}")
            with st.expander("Show Error Details"):
                st.exception(e)

# Footer
st.markdown("---")
st.markdown("""
<div style="text-align: center; color: #666; font-size: 12px; padding: 20px;">
    <strong>MEP Proposal Generator v3.0</strong> | Kimley-Horn Engineering Services<br>
    Complete Template Implementation with Exact Formatting<br>
    <em>For support, contact your IT administrator</em>
</div>
""", unsafe_allow_html=True)