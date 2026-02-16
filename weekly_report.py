#!/usr/bin/env python3
"""
Weekly Analysis Report Generator
Generates: Weekly_Analysis_Report.pdf with charts
"""

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime
import os
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER

sns.set_style("whitegrid")
plt.rcParams['figure.figsize'] = (12, 8)

# All store names for tracking
ALL_STORES = [
    'Jenkins', 'Neon', 'Harlan 1', 'Harlan 2', 'Hyden', 'PMM', 'Isom', 
    'Whitesburg', 'Hazard 2', 'Ermine', 'Hindman 2', 
    'Hindman 1', 'Martin', 'Jackson', 'Hazard 3', 'Dryfork', 'Pound', 
    'Catnip (Nicholasville)', 'Marrowbone', 'Elkhorn City', 'Chloe', 
    'Caney', 'Belfrey', 'Phelps', 'Virgie', 'Harold', 'Allen', 'Goody', 
    'Zebulon', 'Pikeville', 'South', 'North', 'Prestonsburg 1', 'Ivel', 
    'Justiceville', 'Salyersville', 'Grundy', 'West Liberty', 
    'Prestonsburg 2', 'Prestonsburg 3'
]

def load_and_process_data():
    """Load and process ticket data"""
    print("\nLoading data...")
    active = pd.read_csv('active_tickets.csv')
    closed = pd.read_csv('closed_tickets.csv')
    
    # Filter to IT Helpdesk team only
    if 'Team' in active.columns:
        active = active[active['Team'] == 'IT Helpdesk']
    if 'Team' in closed.columns:
        closed = closed[closed['Team'] == 'IT Helpdesk']
    
    print(f"✓ Filtered to IT Helpdesk team only")
    
    # Filter closed to last 7 days using Last Modified Date
    closed['Last Modified Date'] = pd.to_datetime(closed['Last Modified Date'], errors='coerce', utc=True)
    seven_days_ago = pd.Timestamp.now(tz='UTC') - pd.Timedelta(days=7)
    closed = closed[closed['Last Modified Date'] >= seven_days_ago]
    
    # Remove timezone info
    for col in closed.columns:
        if pd.api.types.is_datetime64_any_dtype(closed[col]):
            closed[col] = closed[col].dt.tz_localize(None)
    
    for col in active.columns:
        if pd.api.types.is_datetime64_any_dtype(active[col]):
            active[col] = active[col].dt.tz_localize(None)
    
    # Treat "Resolved" status as closed too
    active_copy = active.copy()
    resolved = active_copy[active_copy['Status'].str.lower().str.contains('resolved', na=False)]
    active = active_copy[~active_copy['Status'].str.lower().str.contains('resolved', na=False)]
    
    if len(resolved) > 0:
        closed = pd.concat([closed, resolved], ignore_index=True)
    
    # Add Status_Type
    active['Status_Type'] = 'Active'
    closed['Status_Type'] = 'Closed'
    
    all_tickets = pd.concat([active, closed], ignore_index=True)
    
    print(f"✓ Active tickets: {len(active)}")
    print(f"✓ Closed/Resolved tickets (last 7 days): {len(closed)}")
    
    return all_tickets

def generate_weekly_analysis_pdf(all_tickets):
    """Generate Weekly Analysis Report as PDF with embedded charts"""
    print("\n" + "=" * 50)
    print("WEEKLY ANALYSIS REPORT")
    print("=" * 50)
    
    # Only use closed tickets from last 7 days
    analysis_tickets = all_tickets[all_tickets['Status_Type'] == 'Closed'].copy()
    
    # Exclude Jacob Sexton
    analysis_tickets = analysis_tickets[analysis_tickets['Requester'].str.lower() != 'jacob sexton'].copy()
    
    # Combine Isom Deli with Isom
    analysis_tickets['Requester'] = analysis_tickets['Requester'].str.replace('Isom Deli', 'Isom', case=False, regex=False)
    
    # Create charts directory
    os.makedirs('charts', exist_ok=True)
    
    # ========== GENERATE CHARTS ==========
    print("\n📊 Generating charts...")
    
    # Chart 1: Top Stores Pie Chart - Only include actual stores from ALL_STORES list
    # Filter to only tickets from stores in the list
    store_tickets = analysis_tickets[
        analysis_tickets['Requester'].apply(
            lambda x: any(store.lower() in str(x).lower() for store in ALL_STORES)
        )
    ]
    
    store_counts = store_tickets.groupby('Requester').size().sort_values(ascending=False).head(10)
    
    if len(store_counts) > 0:
        plt.figure(figsize=(10, 8))
        colors_pie = sns.color_palette("Set3", len(store_counts))
        plt.pie(store_counts.values, labels=store_counts.index, autopct='%1.1f%%',
                colors=colors_pie, startangle=90)
        plt.title('Top 10 Stores - Closed Tickets (Last 7 Days)', fontsize=16, fontweight='bold')
        plt.tight_layout()
        plt.savefig('charts/stores_pie.png', dpi=150, bbox_inches='tight')
        plt.close()
    
    # Chart 2: Top Stores Bar Chart - Only include actual stores from ALL_STORES list
    store_counts_bar = store_tickets.groupby('Requester').size().sort_values(ascending=False).head(15)
    
    if len(store_counts_bar) > 0:
        plt.figure(figsize=(12, 8))
        plt.barh(range(len(store_counts_bar)), store_counts_bar.values, color='#2ecc71')
        plt.yticks(range(len(store_counts_bar)), store_counts_bar.index)
        plt.xlabel('Number of Closed Tickets', fontsize=12)
        plt.ylabel('Store', fontsize=12)
        plt.title('Top 15 Stores by Closed Tickets (Last 7 Days)', fontsize=16, fontweight='bold')
        plt.gca().invert_yaxis()
        # Force integer x-axis
        ax = plt.gca()
        ax.xaxis.set_major_locator(plt.MaxNLocator(integer=True))
        plt.tight_layout()
        plt.savefig('charts/top_stores.png', dpi=150, bbox_inches='tight')
        plt.close()
    
    # Chart 3: Techs Bar Chart - Only show specific techs
    assigned_tickets = analysis_tickets[analysis_tickets['Assignee'].notna() & (analysis_tickets['Assignee'] != '')]
    
    # Filter to only the 4 main techs
    tech_filter = assigned_tickets['Assignee'].str.contains('Jacob|Richard|Jon|Rick', case=False, na=False)
    filtered_techs = assigned_tickets[tech_filter]
    
    tech_counts = filtered_techs.groupby('Assignee').size().sort_values(ascending=False)
    
    if len(tech_counts) > 0:
        plt.figure(figsize=(10, 6))
        plt.barh(range(len(tech_counts)), tech_counts.values, color='#3498db')
        plt.yticks(range(len(tech_counts)), tech_counts.index)
        plt.xlabel('Number of Closed Tickets', fontsize=12)
        plt.ylabel('Assignee', fontsize=12)
        plt.title('Closed Tickets by Assignee (Last 7 Days)', fontsize=16, fontweight='bold')
        plt.gca().invert_yaxis()
        # Force integer x-axis
        ax = plt.gca()
        ax.xaxis.set_major_locator(plt.MaxNLocator(integer=True))
        plt.tight_layout()
        plt.savefig('charts/techs.png', dpi=150, bbox_inches='tight')
        plt.close()
    
    # ========== BUILD PDF ==========
    print("\n📄 Building PDF...")
    
    pdf_file = 'Weekly_Analysis_Report.pdf'
    doc = SimpleDocTemplate(pdf_file, pagesize=letter,
                           rightMargin=0.5*inch, leftMargin=0.5*inch,
                           topMargin=0.5*inch, bottomMargin=0.5*inch)
    
    elements = []
    styles = getSampleStyleSheet()
    
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=20,
        textColor=colors.HexColor('#2c3e50'),
        spaceAfter=20,
        alignment=TA_CENTER
    )
    
    heading_style = ParagraphStyle(
        'CustomHeading',
        parent=styles['Heading2'],
        fontSize=14,
        textColor=colors.HexColor('#34495e'),
        spaceAfter=10
    )
    
    # Title Page
    elements.append(Paragraph("Weekly Analysis Report", title_style))
    elements.append(Paragraph(f"Closed Tickets - Last 7 Days", styles['Normal']))
    elements.append(Paragraph(f"Generated: {datetime.now().strftime('%B %d, %Y')}", styles['Normal']))
    elements.append(Spacer(1, 0.3*inch))
    
    # Add Stores Pie Chart
    elements.append(Paragraph("Top Stores with Closed Tickets", heading_style))
    if os.path.exists('charts/stores_pie.png'):
        img = Image('charts/stores_pie.png', width=6*inch, height=4.8*inch)
        elements.append(img)
        elements.append(Spacer(1, 0.2*inch))
    
    # Add Top Stores Chart
    elements.append(Paragraph("Top Stores with Closed Tickets", heading_style))
    if os.path.exists('charts/top_stores.png'):
        img = Image('charts/top_stores.png', width=6.5*inch, height=4.3*inch)
        elements.append(img)
        elements.append(Spacer(1, 0.2*inch))
    
    elements.append(PageBreak())
    
    # Add Techs Chart
    elements.append(Paragraph("Closed Tickets by Assignee", heading_style))
    if os.path.exists('charts/techs.png'):
        img = Image('charts/techs.png', width=6*inch, height=3.6*inch)
        elements.append(img)
        elements.append(Spacer(1, 0.3*inch))
    
    # Complete Store List
    elements.append(Paragraph("Complete Store Report", heading_style))
    elements.append(Paragraph("All Stores - Closed Tickets (Last 7 Days)", styles['Normal']))
    elements.append(Spacer(1, 0.2*inch))
    
    # Build complete store list with counts
    store_report_data = [['Store', 'Closed Tickets']]
    store_counts_list = []
    
    for store in ALL_STORES:
        count = len(analysis_tickets[analysis_tickets['Requester'].str.contains(store, case=False, na=False)])
        store_counts_list.append((store, count))
    
    # Sort by count descending
    store_counts_list.sort(key=lambda x: x[1], reverse=True)
    
    for store, count in store_counts_list:
        store_report_data.append([store, str(count)])
    
    store_table = Table(store_report_data, colWidths=[4*inch, 1.5*inch])
    store_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#3498db')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('FONTSIZE', (0, 1), (-1, -1), 9),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.lightgrey])
    ]))
    
    elements.append(store_table)
    
    # Build PDF
    doc.build(elements)
    
    print(f"✓ Saved: {pdf_file}")

def main():
    print("=" * 50)
    print("WEEKLY ANALYSIS REPORT GENERATOR")
    print("=" * 50)
    
    # Load and process data
    all_tickets = load_and_process_data()
    
    # Generate PDF
    generate_weekly_analysis_pdf(all_tickets)
    
    print("\n" + "=" * 50)
    print("✅ REPORT GENERATED!")
    print("=" * 50)
    print("\nFile created: Weekly_Analysis_Report.pdf")
    print("\n")

if __name__ == "__main__":
    main()
