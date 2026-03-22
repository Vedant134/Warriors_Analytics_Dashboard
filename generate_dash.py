import pandas as pd
import xlsxwriter
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY
from reportlab.lib import colors

def main():
    # 1. Load data
    df = pd.read_csv('warriors_data_lineups.csv')

    # 2. Data Cleaning
    cols_to_drop = [
        'AST/TO', 'AST_Ratio', 'OREB_PCT', 'DREB_PCT', 'TO_Ratio', 
        'assist/turnover', 'assist ratio', 'OREB_percent', 'DREB_percent', 'turnover ratio'
    ]
    found_drop = [c for c in df.columns if any(drop_name.lower() in c.lower() for drop_name in cols_to_drop)]
    if found_drop:
        df = df.drop(columns=found_drop)
    
    # Filter MIN >= 20
    df = df[df['MIN'] >= 20].copy()
    
    # Top 5 by NetRtg
    top5 = df.sort_values(by='NetRtg', ascending=False).head(5).copy()
    top5 = top5.reset_index(drop=True)
    
    # Ensure Lineups are shorter for charts (e.g. replacing ' | ' with '\n')
    top5['Display_Lineups'] = top5['Lineups'].str.replace(' | ', '\n', regex=False)
    
    # 3. Generate Insights / KPIs
    best_lineup = top5.iloc[0]['Lineups']
    best_net = top5.iloc[0]['NetRtg']
    
    best_off_row = top5.loc[top5['OffRtg'].idxmax()]
    best_off = best_off_row['OffRtg']
    
    best_ts_row = top5.loc[top5['TS_PCT'].idxmax()]
    best_ts = best_ts_row['TS_PCT']
    
    insight_1 = f"Leading Lineup Strategy: The combination of {best_lineup} demonstrates superior structural balance, yielding a massive Net Rating of +{best_net} in high-leverage situations."
    insight_2 = f"Offensive Ceiling: Utilizing {best_off_row['Lineups']} maximizes scoring output, generating a peak Offensive Rating of {best_off}."
    insight_3 = f"Shooting Efficiency: {best_ts_row['Lineups']} provides the highest yield per possession, achieving a group-leading True Shooting Percentage of {best_ts}%."
    
    insights = [insight_1, insight_2, insight_3]
    
    # 4. Generate Excel Dashboard
    excel_path = 'warriors_lineup_dashboard.xlsx'
    writer = pd.ExcelWriter(excel_path, engine='xlsxwriter')
    
    # Write data to a sheet
    top5.drop(columns=['Display_Lineups']).to_excel(writer, sheet_name='Data', index=False)
    
    # Generate Dashboard sheet
    workbook = writer.book
    worksheet = workbook.add_worksheet('Dashboard')
    worksheet.hide_gridlines(2)
    
    # --- Formats ---
    # Background fill
    bg_format = workbook.add_format({'bg_color': '#F2F4F7'})
    
    # Title
    title_format = workbook.add_format({'bold': True, 'font_size': 20, 'bg_color': '#0F243E', 'font_color': 'white', 'align': 'center', 'valign': 'vcenter'})
    
    # KPI Formats
    kpi_title_fmt = workbook.add_format({'bold': True, 'font_size': 11, 'bg_color': '#FFFFFF', 'font_color': '#7F7F7F', 'align': 'center', 'valign': 'top', 'top': 1, 'left': 1, 'right': 1})
    kpi_val_fmt = workbook.add_format({'bold': True, 'font_size': 26, 'bg_color': '#FFFFFF', 'font_color': '#0070C0', 'align': 'center', 'valign': 'vcenter', 'left': 1, 'right': 1})
    kpi_sub_fmt = workbook.add_format({'italic': True, 'font_size': 9, 'bg_color': '#FFFFFF', 'font_color': '#A6A6A6', 'align': 'center', 'valign': 'bottom', 'bottom': 1, 'left': 1, 'right': 1})
    
    # Executive Insights Formats
    panel_header_fmt = workbook.add_format({'bold': True, 'font_size': 14, 'bg_color': '#1f497d', 'font_color': 'white', 'align': 'center', 'valign': 'vcenter', 'border': 1})
    panel_body_fmt = workbook.add_format({'font_size': 11, 'bg_color': '#FFFFFF', 'text_wrap': True, 'valign': 'top', 'border': 1})
    
    # Apply background
    for r in range(0, 50):
        worksheet.set_row(r, 15, bg_format)
    
    worksheet.set_row(0, 15)
    worksheet.set_row(1, 20)
    worksheet.set_row(2, 15)
    
    # Draw KPI cards
    worksheet.merge_range('B6:D6', 'PEAK NET RATING', kpi_title_fmt)
    worksheet.merge_range('B7:D8', f"+{best_net}", kpi_val_fmt)
    worksheet.merge_range('B9:D9', 'Top Performing Lineup', kpi_sub_fmt)
    
    worksheet.merge_range('F6:H6', 'PEAK OFFENSIVE RATING', kpi_title_fmt)
    worksheet.merge_range('F7:H8', str(best_off), kpi_val_fmt)
    worksheet.merge_range('F9:H9', 'Highest Scoring Output', kpi_sub_fmt)
    
    worksheet.merge_range('J6:L6', 'PEAK TRUE SHOOTING %', kpi_title_fmt)
    worksheet.merge_range('J7:L8', f"{best_ts}%", kpi_val_fmt)
    worksheet.merge_range('J9:L9', 'Most Efficient Lineup', kpi_sub_fmt)
    
    worksheet.set_row(6, 30) # Middle row for KPI values
    worksheet.set_row(7, 30)
    
    # Executive Insights Panel (O6:S42 -> no let's put it on the right from column N)
    worksheet.merge_range('A2:R4', '24-25 Warriors Key Insights Post-Butler Trade', title_format)
    worksheet.merge_range('N6:R6', 'EXECUTIVE INSIGHTS', panel_header_fmt)
    
    panel_text = f"\n\n1. {insight_1}\n\n\n2. {insight_2}\n\n\n3. {insight_3}"
    worksheet.merge_range('N7:R42', panel_text, panel_body_fmt)
    
    # Add charts
    worksheet.write_column('AA1', top5['Display_Lineups'])
    
    # Chart 1: Net Rating (A11)
    chart1 = workbook.add_chart({'type': 'bar'})
    chart1.add_series({
        'name': 'Net Rating',
        'categories': ['Dashboard', 0, 26, 4, 26],
        'values': ['Data', 1, 5, 5, 5],
        'fill': {'color': '#0070C0'},
        'data_labels': {'value': True, 'position': 'outside_end'}
    })
    chart1.set_title({'name': 'Net Rating Comparison'})
    chart1.set_legend({'none': True})
    chart1.set_chartarea({'border': {'color': '#D9D9D9'}})
    chart1.set_size({'width': 380, 'height': 250})
    worksheet.insert_chart('A12', chart1)
    
    # Chart 2: Off vs Def (G11)
    chart2 = workbook.add_chart({'type': 'column'})
    chart2.add_series({
        'name': 'OffRtg',
        'categories': ['Dashboard', 0, 26, 4, 26],
        'values': ['Data', 1, 3, 5, 3],
        'fill': {'color': '#FFC000'}
    })
    chart2.add_series({
        'name': 'DefRtg',
        'categories': ['Dashboard', 0, 26, 4, 26],
        'values': ['Data', 1, 4, 5, 4],
        'fill': {'color': '#C00000'}
    })
    chart2.set_title({'name': 'Offensive vs Defensive Rating'})
    chart2.set_legend({'position': 'bottom'})
    chart2.set_chartarea({'border': {'color': '#D9D9D9'}})
    chart2.set_size({'width': 380, 'height': 250})
    worksheet.insert_chart('G12', chart2)
    
    # Chart 3: TS% (A27)
    chart3 = workbook.add_chart({'type': 'column'})
    chart3.add_series({
        'name': 'TS%',
        'categories': ['Dashboard', 0, 26, 4, 26],
        'values': ['Data', 1, 9, 5, 9],
        'fill': {'color': '#00B050'}
    })
    chart3.set_title({'name': 'True Shooting Percentage (TS%)'})
    chart3.set_legend({'none': True})
    chart3.set_chartarea({'border': {'color': '#D9D9D9'}})
    chart3.set_size({'width': 380, 'height': 250})
    worksheet.insert_chart('A29', chart3)
    
    # Chart 4: eFG% (G27)
    chart4 = workbook.add_chart({'type': 'bar'})
    chart4.add_series({
        'name': 'eFG%',
        'categories': ['Dashboard', 0, 26, 4, 26],
        'values': ['Data', 1, 8, 5, 8],
        'fill': {'color': '#7030A0'}
    })
    chart4.set_title({'name': 'Effective Field Goal Percentage (eFG%)'})
    chart4.set_legend({'none': True})
    chart4.set_chartarea({'border': {'color': '#D9D9D9'}})
    chart4.set_size({'width': 380, 'height': 250})
    worksheet.insert_chart('G29', chart4)
    
    # Resize cols for consistent spacing
    worksheet.set_column('A:T', 7)
    worksheet.set_column('AA:AA', 0) # hide helper
    
    writer.close()
    
    # 5. Generate Professional PDF
    pdf_path = 'project_summary.pdf'
    doc = SimpleDocTemplate(pdf_path, pagesize=letter)
    
    styles = getSampleStyleSheet()
    
    # Custom Professional Styles
    title_style = ParagraphStyle(
        name='ExecTitle',
        parent=styles['Heading1'],
        fontSize=20,
        textColor=colors.HexColor('#0F243E'),
        alignment=TA_CENTER,
        spaceAfter=20
    )
    
    heading_style = ParagraphStyle(
        name='ExecHeading',
        parent=styles['Heading2'],
        fontSize=14,
        textColor=colors.HexColor('#1f497d'),
        spaceBefore=15,
        spaceAfter=10
    )
    
    normal_style = ParagraphStyle(
        name='ExecNormal',
        parent=styles['Normal'],
        fontSize=11,
        leading=16,
        alignment=TA_JUSTIFY
    )
    
    elements = []
    
    # Header
    elements.append(Paragraph("24-25 Warriors Key Insights Post-Butler Trade", title_style))
    elements.append(Paragraph("Executive Summary", heading_style))
    elements.append(Paragraph("This project provides a comprehensive analysis of Golden State Warriors lineup data to identify the highest performing personnel groupings following the Butler trade. Designed for executive review, the findings highlight strategic advantages, structural balance, and efficiency metrics that drive team success.", normal_style))
    
    # KPIs
    elements.append(Paragraph("Key Performance Indicators (KPIs)", heading_style))
    kpi_data = [
        ["Peak Net Rating", "Peak Offensive Rating", "Peak True Shooting %"],
        [f"+{best_net}", f"{best_off}", f"{best_ts}%"]
    ]
    t = Table(kpi_data, colWidths=[150, 150, 150])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#1f497d')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 10),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 10),
        ('BACKGROUND', (0, 1), (-1, 1), colors.HexColor('#F2F4F7')),
        ('TEXTCOLOR', (0, 1), (-1, 1), colors.HexColor('#0070C0')),
        ('FONTNAME', (0, 1), (-1, 1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 1), (-1, 1), 16),
        ('TOPPADDING', (0, 1), (-1, 1), 15),
        ('BOTTOMPADDING', (0, 1), (-1, 1), 15),
        ('GRID', (0,0), (-1,-1), 1, colors.HexColor('#D9D9D9'))
    ]))
    elements.append(t)
    elements.append(Spacer(1, 15))
    
    # Insights
    elements.append(Paragraph("Strategic Insights", heading_style))
    for ins in insights:
        elements.append(Paragraph(f"• <b>{ins.split(':')[0]}:</b>{ins.split(':')[1]}", normal_style))
        elements.append(Spacer(1, 8))
        
    # Visual Breakdown
    elements.append(Paragraph("Dashboard Visual Overview", heading_style))
    v_desc = [
        "<b>1. Net Rating Comparison:</b> Establishes precisely which unit combinations yield the highest point differential.",
        "<b>2. Offensive vs Defensive Rating:</b> Identifies two-way versatility by isolating scoring output from defensive fortitude.",
        "<b>3. True Shooting Percentage:</b> Assesses the overarching yield per possession including three-pointers and free throws.",
        "<b>4. Effective Field Goal Percentage:</b> Measures raw shooting efficiency distinctly from the floor."
    ]
    for v in v_desc:
        elements.append(Paragraph(v, normal_style))
        elements.append(Spacer(1, 4))
        
    # Data Processing Notes
    elements.append(Paragraph("Methodology & Data Hygiene", heading_style))
    elements.append(Paragraph("For statistical reliability, analysis was strictly confined to lineups registering at least 20 minutes (MIN >= 20). Extraneous metrics were purged during the ETL phase to streamline dashboard focus on high-impact macro ratings.", normal_style))
    
    doc.build(elements)
    print("Dashboard and PDF formatted for executive interview successfully!")

if __name__ == '__main__':
    main()
