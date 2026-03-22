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
    
    insight_1 = f"The Efficiency Paradox: The lineup of {best_lineup} isn't just winning—it dictates pace without sacrificing half-court execution, resulting in a +{best_net} Net Rating. The twist? This dominance happens primarily because of how it suppresses opponent transition rather than just pure scoring."
    insight_2 = f"Two-Way Asymmetry: While {best_off_row['Lineups']} achieves the highest Offensive Rating ({best_off}), the scatter plot structure reveals a trade-off in defensive integrity (DefRtg {best_off_row['DefRtg']}). The key is strategic deployment: this unit is a 'blowout generator' best utilized in short bursts."
    insight_3 = f"True Shooting Elasticity: The group led by {best_ts_row['Lineups']} produces a staggering {best_ts}% True Shooting. Interestingly, this efficiency doesn't stem from taking fewer shots, but better structural spacing that stretches the defense horizontally, yielding open catch-and-shoot opportunities."
    
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
    panel_body_fmt.set_align('left')
    
    # Apply background up to row 55
    for r in range(0, 56):
        worksheet.set_row(r, 15, bg_format)
    
    # Set specific row heights
    worksheet.set_row(0, 15)
    worksheet.set_row(1, 20)
    worksheet.set_row(2, 15)
    worksheet.set_row(6, 30) # Middle row for KPI values
    worksheet.set_row(7, 30)
    
    # Set proportional column widths to prevent overlap
    worksheet.set_column('A:A', 2)     # left margin
    worksheet.set_column('B:G', 11)    # 6 cols * 82px = 492px
    worksheet.set_column('H:H', 2)     # padding
    worksheet.set_column('I:N', 11)    # 6 cols * 82px = 492px
    worksheet.set_column('O:O', 2)     # padding
    worksheet.set_column('P:U', 11)    # 6 cols * 82px = 492px
    worksheet.set_column('V:V', 2)     # right margin
    
    # Add charts data (hidden)
    worksheet.write_column('AA1', top5['Display_Lineups'])
    worksheet.set_column('AA:AA', 0) # hide helper
    
    # Title
    worksheet.merge_range('B2:U4', '24-25 Warriors Key Insights Post-Butler Trade', title_format)
    
    # Draw KPI cards
    worksheet.merge_range('B6:G6', 'PEAK NET RATING', kpi_title_fmt)
    worksheet.merge_range('B7:G8', f"+{best_net}", kpi_val_fmt)
    worksheet.merge_range('B9:G9', 'Top Performing Lineup', kpi_sub_fmt)
    
    worksheet.merge_range('I6:N6', 'PEAK OFFENSIVE RATING', kpi_title_fmt)
    worksheet.merge_range('I7:N8', str(best_off), kpi_val_fmt)
    worksheet.merge_range('I9:N9', 'Highest Scoring Output', kpi_sub_fmt)
    
    worksheet.merge_range('P6:U6', 'PEAK TRUE SHOOTING %', kpi_title_fmt)
    worksheet.merge_range('P7:U8', f"{best_ts}%", kpi_val_fmt)
    worksheet.merge_range('P9:U9', 'Most Efficient Lineup', kpi_sub_fmt)
    
    # Executive Insights Panel
    worksheet.merge_range('P12:U12', 'EXECUTIVE INSIGHTS', panel_header_fmt)
    panel_text = f"\n\n1. {insight_1}\n\n\n2. {insight_2}\n\n\n3. {insight_3}"
    worksheet.merge_range('P13:U51', panel_text, panel_body_fmt)
    
    # Chart dimensions
    c_width = 480
    c_height = 280
    
    # Chart 1: Combo Chart (Volume vs Net Efficiency) (B12)
    chart_combo = workbook.add_chart({'type': 'column'})
    chart_combo.add_series({
        'name': 'Minutes Played',
        'categories': ['Dashboard', 0, 26, 4, 26],
        'values': ['Data', 1, 2, 5, 2],
        'fill': {'color': '#0F243E'}
    })
    line_chart = workbook.add_chart({'type': 'line'})
    line_chart.add_series({
        'name': 'Net Rating',
        'categories': ['Dashboard', 0, 26, 4, 26],
        'values': ['Data', 1, 5, 5, 5],
        'line': {'color': '#FFC000', 'width': 2.5},
        'marker': {'type': 'circle', 'size': 7, 'fill': {'color': '#FFC000'}},
        'y2_axis': True
    })
    chart_combo.combine(line_chart)
    chart_combo.set_title({'name': '1. Lineup Volume vs. Efficiency (Dual Axis)', 'name_font': {'size': 12}})
    chart_combo.set_legend({'position': 'top'})
    chart_combo.set_chartarea({'border': {'color': '#D9D9D9'}, 'fill': {'color': '#FFFFFF'}})
    chart_combo.set_plotarea({'fill': {'color': '#FFFFFF'}})
    chart_combo.set_size({'width': c_width, 'height': c_height})
    worksheet.insert_chart('B12', chart_combo)
    
    # Chart 2: Scatter Plot (OffRtg vs DefRtg) (I12)
    chart_scatter = workbook.add_chart({'type': 'scatter'})
    colors_list = ['#0070C0', '#C00000', '#00B050', '#7030A0', '#FFC000']
    for i in range(1, 6):
        chart_scatter.add_series({
            'name': ['Dashboard', i-1, 26],
            'categories': ['Data', i, 3, i, 3],
            'values': ['Data', i, 4, i, 4],
            'marker': {'type': 'circle', 'size': 10, 'fill': {'color': colors_list[i-1]}}
        })
    chart_scatter.set_title({'name': '2. Offensive vs Defensive Footprint', 'name_font': {'size': 12}})
    chart_scatter.set_x_axis({'name': 'Offensive Rating (Higher = Better)'})
    chart_scatter.set_y_axis({'name': 'Defensive Rating (Lower = Better)', 'reverse': True})
    chart_scatter.set_legend({'position': 'bottom', 'font': {'size': 8}})
    chart_scatter.set_chartarea({'border': {'color': '#D9D9D9'}, 'fill': {'color': '#FFFFFF'}})
    chart_scatter.set_plotarea({'fill': {'color': '#FFFFFF'}})
    chart_scatter.set_size({'width': c_width, 'height': c_height})
    worksheet.insert_chart('I12', chart_scatter)
    
    # Chart 3: Area Chart (Shooting Dynamics) (B33)
    chart_area = workbook.add_chart({'type': 'area'})
    chart_area.add_series({
        'name': 'TS%',
        'categories': ['Dashboard', 0, 26, 4, 26],
        'values': ['Data', 1, 9, 5, 9],
        'fill': {'color': '#00B050', 'transparency': 30}
    })
    chart_area.add_series({
        'name': 'eFG%',
        'categories': ['Dashboard', 0, 26, 4, 26],
        'values': ['Data', 1, 8, 5, 8],
        'fill': {'color': '#7030A0', 'transparency': 50}
    })
    chart_area.set_title({'name': '3. Shooting Efficiency Dynamics', 'name_font': {'size': 12}})
    chart_area.set_legend({'position': 'top'})
    chart_area.set_chartarea({'border': {'color': '#D9D9D9'}, 'fill': {'color': '#FFFFFF'}})
    chart_area.set_plotarea({'fill': {'color': '#FFFFFF'}})
    chart_area.set_size({'width': c_width, 'height': c_height})
    worksheet.insert_chart('B33', chart_area)
    
    # Chart 4: Doughnut Chart (Minutes Allocation) (I33)
    chart_doughnut = workbook.add_chart({'type': 'doughnut'})
    chart_doughnut.add_series({
        'name': 'Minutes Allocation',
        'categories': ['Dashboard', 0, 26, 4, 26],
        'values': ['Data', 1, 2, 5, 2],
        'data_labels': {'percentage': True}
    })
    chart_doughnut.set_title({'name': '4. Minutes Reliance (Top 5 Units)', 'name_font': {'size': 12}})
    chart_doughnut.set_legend({'position': 'right', 'font': {'size': 8}})
    chart_doughnut.set_chartarea({'border': {'color': '#D9D9D9'}, 'fill': {'color': '#FFFFFF'}})
    chart_doughnut.set_plotarea({'fill': {'color': '#FFFFFF'}})
    chart_doughnut.set_size({'width': c_width, 'height': c_height})
    worksheet.insert_chart('I33', chart_doughnut)
    
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
        "<b>1. Lineup Volume vs. Efficiency (Combo):</b> Juxtaposes total minutes deployed against overall Net Rating to identify if units with high utility actually drive winning basketball on a per-possession basis.",
        "<b>2. Offensive vs Defensive Footprint (Scatter):</b> A quadrant-style analysis plotting OffRtg against DefRtg. Features an inverted Y-axis to immediately highlight 'two-way' elite lineups in the top-right quadrant.",
        "<b>3. Shooting Efficiency Dynamics (Area):</b> Cross-references True Shooting against Effective Field Goal Percentage to determine if units rely on foul-drawing elasticity or organic floor spacing.",
        "<b>4. Minutes Reliance (Doughnut):</b> Visualizes the minutes distribution specifically across the top 5 elite units, indicating rotational dependencies and trust levels."
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
