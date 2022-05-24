#SCOPE OF REPORT TABLE (PAGE 15 OF SAMPLE PDF REPORT)
data=[['Base Year for Estimation','2019'],
['Forecast Period','2020–2027'],
['Forecast Unit','Value (USD Million)'],
['Regionals & Countries Covered',"""region/regions selected by user and their countries"""],
['By Type',"""by type segmentation entered by user"""],
['Application','Municipal, Industrial and others'],
['Companies covered',"""companies entered by user"""]]

t=Table(data)
t.setStyle(TableStyle([('BACKGROUND',(1,1),(-2,-2),colors.white),
('TEXTCOLOR',(0,0),(-1,-1),colors.black),
('BACKGROUND', (0, 0), (0, 10), colors.lightblue),
('FONTSIZE',(0,0),(-1,-1), 12),
('INNERGRID', (0,0), (-1,-1), 0.25, colors.black),
('GRID',(0,0),(-1,-1),0.5,colors.black),
('ALIGN',(0,0),(0,6),'CENTER'),
('VALIGN',(0,0),(0,6),'MIDDLE')]))
story.append(t)
story.append(PageBreak())

#MARKET SNAPSHOT TABLE (PG 20 OF SAMPLE REPORT)
#includes global market share in general and for all the types, regions and applications

data1=[['Parameter','2019','2027'],
['Global  Market Size (USD Billion)',"""for 2019""","""for 2027"""],
['By Application',"Municipal: $",'Municipal: $'],
['','Industrial: $','Industrial: $'],
['Geographic Share',"""for all regions chosen by user in 2019""","""or all regions chosen by user in 2027"""],
['By Country',"""for all countries chosen by user in 2019""","""for all countries chosen by user in 2027"""],
['CAGR %'],
['Drivers'],
['Restraints']]

t1=Table(data1)
t1.setStyle(TableStyle([('BACKGROUND',(1,1),(-2,-2),colors.white),
('TEXTCOLOR',(0,0),(-1,-1),colors.black),
('BACKGROUND', (0, 0), (0, 10), colors.lightblue),
('FONTSIZE',(0,0),(-1,-1), 12),
('INNERGRID', (0,0), (-1,-1), 0.25, colors.black),
('GRID',(0,0),(-1,-1),0.5,colors.black),
('ALIGN',(0,0),(0,6),'CENTER'),
('VALIGN',(0,0),(0,6),'MIDDLE')]))
story.append(t1)
