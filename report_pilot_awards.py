#!/usr/bin/env/python
"""
    report_pilot_awards.py -- Given a list of people involved with pilot
    awards, list their publications and grants in an RTF file suitable for
    use in Word

    Version 0.0 MC 2013-12-26
    --  Getting started with rtf-ng
    Version 0.1 MC 2013-12-28
    --  works as expected
    Version 0.2 MC 2014-02-18
    --  source formatting, pylint
"""

__author__ = "Michael Conlon"
__copyright__ = "Copyright 2014, University of Florida"
__license__ = "BSD 3-Clause license"
__version__ = "0.2"

import vivotools as vt
import csv
import datetime

# I wish rtf-ng was more organized and we didn't need all the imports below,
# but it isn't and we do.

from rtfng.Renderer import Renderer
from rtfng.Elements import Document, PAGE_NUMBER, TOTAL_PAGES
from rtfng.Styles import TextStyle, ParagraphStyle
from rtfng.document.section import Section
from rtfng.document.paragraph import Paragraph
from rtfng.PropertySets import TabPropertySet, TextPropertySet, \
     ParagraphPropertySet

# Start here

print str(datetime.datetime.now())
print "Creating UFID dictionary"
ufid_dictionary = vt.make_ufid_dictionary()
print "Dictionary has", len(ufid_dictionary), "people"
print str(datetime.datetime.now())
csvReader = csv.reader(open('Award Data Collection Effort Sample.csv', 'rb'),
                       delimiter='|')
n = 0

previousYear = None
previousRFA = None
report_title = 'Sample CTSI Pilot Award Report'

doc = Document()
ss = doc.StyleSheet

# Improve the style sheet

ps = ParagraphStyle('Title',
                    TextStyle(TextPropertySet(ss.Fonts.Arial, 44)).Copy(),
                    ParagraphPropertySet(space_before=60, space_after=60))
ss.ParagraphStyles.append(ps)
ps = ParagraphStyle('Heading 3',
                    TextStyle(TextPropertySet(ss.Fonts.Arial, 22)).Copy(),
                    ParagraphPropertySet(space_before=60, space_after=60))
ss.ParagraphStyles.append(ps)
ps = ParagraphStyle('Heading 4',
                    TextStyle(TextPropertySet(ss.Fonts.Arial, 22)).Copy(),
                    ParagraphPropertySet(space_before=60, space_after=60))
ss.ParagraphStyles.append(ps)

section = Section()
doc.Sections.append(section)

section.FirstHeader.append(' ')
p = Paragraph(ss.ParagraphStyles.Normal)
p.append(PAGE_NUMBER, ' of ', TOTAL_PAGES)
section.FirstFooter.append(p)

section.Header.append(report_title)
p = Paragraph(ss.ParagraphStyles.Normal)
p.append(PAGE_NUMBER, ' of ', TOTAL_PAGES)
section.Footer.append(p)

p = Paragraph(ss.ParagraphStyles.Title)
p.append(report_title)
section.append(p)


for row in csvReader:
    n = n + 1
    if n == 1:
        continue
    i = 0
    for r in row:
        row[i] = r.strip()
        i = i + 1

    [year, rfa, title, name, ufid, email, role, status, program] = row

    if previousYear != year:
        p = Paragraph(ss.ParagraphStyles.Heading1)
        p. append(year)
        section.append(p)
        previousYear = year
        p = Paragraph(ss.ParagraphStyles.Heading2)
        p. append(rfa)
        section.append(p)
        previousRFA = rfa
    elif previousRFA != rfa:
        p = Paragraph(ss.ParagraphStyles.Heading2)
        p. append(rfa)
        section.append(p)
        previousRFA = rfa

    [found_person, person_uri] = vt.find_person(ufid, ufid_dictionary)
    if not found_person:
        print '\n', n, name, "not Found", "ufid", ufid
        continue
    person = vt.get_person(person_uri, get_publications=True, get_grants=True)

    p = Paragraph(ss.ParagraphStyles.Heading3)
    p. append(title)
    section.append(p)
    p = Paragraph(ss.ParagraphStyles.Heading4)
    p. append(name)
    section.append(p)

    if len(person['grants']) > 0:
        for grant in person['grants']:
            try:
                print 'grant',\
                      int(grant['start_date']['date']['year']), int(year)
            except:
                continue
            if int(grant['start_date']['date']['year']) >= int(year):
                para_props = ParagraphPropertySet()
                para_props.SetFirstLineIndent(TabPropertySet.DEFAULT_WIDTH*-1)
                para_props.SetLeftIndent(TabPropertySet.DEFAULT_WIDTH*2)
                p = Paragraph(ss.ParagraphStyles.Normal, para_props)
                p.append(vt.string_from_grant(grant))
                section.append(p)

    if len(person['publications']) > 0:
        for publication in person['publications']:
            try:
                print 'pub', int(publication['date']['year']), int(year)
            except:
                continue
            if int(publication['date']['year']) >= int(year):
                para_props = ParagraphPropertySet()
                para_props.SetFirstLineIndent(TabPropertySet.DEFAULT_WIDTH*-1)
                para_props.SetLeftIndent(TabPropertySet.DEFAULT_WIDTH*2)
                p = Paragraph(ss.ParagraphStyles.Normal, para_props)
                p.append(vt.string_from_document(publication))
                section.append(p)

print str(datetime.datetime.now())
Renderer().Write(doc, file("report.rtf", "w"))
