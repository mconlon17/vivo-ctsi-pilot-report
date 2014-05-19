#!/usr/bin/env/python
"""
    publication_report.py -- Given a list of people, produce a Word document
    listing publications for each, most recent first

    Version 0.1 MC 2014-03-25
    --  Works as expected
"""

__author__ = "Michael Conlon"
__copyright__ = "Copyright 2014, University of Florida"
__license__ = "BSD 3-Clause license"
__version__ = "0.1"

from vivotools import read_csv
from vivotools import get_person
from vivotools import find_vivo_uri
from vivotools import string_from_document
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

YEAR = None  # Pubs must be on or after this year
REPORT_TITLE = 'Publication Report'

print str(datetime.datetime.now())
ufids = read_csv("weller2.csv")

previousPerson = None

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

section.Header.append(REPORT_TITLE)
p = Paragraph(ss.ParagraphStyles.Normal)
p.append(PAGE_NUMBER, ' of ', TOTAL_PAGES)
section.Footer.append(p)

p = Paragraph(ss.ParagraphStyles.Title)
p.append(REPORT_TITLE)
section.append(p)


for row in ufids.values():
    ufid = row['ufid']
    if ufid.startswith('http'):
        uri = ufid
        ufid = None
    else:
        uri = find_vivo_uri("ufVivo:ufid", ufid)
    if uri is None:
        print ufid,' not found'
        person = {}
        person['display_name'] = 'UFID '+ ufid + ' not found'
        p = Paragraph(ss.ParagraphStyles.Heading1)
        p. append(person['display_name'])
        section.append(p)
        previousPerson = person
        continue
    person = get_person(uri, get_publications=True)
    print uri
    print person
    print person['display_name'], len(person['publications']), "publications"

    if previousPerson != person:
        p = Paragraph(ss.ParagraphStyles.Heading1)
        p. append(person['display_name'])
        section.append(p)
        previousPerson = person

    if len(person['publications']) > 0:
        for pub in person['publications']:
            if 'date' not in pub:
                pub['date'] = {'year': None}
        for publication in sorted(person['publications'],
                                  key=lambda pub: pub['date']['year'], reverse=True):
            if publication['date']['year'] is not None and (YEAR is None or  \
                                int(publication['date']['year']) >= int(YEAR)):
                para_props = ParagraphPropertySet()
                para_props.SetFirstLineIndent(TabPropertySet.DEFAULT_WIDTH*-1)
                para_props.SetLeftIndent(TabPropertySet.DEFAULT_WIDTH*2)
                p = Paragraph(ss.ParagraphStyles.Normal, para_props)
                p.append(string_from_document(publication))
                section.append(p)

print str(datetime.datetime.now())
Renderer().Write(doc, file("publication_report.rtf", "w"))
