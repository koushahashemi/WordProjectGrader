"""Instructor/student feedback, distinct from the XML diagnostic audit trail."""
import re
from grader import CRITERIA

REFERENCES={
 CRITERIA[0]:'Project Instruction 3; Brand Style Guide, Typography (section 3).',
 CRITERIA[1]:'Project Instruction 4; Brand Style Guide, Normal (section 3) and optional callouts (section 7).',
 CRITERIA[2]:'Project Instruction 5; Brand Style Guide, Typography (section 3).',
 CRITERIA[3]:'Project Instruction 6; Brand Style Guide, Table of Contents (section 4).',
 CRITERIA[4]:'Project Instruction 7; Brand Style Guide, Captions / Notes (section 3).',
 CRITERIA[5]:'Project Instruction 8; Brand Style Guide, List of Figures (section 5).',
 CRITERIA[6]:'Project Instruction 9; Brand Style Guide, Bullets and Numbering (section 6).',
 CRITERIA[7]:'Project Instructions 10–11 (APA 6 and updating the first source).',
 CRITERIA[8]:'Project Instruction 13 (Word-generated references on a new page).',
 CRITERIA[9]:'Project Instruction 12; Brand Style Guide, Header & Footer (section 9).',
 CRITERIA[10]:'Project Instruction 13 (new section, independent footer and A/B/C numbering).',
 CRITERIA[11]:'Brand Style Guide as a whole: typography (3), contents/figures lists (4–5), lists (6), margins (8), and header/footer (9).',
}


def explanation(criterion,check):
    name=check['name'];expected=check['expected'];found=check['found']
    specific={
      'Separate references section':'The bibliography does not have a separate Word section. Insert a Next Page section break before it; a page break alone does not create a section.',
      'References footer unlinked':'The bibliography footer is still linked to the previous section. Turn off “Link to Previous” for that footer so its numbering can be changed independently.',
      'References letter numbering':'The bibliography page numbering is not set to A, B, C. Choose uppercase letters in Format Page Numbers for the references section.',
      'References numbering starts at A':'The references section does not explicitly restart at A. Set its starting page number to A.',
      'New section starts next page':'The references section is not configured to start on the next page. Use a Next Page section break before the bibliography.',
      'Main section decimal numbering':'The main document’s numbering has changed from ordinary numbers. Keep its numbering unchanged when formatting the bibliography section.',
      'Bibliography after last section break':'The bibliography is not located after the final section break. Place it in the new references section.',
      'Title-page number hidden':'The title page’s page number is not configured to be hidden. Use Different First Page with a blank first-page footer, or a footer-free cover followed by page 2.',
      'Main page numbering starts at 1':'The title page is not configured to count as page 1. Hide its printed number and let the following page display 2, rather than restarting at 1.',
      'Footer right aligned':'The footer content is not right-aligned. Right-align the page number and document title, or use a right-aligned tab stop.',
      'Document Title field':'The cover title is not linked to the document’s Title property. Insert a Document Title field or linked Title content control.',
      'Student information lines':'The cover page is missing one or more required lines: document title, name, student number and set.',
      'Blank lines between information':'Leave a blank line between the title, name, student number and set on the cover.',
      'List title style':'The figures-list heading needs Heading 1 or the accepted TOCHeading style.',
      'List of Figures title':'Add a heading for the figures list. “List of Figures” and “Table of Figures” are both accepted.',
      'Caption above image':'A caption is not directly above its figure. Move the caption above the corresponding picture.',
      'Caption style applied':'A figure caption is not using the Caption style (or a style based on it). Apply the modified Caption style.',
      'Caption text present':'There are no recognizable figure captions to evaluate. Insert Word captions for the four pictures.',
      'Four caption fields':'Not all four figures have real Word captions. Use Insert Caption for each figure; typing a label alone does not create a caption field.',
      'Four figures present':'One or more of the four original figures is missing from the document.',
      'Original caption text':'A required caption is missing or substantially changed. Expected title: '+expected+'. Punctuation and minor spelling differences are accepted.',
      'Applied list numbering':'A required list item is missing or is not part of a real Word list. Apply a bulleted/numbered list style rather than typing the marker.',
      'Bullet color':'A bullet marker has the wrong color. Set the marker to Soft Coral (#FF6F61); the list text should remain Dark Gray.',
      'Bullet shape':'A list uses a different bullet shape. Use a solid circle or square.',
      'Numbering format':'A numbered list uses a different numbering scheme. Use decimal numbering as specified in the guide.',
      'APA 6 bibliography setting':'The document’s citation style is not APA 6. Select APA Sixth Edition in Word’s bibliography settings.',
      'Citation fields':'The document has no recognizable Word citation fields. Insert citations using Word’s citation tools.',
      'Citation sources resolved':'At least one citation cannot be matched to a saved source. Check its source in Manage Sources.',
      'Word bibliography field':'The references are not generated using Word’s Bibliography tool. Insert a Word bibliography in the references section.',
      'Bibliography saved entries':'The bibliography field has no saved reference entries. Update the bibliography after checking the document’s sources.',
      'Body paragraphs present':'The expected assignment body text could not be identified. Check that the original Bears text is retained.',
      'Body text coverage':'Some original body text could not be matched. Check that assignment paragraphs have not been removed or substantially rewritten.',
      'Body paragraph style':'A body paragraph is not using the Normal style. Apply the modified Normal style to body text.',
      'Title bottom border':'The title’s bottom border does not match the guide. Use a 1-point Dark Teal bottom border.',
      'Heading style application':'A required heading is missing or uses the wrong heading level. Apply the appropriate modified Heading 1, 2 or 3 style.',
      'Generated entry styles applied':'The generated contents/list entries are not using their expected Word styles. Format the generated entry styles to match the guide.',
      'Leader dots':'The generated list is missing dot leaders between entry text and page numbers. Enable dot leaders in its tab settings.',
      'Field has saved results':'The generated list has no saved entries. Update it in Word and save the document.',
      'Manual entry formatting':'The manually created list’s text does not match the guide’s Segoe UI, 11 pt, Dark Gray formatting.',
      'TOC title':'Add a “Contents” or “Table of Contents” heading with the required title formatting.',
      'Title style':'The cover title is not using the Title style. Apply the modified Title style.',
      'Typed cover title style':'The typed cover title is not using the required Title style.',
      'Student information style':'A student-information line on the cover is not using Normal style.',
    }
    if name=='Overall style-guide adherence':return found+' Review the guide as a complete system: named styles, brand colors, headings, supporting elements and page layout.'
    if name in specific:return specific[name]
    if name.startswith('Section ') and name.endswith('footer fields'):
        return 'A section footer is missing a page number or the linked Document Title. Insert both; a total-page-count field does not replace the title.'
    if name=='Displayed run formatting':
        sizes=sorted(set(re.findall(r'size: (\d+(?:\.\d+)? pt)',found)))
        detail=(' Detected size: '+', '.join(sizes)+'.') if sizes else ''
        if criterion==CRITERIA[9]:return 'Footer text does not match the required Segoe UI, 9 pt, Dark Gray formatting.'+detail+' Update the footer text or its style.'
        return f'Some text overrides the required formatting. It should use {expected}.{detail} Apply the correct style and remove conflicting direct formatting.'
    if ' style ' in name:
        prop=name.rsplit(' style ',1)[1]
        return f'The {name.rsplit(" style ",1)[0]} style has the wrong {prop}. Expected {expected}; found {found}. Modify the style itself.'
    if name=='List indentation':return f'The bullet/list position does not match the accepted indentation. {expected}. Detected {found}.'
    if re.match(r'Section \d+ (top|bottom|left|right) margin',name):return f'The {name.lower()} is {found}; the guide specifies {expected}. Adjust Layout → Margins.'
    if name.startswith('First source '):return f'The first citation’s {name[len("First source "):].lower()} needs correction. Expected “{expected}”; found “{found}”. Edit the source in Word’s Manage Sources.'
    if name=='Automatic Word field':return 'The contents/figures list is missing the required automatically generated Word field. Insert it using Word’s Table of Contents or Table of Figures tool.'
    if name=='Manually created list content':return 'The list heading is present, but recognizable entries are missing. Generate or complete the list.'
    if name=='Font size and style':return 'The footer is missing, so its font and size cannot meet the requirement. Add a footer using Segoe UI, 9 pt, Dark Gray.'
    if name=='Right alignment':return 'The required footer is missing or is not right-aligned.'
    if name=='Page number and Document Title field':return 'The footer is missing the required page number and linked Document Title.'
    if name=='Title page counted as one, number hidden':return 'Configure the title page to count as page 1 while hiding its printed page number.'
    return 'This part of the document does not meet the requirement. Check the cited instruction and apply the required Word feature or formatting.'


def feedback_text(result,criterion=None):
    if 'error' in result:return 'This document could not be graded: '+result['error']
    rows=[r for r in result['criteria'] if r['score']<r['maximum'] and (criterion is None or r['criterion']==criterion)]
    if not rows:return 'No marks lost'+(' for this criterion.' if criterion else ' in this submission.')
    lines=[f'Feedback on lost marks — {result["filename"]}',f'Total: {result["total"]}/60','']
    for row in rows:
        lines += [f'{row["criterion"]}: {row["score"]}/5 ({5-row["score"]} '+('mark lost)' if 5-row['score']==1 else 'marks lost)'),REFERENCES[row['criterion']]]
        if row.get('overall_assessment'):
            lines.append('This is one overall style-guide assessment; the domain evidence is not a list of separate mark deductions.')
        if row.get('scoring_groups'):
            for g in row['scoring_groups']:
                if not g['passed']:lines.append(f'• {g["name"]}: not met. '+('One mark deducted.' if row['criterion']==CRITERIA[9] else f'{g["weight"]} '+('mark allocated to this check.' if g['weight']==1 else 'marks allocated to this check.')))
        seen=set();seen_issues=set()
        for c in row['checks']:
            if not c.get('scored',True) or c['passed'] is not False or c.get('accepted'):continue
            issues=set(c.get('issues',[]))
            if issues and issues<=seen_issues:continue
            seen_issues.update(issues)
            message=explanation(row['criterion'],c)
            if message in seen:continue
            seen.add(message)
            location=c['location']
            # Technical IDs and XML are never included; only meaningful paragraph locations.
            if location.startswith('Paragraph '):message+=' Location: '+location
            lines.append('• '+message)
        if row.get('scoring_groups') and row['criterion']==CRITERIA[9] and row['score']==0:
            lines.append('All four footer checks failed, so this criterion receives 0/5.')
        lines.append('')
    return '\n'.join(lines)
