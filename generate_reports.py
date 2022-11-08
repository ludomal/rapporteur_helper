from docx import Document
import docx

from lxml import html, etree
import requests
from pprint import pprint
import re
import traceback

# questions = range(1, 21)
questions = [7]
meetingDate = "220607"
# meetingDate = "230118"


def add_hyperlink(paragraph, text, url, format = None):
    # :param paragraph: The paragraph we are adding the hyperlink to.
    # :param url: A string containing the required url
    # :param text: The text displayed for the url
    #     :return: The hyperlink object

    # This gets access to the document.xml.rels file and gets a new relation id value
    part = paragraph.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)

    # Create the w:hyperlink tag and add needed values
    hyperlink = docx.oxml.shared.OxmlElement('w:hyperlink')
    hyperlink.set(docx.oxml.shared.qn('r:id'), r_id, )

    # Create a w:r element
    new_run = docx.oxml.shared.OxmlElement('w:r')

    # Create a new w:rPr element
    rPr = docx.oxml.shared.OxmlElement('w:rPr')

    if format == 'italic':
        rStyle = docx.oxml.shared.OxmlElement('w:i')
        rPr.append(rStyle)
    if format == 'bold':
        rStyle = docx.oxml.shared.OxmlElement('w:b')
        rPr.append(rStyle)

    # Join all the xml elements together and add the required text to the w:r element
    new_run.append(rPr)
    new_run.text = text

    hyperlink.append(new_run)

    paragraph._p.append(hyperlink)

    return hyperlink

def insert_paragraph_after(paragraph, text=None, style=None):
    """Insert a new paragraph after the given paragraph."""
    new_p = docx.oxml.shared.OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    new_para = docx.text.paragraph(new_p, paragraph)
    if text:
        new_para.add_run(text)
    if style is not None:
        new_para.style = style
    return new_para

def get_documents(docSection, endpoint):
    print(f"Retrieving documents from: {endpoint['url']}")
    x = requests.get(endpoint['url'])
    tree = html.fromstring(x.content)

    # Find and parse all rows (<tr>) in the document
    rows = tree.xpath('//tr')

    docSection.text = ""

    # Parse  in descending order
    for i in range(len(rows) -1, 0, -1):
        row = rows[i]
        columns = row.xpath('.//td')

        try:
            # A document row should have the attributes below
            # If not, then the row is ignored

            # Link and document number should be in the second column
            link = hostname + '/' + columns[1].xpath('.//a')[0].attrib['href'].strip()
            number = columns[1].xpath('.//a/strong/text()')[0].strip().replace('[ ', endpoint['prefix']).replace(' ]', '')
            try:
                revision = columns[1].xpath('.//font/text()')[0]
                x = re.search(r"([\d]+)\)", revision)
                revision = x.group(1)
                number = f"{number}r{revision}"
            except Exception as e:
                # print(e)
                pass

            # Title should be in third row
            title = columns[2].xpath('.//text()')[0].strip()
            sources = columns[3].xpath('.//a')
            src = []
            for source in sources:
                src.append(dict(link = f'{hostname}/{source.attrib["href"]}', text = source.text.strip()))

            # Relevant questions should be in fourth column
            questions = columns[4].xpath('.//a')
            q = []
            for quest in questions:
                q.append(dict(link = f'{hostname}/{quest.attrib["href"]}', text = quest.text.strip().replace('/12', '')))

            # Generate word document block for this document
            # p = document.add_paragraph()
            p = docSection.insert_paragraph_before()

            add_hyperlink(p, f"{number} - {title}", link, 'bold')

            p.add_run('\nSources: ')
            for item in src:
                add_hyperlink(p, item['text'], item['link'])
                if src[-1] != item:
                    p.add_run(' | ')

            p.add_run('\nQuestions: ')
            for item in q:
                add_hyperlink(p, item['text'], item['link'])
                if q[-1] != item:
                    p.add_run(', ')

            p.add_run('\nSummary:\n')


        except Exception as e:
            # print(e)
            # traceback.print_exc()
            pass

def get_questions_details():
    info = {}

    url="https://www.itu.int/net4/ITU-T/lists/loqr.aspx?Group=12&Period=17"
    try:
        x = requests.get(url)
        tree = html.fromstring(x.content)
    except Exception as e:
        print(url)
        raise(e)

    # Find and parse all rows (<tr>) in the document
    rows = tree.xpath('//tr')


    for row in rows:
        try:
            columns = row.xpath('.//td')
            # This should be three elements
            tmp = columns[0].xpath('.//span/text()')
            res = re.search(r'Q(\d+)/12.*WP(\d+)/12', tmp[0])
            qNum = res.group(1)
            wpNum = res.group(2)
            info[int(qNum)] = dict(wp = wpNum, title = tmp[2])
            # print(info[qNum])
        except:
            pass

    return info

def find_element(document, text):
    for paragraph in document.paragraphs:
        if text in paragraph.text:
            return paragraph

def replace(find, replace):
    for paragraph in document.paragraphs:
        foundInRun = False
        for run in paragraph.runs:
            if find in run.text:
                run.text = run.text.replace(find, replace)
                # print(f'run: {find}')
                foundInRun = True
        if foundInRun == False:
            if find in paragraph.text:
                paragraph.text = paragraph.text.replace(find, replace)
                # print(f'paragraph: {find}')

    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    foundInRun = False
                    for run in paragraph.runs:
                        if find in run.text:
                            run.text = paragraph.text.replace(find, replace)
                            # print(f'run: {find}')
                            foundInRun = True
                    if foundInRun == False:
                        if find in paragraph.text:
                            paragraph.text = paragraph.text.replace(find, replace)
                            # print(f'paragraph: {find}')



if __name__ == '__main__':
    try:
        questionInfo = get_questions_details()
    except Exception as e:
        print("Error - Cannot fetch question details from ITU-T website")
        raise(e)

    for question in questions:
        print(f"Generating report for Q{question}")
        try:
            hostname = 'https://www.itu.int'
            endpoints = [
                dict(url = f'{hostname}/md/meetingdoc.asp?lang=en&parent=T22-SG12-{meetingDate}-C&question=Q{question}/12', prefix='C-', title='Contributions'),
                dict(url = f'{hostname}/md/meetingdoc.asp?lang=en&parent=T22-SG12-{meetingDate}-TD&question=Q{question}/12', prefix='TD-', title='Temporary Documents'),
            ]

            with open('template.docx', 'rb') as f:
                document = Document(f)

            # Insert contributions
            endpoint = endpoints[0]
            docSection = find_element(document, 'Copy table of contributions')
            get_documents(docSection, endpoint)

            # Insert temporary documents
            endpoint = endpoints[1]
            docSection = find_element(document, 'Copy the TD table')
            get_documents(docSection, endpoint)

            # Replace question number
            replace('X/12', f'{question}/12')
            replace('x/12', f'{question}/12')
            replace('t22sg12qX@lists.itu.int', f't22sg12q{question}@lists.itu.int')

            # Replace working party number
            replace('Working Party y/12', f"Working Party {questionInfo[question]['wp']}/12")

            # Replace question title
            replace('[title of question]', questionInfo[question]['title'])
            replace('Title of question', questionInfo[question]['title'])


            document.save(f'Q{question}_status_report.docx')
        except:
            traceback.print_stack()
            pprint(questionInfo)
            raise()
