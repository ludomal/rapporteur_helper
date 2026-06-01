import docx
from docx.document import Document


def create_hyperlink(document: Document, text: str, url, format=["None", "bold", "italic", "hyperlink", "button"][0]):
    # Create the w:hyperlink tag and add needed values
    hyperlink = docx.oxml.shared.OxmlElement("w:hyperlink")

    # Access to the document settings (DocumentPart) to create a new relation id value
    documentPart = document.part
    r_id = documentPart.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)

    # Attach the relation ID to the hyperlink object
    hyperlink.set(
        docx.oxml.shared.qn("r:id"),
        r_id,
    )

    # Create a w:r element
    new_run = docx.oxml.shared.OxmlElement("w:r")

    # Create a new w:rPr element
    rPr = docx.oxml.shared.OxmlElement("w:rPr")

    if format == "italic":
        rStyle = docx.oxml.shared.OxmlElement("w:i")
        rPr.append(rStyle)
    if format == "bold":
        rStyle = docx.oxml.shared.OxmlElement("w:b")
        rPr.append(rStyle)

    if format == "hyperlink":
        rStyle = docx.oxml.shared.OxmlElement("w:rStyle")
        rStyle.set(docx.oxml.shared.qn("w:val"), "Hyperlink")
        rPr.append(rStyle)

    # Join all the xml elements together and add the required text to the w:r element
    new_run.append(rPr)
    new_run.text = f"{text}"

    hyperlink.append(new_run)

    if format == "button":
        # Create a link button with the Hyperlink style, prettier than coloring the whole text
        new_run = docx.oxml.shared.OxmlElement("w:r")
        new_run.text = "  "
        hyperlink.append(new_run)

        new_run = docx.oxml.shared.OxmlElement("w:r")
        # Unicode character for "download"
        # https://www.fileformat.info/info/unicode/char/2913/fontsupport.htm

        rStyle = docx.oxml.shared.OxmlElement("w:rStyle")
        rStyle.set(docx.oxml.shared.qn("w:val"), "Hyperlink")
        rPr = docx.oxml.shared.OxmlElement("w:rPr")
        rPr.append(rStyle)
        new_run.append(rPr)

        new_run.text = "\u2913"

        hyperlink.append(new_run)

    return hyperlink


def add_hyperlink(paragraph, text, url, format=["None", "bold", "italic", "hyperlink", "button"][0]):
    # :param paragraph: The paragraph we are adding the hyperlink to.
    # :param text: The text displayed for the url
    # :param url: A string containing the required url
    # :param format: Style to apply to the text ['None', 'bold','italic']
    #     :return: The hyperlink object

    # Create hyperlink
    hyperlink = create_hyperlink(paragraph, text, url, format)

    # Append hyperlink to paragraph
    paragraph._p.append(hyperlink)

    return hyperlink
