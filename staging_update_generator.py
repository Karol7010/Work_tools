import csv
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def add_hyperlink(paragraph, text, url):
    """
    Add a hyperlink to a paragraph.
    Compatible with older python-docx versions.
    """
    # This gets access to the document.xml.rels file and creates a new relationship
    part = paragraph.part
    r_id = part.relate_to(
        url,
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
        is_external=True
    )

    # Create the w:hyperlink tag and add needed values
    hyperlink = OxmlElement('w:hyperlink')
    hyperlink.set(qn('r:id'), r_id)

    # Create a w:r element
    new_run = OxmlElement('w:r')

    # Add formatting
    rPr = OxmlElement('w:rPr')

    color = OxmlElement('w:color')
    color.set(qn('w:val'), '0000FF')
    rPr.append(color)

    u = OxmlElement('w:u')
    u.set(qn('w:val'), 'single')
    rPr.append(u)

    new_run.append(rPr)

    # Create a w:t element and set the text
    text_elem = OxmlElement('w:t')
    text_elem.text = text
    new_run.append(text_elem)

    # Append the hyperlink to the paragraph
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)

    return hyperlink

def generate_docx(date, blocking, major, minor, filename='update.docx'):
    doc = Document()
    doc.add_heading(f'Day {date} PCC Staging Update', level=1)

    # Blocking issues
    doc.add_paragraph('• blocking issue:')
    for element in blocking:
        p_blocking = doc.add_paragraph('    - ')
        for job, link in element["links"].items():
            add_hyperlink(p_blocking, job, link)
            p_blocking.add_run(', ')
        p_blocking.add_run(f' - {element["description"]}')

    # Major issue
    doc.add_paragraph('• major issue:')
    for element in major:
        p_major = doc.add_paragraph('    - ')
        for job, link in element["links"].items():
            add_hyperlink(p_major, job, link)
            p_major.add_run(', ')
        p_major.add_run(f' - {element["description"]}')

    # Minor issue
    doc.add_paragraph('• Minor issue:')
    for element in minor:
        p_minor = doc.add_paragraph('    - ')
        for job, link in element["links"].items():
            add_hyperlink(p_minor, job, link)
            p_minor.add_run(', ')
        p_minor.add_run(f' - {element["description"]}')
    doc.save(filename)
    print(f"Document saved as {filename}")

def csv_to_issues(row_no):
    with open("files/PCC staging TR status summary  (2025_PCC_Staging_TRs) (1).csv") as csvfile:
        Issues = []
        reader = csv.reader(csvfile)
        next(reader)
        # For master
        for row in reader:
            job = row[row_no].upper().strip()
            headline = row[1]
            comment = row[4]
            if job.strip():
                d = ' '.join([headline, comment])
                d = d.replace('\n', '')
                links = {}
                for single_job in job.split(' '):
                    if "OAM" in single_job:
                        job_no = single_job.replace("OAM", "")
                        links[single_job] = f"https://jenkins-blue-grey.karle005.rnd.gic.ericsson.se/job/staging-tests-jcat-fw--PCC-Staging-OAM-SM/{job_no}/"
                    elif "FT" in single_job:
                        job_no = single_job.replace("FT", "")
                        links[single_job] = f"https://jenkins-blue-grey.karle005.rnd.gic.ericsson.se/job/pcc-staging-deployment--Staging-footprint/{job_no}/"
                    elif "UPG" in single_job:
                        job_no = single_job.replace("UPG", "")
                        links[single_job.strip()] = f"https://jenkins-blue-grey.karle005.rnd.gic.ericsson.se/job/staging-tests-jcat-fw--PCC-Staging-Upgrade-SM/{job_no}/"
                    elif "STAB" in single_job:
                        job_no = single_job.replace("STAB", "")
                        links[single_job] = f"https://jenkins-blue-grey.karle005.rnd.gic.ericsson.se/job/staging-tests-jcat-fw--PCC-Staging-Stability-SM/{job_no}/"
                desc_found = False
                for i in range(len(Issues)):
                    if d in Issues[i]:
                        Issues[i]["links"] = Issues[i]["links"] | links
                        desc_found = True
                        break  # if you only want the first match
                if not desc_found:
                    Issues.append({"description": d,"links": links})
        return Issues

with open("files/PCC staging TR status summary  (2025_PCC_Staging_TRs).csv") as csvfile:
    master_issues = csv_to_issues(13)
    release_issues = csv_to_issues(12)
    all_issues = [*master_issues, *release_issues]
    blocking = []
    major = []
    minor = []
    print(all_issues)

    for issue in all_issues:
        if len(issue['links']) < 3:
            minor.append(issue)
        elif len(issue['links']) < 5:
            major.append(issue)
        else:
            major.append(issue)
    generate_docx(
        date='2025-05-30',
        blocking=blocking,
        major=major,
        minor=minor
    )






