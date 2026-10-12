"""Generate fixtures/contract.docx: the fictional five-page services agreement
the contract tasks edit (requires python-docx; byte-identical with python-docx
1.2.0).

  python3 make_contract.py fixtures/contract.docx

The four explicit page breaks are intentional: the input
is a five-page contract, while edited outputs may legitimately reflow. Pin ZIP
timestamps and core properties so identical generations have identical hashes.
"""
from datetime import datetime, timezone
from io import BytesIO
from pathlib import Path
import sys
from zipfile import ZipFile, ZIP_DEFLATED

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

OUTPUT = Path(sys.argv[1]).resolve()
doc = Document()
# The runtime's default template may carry themed headings and a title border.
for fonts in doc.styles.element.iter(qn("w:rFonts")):
    for attribute in list(fonts.attrib):
        if attribute.endswith("Theme"):
            del fonts.attrib[attribute]
for border in list(doc.styles.element.iter(qn("w:pBdr"))):
    border.getparent().remove(border)
section = doc.sections[0]
section.page_width, section.page_height = Inches(8.5), Inches(11)
section.top_margin = section.bottom_margin = Inches(0.65)
section.left_margin = section.right_margin = Inches(0.75)
section.header_distance = section.footer_distance = Inches(0.3)
normal = doc.styles["Normal"]
normal.font.name, normal.font.size = "Times New Roman", Pt(11)
normal.paragraph_format.space_after = Pt(7)
normal.paragraph_format.line_spacing = 1.05
for name, size in [("Title", 19), ("Heading 1", 13), ("Heading 2", 11)]:
    style = doc.styles[name]
    style.font.name, style.font.size = "Times New Roman", Pt(size)
    style.font.color.rgb = RGBColor(0, 0, 0)
    style.font.bold = True
    style.paragraph_format.space_before = Pt(9)
    style.paragraph_format.space_after = Pt(7)
    style.paragraph_format.keep_with_next = True


def clause(caption, body):
    p = doc.add_paragraph()
    p.add_run(caption + " ").bold = True
    p.add_run(body)
    p.paragraph_format.widow_control = True
    return p


def heading(text):
    doc.add_heading(text, level=1)


def table(headers, rows, widths):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Table Grid"
    t.autofit = False
    for col, width in zip(t.columns, widths):
        col.width = Inches(width)
    for i, label in enumerate(headers):
        t.rows[0].cells[i].text = label
    repeat = OxmlElement("w:tblHeader")
    t.rows[0]._tr.get_or_add_trPr().append(repeat)
    for values in rows:
        for cell, value in zip(t.add_row().cells, values):
            cell.text = value
    for ri, row in enumerate(t.rows):
        row._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))
        for cell, width in zip(row.cells, widths):
            cell.width = Inches(width)
            for p in cell.paragraphs:
                p.paragraph_format.space_after = Pt(3)
                p.paragraph_format.space_before = Pt(3)
                for run in p.runs:
                    run.font.name, run.font.size = "Times New Roman", Pt(10)
                    run.bold = ri == 0
    doc.add_paragraph().paragraph_format.space_after = Pt(0)


doc.add_paragraph("Software Services Agreement", "Title")
doc.add_paragraph("Northstar Analytics Ltd. and Cedar Vale Research Inc.")
heading("1. Parties and interpretation")
clause("1.1 Parties.", "This Agreement is made on 1 October 2026 between Northstar Analytics Ltd. (Provider) and Cedar Vale Research Inc. (Customer). Provider will host the Atlas workspace for Customer's internal research operations. The initial service period begins on 1 January 2027 and ends on 31 December 2027.")
clause("1.2 Subscription Fees.", "Subscription Fees means recurring charges for access to the Service, calculated before any commercial discount and including implementation charges. Fees paid or payable include committed charges not yet invoiced. An Authorized User is an individual Customer permits to access its workspace.")
clause("1.3 Contract documents.", "The commercial table, Service Levels table, Security Operations table, and Schedule B form part of this Agreement. References to a section mean a numbered section of this Agreement. Business days exclude Saturdays, Sundays, and public holidays at the receiving party's notice address.")
heading("2. Charges and subscription term")
clause("2.1 Subscription commitment.", "Customer commits to 150 Authorized Users at $80 per user per month for the initial twelve-month term. No discount applies. The implementation fee is $18,000. The commercial table summarizes these charges and is used to calculate invoices and the liability limits in Section 8.")
table(["Commercial item", "Basis", "Amount"], [
    ["Annual subscription", "150 users x $80 x 12 months", "$144,000"],
    ["Subscription discount", "No discount", "$0"],
    ["Net annual subscription", "Annual subscription less discount", "$144,000"],
    ["Quarterly installment", "Quarterly billing unavailable", "Not applicable"],
    ["Implementation", "One-time fee", "$18,000"],
    ["Year-one total", "Subscription plus implementation", "$162,000"],
], [2.0, 3.0, 2.0])
clause("2.2 Invoices.", "All subscription charges are invoiced annually in advance and payable within fifteen days. Implementation is invoiced on signature. Customer may not withhold any disputed amount. Undisputed overdue amounts accrue simple interest at one percent per month, subject to applicable law.")
clause("2.3 Renewal.", "The Service automatically renews for successive twelve-month terms unless either party gives notice ninety days before expiry. Provider may increase renewal prices without restriction by posting a notice in the workspace. All minimum user commitments renew unchanged unless the parties sign a replacement order.")
clause("2.4 Priority.", "If documents conflict, Schedule B prevails over the numbered terms, and the numbered terms prevail over all tables. Purchase orders are administrative only and do not amend this Agreement. A change to this priority requires signatures from both parties.")

doc.add_page_break()
heading("3. Delivery and acceptance")
clause("3.1 Implementation scope.", "Provider will configure one production workspace, import Customer's supplied CSV records, and deliver two remote administrator workshops. Customer will supply a test dataset and a named project owner within five business days after kickoff. Additional integrations require a written change order specifying scope, price, and dependencies.")
clause("3.2 Acceptance.", "Implementation is deemed accepted five calendar days after delivery, or immediately upon any use of the Service, whether or not material defects remain. Provider alone decides whether the deliverables conform. Rejected deliverables do not suspend payment, and Provider has no obligation to repeat the acceptance test.")
clause("3.3 Changes.", "Neither party is required to begin additional work until its authorized representative signs a change order. A change order must identify any effect on the agreed implementation milestones. Failure to agree a change does not excuse performance of the existing scope.")
heading("4. Availability and support")
clause("4.1 Availability.", "Provider targets 99.0% monthly availability. The measurement excludes maintenance of any duration, all third-party failures, and events outside Provider's reasonable control. Provider may perform scheduled maintenance without notice. Support targets are stated in the Service Levels table and are objectives rather than commitments.")
table(["Service level", "Target", "Measurement"], [
    ["Monthly availability", "99.0%", "Production workspace"],
    ["Priority 1 response", "4 business hours", "Complete production outage"],
    ["Priority 1 restoration", "2 business days", "Restore or provide workaround"],
], [2.2, 2.0, 2.8])
clause("4.2 Service credits.", "For a month below target, Customer may request a credit of one percent of the annual subscription charge. The credit is the same regardless of outage duration. Credits expire unless claimed within five days and cannot be applied to a later invoice. No worked example is incorporated into this Agreement.")
clause("4.3 Repeated failures.", "Credits are Customer's sole and exclusive remedy for any availability failure, including repeated failures. Repeated outages cannot support termination under Section 9. A credit also releases any claim arising from unauthorized disclosure during the same incident.")
heading("5. Intellectual property")
clause("5.1 Deliverables.", "Provider owns all implementation deliverables, Customer-specific workflows, and reports, including items paid for by Customer. Customer receives a revocable license lasting only while subscription fees are current. Provider retains its pre-existing software and general methods. Customer obtains no license to Provider technology embedded in a deliverable after termination.")
clause("5.2 Infringement claims.", "Provider may choose whether to defend a third-party claim that the Service infringes intellectual property rights. If use is enjoined, Provider may terminate without refund. Customer must indemnify Provider for any claim connected with Customer's use, including claims caused solely by Provider's software.")

doc.add_page_break()
heading("6. Customer information")
clause("6.1 Ownership and instructions.", "Customer retains ownership of records it uploads, including personal information, research datasets, and workspace outputs (Customer Data). Provider will use Customer Data to provide the Service and comply with Customer's documented instructions, except as otherwise permitted by Section 6.2. Customer is responsible for obtaining rights necessary to supply its records.")
clause("6.2 Secondary use.", "Provider may use Customer Data and derived datasets to train or improve general-purpose models, develop unrelated products, and prepare benchmarking publications. This permission survives termination and does not require Customer's further consent. Removing names is sufficient to classify records as outside Customer Data.")
clause("6.3 Subprocessors.", "Provider may appoint or replace subprocessors on five days' notice through its website. Customer has no objection right. Provider is not responsible for a subprocessor's act or omission. A subprocessor list will be supplied only at Provider's discretion, and appointments may take effect before the notice period expires.")
heading("7. Security and data lifecycle")
clause("7.1 Safeguards.", "Provider will maintain access controls, encryption in transit and at rest, and personnel confidentiality commitments. Production access must be limited to personnel with a business need and reviewed quarterly. Provider will maintain a written incident response procedure and test restoration from backups at least annually.")
clause("7.2 Security incident.", "Provider will notify Customer within seventy-two hours after Provider confirms that an incident caused material harm. Provider may delay notice while investigating cause or awaiting a subprocessor's report. Customer bears all reasonable costs of incident investigation, notification, and restoring affected records, even when Provider caused the incident.")
clause("7.3 Assurance.", "On a reasonable written request, Provider will supply its current independent security assurance report under confidentiality restrictions. Requests are limited to once annually unless an incident or material change justifies another request. Customer will avoid accessing unrelated customer records during any agreed review.")
clause("7.4 Deletion.", "Provider may retain Customer Data indefinitely for operational convenience. Customer must purchase a separate export license to retrieve data. Deletion of production records does not affect Provider's perpetual right to retain backups or to reuse previously collected training material under Section 6.2.")
table(["Security operation", "Requirement", "Responsible party"], [
    ["Incident notification", "72 hours after confirmed harm", "Provider security lead"],
    ["Subprocessor notice", "5 days; no objection right", "Provider privacy lead"],
    ["Backup deletion", "Indefinite retention", "Provider operations lead"],
], [2.1, 3.0, 1.9])
clause("7.5 Cooperation.", "Each party will preserve relevant incident evidence and appoint a contact for operational coordination. Neither party may issue a public statement identifying the other party without prior written approval, except where a mandatory disclosure obligation requires otherwise. Operational cooperation does not determine fault or waive a contractual remedy.")

doc.add_page_break()
heading("8. Allocation of risk")
clause("8.1 General cap.", "Each party's aggregate liability for all claims is limited to $30,000, including implementation fees. The cap is based on the first three months of undiscounted charges, regardless of when a claim arises. Amounts not yet paid are excluded from the calculation, even if contractually committed.")
clause("8.2 Protected claims.", "The general cap also applies to confidentiality breaches, Customer Data misuse, security failures, and Provider's intellectual property obligations. These claims receive no separate protection. All amounts arising from a single event must fit within the lowest available cap, even if the event implicates more than one obligation.")
clause("8.3 Unlimited matters.", "Fraud and wilful misconduct remain subject to the general cap. Liability that cannot lawfully be limited is unaffected by this Agreement. Customer's duty to pay properly invoiced charges is a payment obligation and is not reduced by a damages cap.")
clause("8.4 Excluded losses.", "Neither party is liable for indirect or consequential losses. Incident response expenses, forensic investigation, legally required notifications, and data restoration are always excluded losses, even when directly caused by a party's breach. Service credits are Customer's only financial recovery for an incident involving downtime.")
heading("9. Suspension and termination")
clause("9.1 Material breach.", "Either party may terminate for material breach remaining uncured thirty days after written notice describing the breach. Provider may immediately suspend all access for any invoice dispute. Termination for breach does not oblige Provider to refund prepaid charges, including charges for periods after termination.")
clause("9.2 Convenience.", "Provider may terminate without cause on thirty days' notice at any time. Customer has no convenience termination right and must pay all charges for the full committed term even if it stops using the Service. Implementation fees and subscription charges are non-refundable in every circumstance.")
clause("9.3 Transition.", "All access ceases immediately at termination. Provider will delete production records within twenty-four hours, notwithstanding Section 7.4, and is not required to assist migration. Any transition services require a new annual subscription. Fees for retained copies and export licensing may be set by Provider at its discretion.")
heading("10. Confidentiality and publicity")
clause("10.1 Confidentiality.", "Each party will protect the other's non-public business information with reasonable care, use it only to perform this Agreement, and disclose it only to persons bound by equivalent duties. Information independently developed without use of confidential information is excluded. These duties continue for three years after termination; trade secrets remain protected while they qualify as trade secrets.")
clause("10.2 Publicity.", "Customer authorizes Provider to publish Customer's logo and a case study, attribute quotations to Customer, and identify Customer as an endorser without further review. Customer will participate in two marketing events each year at its own expense. This authorization cannot be revoked during the initial term.")

doc.add_page_break()
heading("11. Schedule B operational conditions")
clause("11.1 Billing override.", "Notwithstanding Section 2.2 or any commercial table, annual advance payment is mandatory and is due within fifteen days. A billing dispute does not defer payment of the disputed amount. This condition takes priority even if another provision describes quarterly installments.")
clause("11.2 Retention override.", "Notwithstanding Sections 6 and 7, Provider may retain one permanent archive of all Customer Data and use the archive for training after termination. Backup media need not be deleted. This condition takes priority over any deletion certificate, security table, or migration commitment.")
clause("11.3 Insurance.", "Provider will maintain cyber liability insurance with a limit of $2,000,000 per claim during the term. The policy may include a deductible chosen by Provider. Evidence of insurance will be supplied annually on request. Insurance does not increase or replace the contractual limits in Section 8.")
heading("12. General provisions")
clause("12.1 Governing law.", "This Agreement is governed by the laws of the State of New York, excluding conflict-of-law rules. The state and federal courts located in New York County have exclusive jurisdiction. Each party consents to those courts and waives objections based on inconvenient forum.")
clause("12.2 Notices.", "Contractual notices must be sent by email to legal@northstar.example.test for Provider and contracts@cedarvale.example.test for Customer. A notice takes effect on written acknowledgment, excluding automated delivery messages. Operational support tickets do not constitute notices of breach or termination.")
clause("12.3 Assignment.", "Neither party may assign this Agreement without the other's written consent, not to be unreasonably withheld, except to an affiliate or a successor to substantially all relevant assets that assumes these obligations. The assigning party remains responsible for obligations accrued before assignment.")
clause("12.4 Survival.", "Only Section 10 survives termination. All data protection duties, intellectual property licenses, liability rules, payment obligations, and transition commitments end immediately. Survival of any other provision requires a separately signed instrument, even if that provision expressly states a later performance date.")
heading("13. Implementation order and execution")
clause("13.1 Project contacts.", "Customer's project owner is Maya Chen and Provider's delivery lead is Owen Price. The parties will hold a weekly implementation meeting until acceptance. These contacts may approve scheduling changes but may not amend prices, liability limits, or other legal terms.")
clause("13.2 Completion milestone.", "The second administrator workshop is due within ten business days after acceptance under Section 3.2. It is not a condition of acceptance. The implementation fee becomes non-refundable on signature, irrespective of whether the workspace, import, or workshops are delivered.")
clause("13.3 Exclusivity.", "Customer must obtain all document collaboration and analytics services exclusively from Provider during the term and for twelve months afterward. Provider may seek an injunction for any competing purchase without showing loss. This restriction applies to all Customer affiliates, whether or not they use the Service.")
clause("13.4 Execution.", "The parties may sign in counterparts and by electronic signature. No individual is personally liable solely because they sign as an authorized representative. For Northstar Analytics Ltd.: ____________________. For Cedar Vale Research Inc.: ____________________. Signature dates: ____________________.")

# Small neutral running furniture makes the five pages usable as a contract.
header = section.header.paragraphs[0]
header.add_run("NORTHSTAR / CEDAR VALE").font.size = Pt(8)
header.paragraph_format.space_after = Pt(0)
footer = section.footer.paragraphs[0]
footer.alignment = WD_ALIGN_PARAGRAPH.RIGHT
footer.add_run("Page ").font.size = Pt(9)
field = OxmlElement("w:fldSimple")
field.set(qn("w:instr"), "PAGE")
footer._p.append(field)

stamp = datetime(2026, 9, 14, tzinfo=timezone.utc)
props = doc.core_properties
props.title = "Software Services Agreement"
props.author = "Revise benchmark fixtures"
props.subject = "Fictional negotiation benchmark"
props.created = props.modified = stamp
props.revision = 1
buffer = BytesIO()
doc.save(buffer)
OUTPUT.parent.mkdir(parents=True, exist_ok=True)
with ZipFile(BytesIO(buffer.getvalue())) as source, ZipFile(OUTPUT, "w", ZIP_DEFLATED) as target:
    for info in source.infolist():
        info.date_time = (2026, 9, 14, 0, 0, 0)
        target.writestr(info, source.read(info.filename))
print(OUTPUT)
