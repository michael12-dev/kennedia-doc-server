const express = require('express');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  ImageRun, Header, Footer, AlignmentType, BorderStyle, WidthType,
  ShadingType, VerticalAlign, ExternalHyperlink
} = require('docx');

const app = express();
app.use(cors());
app.use(express.json());

const logoImg = fs.readFileSync(path.join(__dirname, 'logo.png'));
const stampImg = fs.readFileSync(path.join(__dirname, 'stamp.jpeg'));

function B(t) { return new TextRun({ text: t || '', bold: true, font: 'Segoe UI', size: 22 }); }
function N(t) { return new TextRun({ text: t || '--', font: 'Segoe UI', size: 22 }); }
function FL(label, val) {
  return new Paragraph({
    spacing: { line: 276, lineRule: 'auto', after: 60 },
    children: [B(label), N('  ' + (val || '--'))]
  });
}
function SP() { return new Paragraph({ spacing: { after: 100 }, children: [new TextRun('')] }); }
function ST(t) {
  return new Paragraph({
    spacing: { after: 80 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: 'C9A84C', space: 2 } },
    children: [new TextRun({ text: t, bold: true, font: 'Segoe UI', size: 22, color: '44546A', allCaps: true })]
  });
}
function TC(children, width, shading) {
  const bdr = { style: BorderStyle.SINGLE, size: 4, color: '44546A' };
  return new TableCell({
    borders: { top: bdr, bottom: bdr, left: bdr, right: bdr },
    width: { size: width, type: WidthType.DXA },
    shading: shading || undefined,
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    verticalAlign: VerticalAlign.CENTER,
    children
  });
}

app.get('/', (req, res) => res.send('Kennedia Doc Generator is running ✅'));

app.post('/generate-doc', async (req, res) => {
  try {
    const s = req.body;
    const dt = new Date(s.submittedAt || Date.now()).toLocaleDateString('en-GB', {
      day: '2-digit', month: 'long', year: 'numeric'
    });

    const photoLines = [];
    [['Street View', s.streetViewURL], ['Gate View', s.gateViewURL], ['Front View', s.frontViewURL]].forEach(([lbl, url]) => {
      if (url) {
        photoLines.push(new Paragraph({
          spacing: { after: 60 },
          children: [
            B(lbl + ': '),
            new ExternalHyperlink({ link: url, children: [new TextRun({ text: url, style: 'Hyperlink', font: 'Segoe UI', size: 20 })] })
          ]
        }));
      } else {
        photoLines.push(FL(lbl + ':', 'Not provided'));
      }
    });

    const doc = new Document({
      sections: [{
        properties: {
          page: {
            size: { width: 12240, height: 15840 },
            margin: { top: 720, right: 1080, bottom: 900, left: 1080 }
          }
        },
        headers: {
          default: new Header({
            children: [
              new Paragraph({
                alignment: AlignmentType.RIGHT, spacing: { after: 0 },
                children: [new TextRun({ text: 'KCL/V/ADD/01', font: 'Segoe UI', size: 18, color: '44546A' })]
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER, spacing: { after: 60 },
                children: [new ImageRun({ data: logoImg, transformation: { width: 180, height: 51 }, type: 'png' })]
              }),
              new Paragraph({
                alignment: AlignmentType.CENTER, spacing: { after: 60 },
                border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: '44546A', space: 1 } },
                children: [
                  new TextRun({ text: 'ADDRESS VERIFICATION FORM', bold: true, font: 'Segoe UI', size: 24, color: '44546A' }),
                  new TextRun({ text: '  (' + dt + ')', font: 'Segoe UI', size: 22, color: '44546A' })
                ]
              })
            ]
          })
        },
        footers: {
          default: new Footer({
            children: [
              new Paragraph({
                spacing: { before: 80, after: 0 },
                border: { top: { style: BorderStyle.SINGLE, size: 4, color: '44546A', space: 1 } },
                children: [new TextRun({ text: 'info@kennediaconsulting.net  OR  helpdesk@kennediaconsulting.net', font: 'Segoe UI', size: 16, color: '44546A' })]
              }),
              new Paragraph({
                spacing: { after: 0 },
                children: [new TextRun({ text: 'Visit Us: 14B Kingsley Emu Street off Chris Efuyemi Onanuga Street, Lekki Phase 1, Lagos, Nigeria', font: 'Segoe UI', size: 16 })]
              })
            ]
          })
        },
        children: [
          SP(), ST('Employee Information'),
          FL('Name of Employee:', s.name),
          FL("Employee's ID:", s.staffId),
          FL('Email Address:', s.email),
          FL('Client Name (Where they work):', s.clientName),
          FL('Address of Employee:', s.address),
          FL('Closest Landmark:', s.landmark),
          FL('Duration at Current Address:', s.durationAddress),
          SP(), ST('Building Details'),
          FL('Description of Building:', s.buildingType),
          FL('Is the Building Painted?', s.isPainted + (s.buildingColour ? '  |  Colour: ' + s.buildingColour : '')),
          FL('Is the Compound Fenced?', s.isFenced),
          FL('Does the Compound Have a Gate?', s.hasGate + (s.gateColour ? '  |  Gate Colour: ' + s.gateColour : '')),
          SP(),
          new Paragraph({
            spacing: { after: 80 },
            children: [new TextRun({ text: 'Any information given will be treated in the strictest of confidence.', italics: true, font: 'Segoe UI', size: 20, color: '44546A' })]
          }),
          new Table({
            width: { size: 10080, type: WidthType.DXA },
            columnWidths: [480, 3600, 6000],
            rows: [
              new TableRow({
                tableHeader: true,
                children: [
                  TC([new Paragraph({ children: [new TextRun({ text: '#', bold: true, font: 'Segoe UI', size: 20, color: 'FFFFFF' })] })], 480, { fill: '44546A', type: ShadingType.CLEAR }),
                  TC([new Paragraph({ children: [new TextRun({ text: 'Question', bold: true, font: 'Segoe UI', size: 20, color: 'FFFFFF' })] })], 3600, { fill: '44546A', type: ShadingType.CLEAR }),
                  TC([new Paragraph({ children: [new TextRun({ text: 'Response', bold: true, font: 'Segoe UI', size: 20, color: 'FFFFFF' })] })], 6000, { fill: '44546A', type: ShadingType.CLEAR })
                ]
              }),
              new TableRow({ children: [TC([new Paragraph({ children: [B('1')] })], 480), TC([new Paragraph({ children: [B('What is your relationship with the employee?')] })], 3600), TC([new Paragraph({ children: [N(s.contactRelationship)] })], 6000)] }),
              new TableRow({ children: [TC([new Paragraph({ children: [B('2')] })], 480), TC([new Paragraph({ children: [B('How long have you known the employee?')] })], 3600), TC([new Paragraph({ children: [N(s.contactDuration)] })], 6000)] }),
              new TableRow({ children: [TC([new Paragraph({ children: [B('3')] })], 480), TC([new Paragraph({ children: [B('How long has the employee lived here?')] })], 3600), TC([new Paragraph({ children: [N(s.durationAddress)] })], 6000)] })
            ]
          }),
          SP(), ST('Contact Person'),
          FL('Name of Contact Person:', s.contactName),
          FL('Phone Number:', s.contactPhone),
          FL('Relationship with Employee:', s.contactRelationship),
          SP(), ST('Uploaded Photos'),
          ...photoLines,
          SP(),
          new Paragraph({
            spacing: { after: 80 },
            border: { top: { style: BorderStyle.SINGLE, size: 4, color: '44546A', space: 1 } },
            children: [B('Is the given address valid?'), N('          ☐  YES          ☐  NO')]
          }),
          new Paragraph({ spacing: { after: 80 }, children: [B('Comments: '), N('_____________________________________________')] }),
          new Paragraph({ spacing: { after: 80 }, children: [B('Date: '), N(dt)] }),
          new Paragraph({
            spacing: { after: 80 },
            children: [
              B('Verification Officer: '), N('Victor Azubuike   '),
              new ImageRun({ data: stampImg, transformation: { width: 90, height: 21 }, type: 'jpeg' })
            ]
          })
        ]
      }]
    });

    const buffer = await Packer.toBuffer(doc);
    const filename = 'AVF_' + (s.name || 'submission').replace(/\s+/g, '_') + '_' + new Date().toISOString().slice(0, 10) + '.docx';

    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.send(buffer);

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
