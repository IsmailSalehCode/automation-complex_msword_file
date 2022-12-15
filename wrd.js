import * as fs from "fs";
import pkg from "docx";
const {
  AlignmentType,
  Document,
  Packer,
  Paragraph,
  TextRun,
  Header,
  Footer,
  PageNumber,
} = pkg;
import pkg2 from "prompt-sync";
const prompt = pkg2({ sigint: true });

const fname = prompt("Enter file name: ");
generateDoc(fname);

function generateDoc(fname) {
  const marginNum = 550;
  const doc = new Document({
    sections: [
      {
        // https://docx.js.org/#/usage/page-numbers
        properties: {
          page: {
            margin: {
              top: marginNum,
              right: marginNum,
              bottom: marginNum,
              left: marginNum,
            },
          },
        },
        headers: {
          default: new Header({
            children: [new Paragraph(fname)],
          }),
        },
        footers: {
          default: new Footer({
            children: [
              new Paragraph({
                alignment: AlignmentType.END,
                children: [
                  new TextRun(fname + " "),
                  new TextRun({
                    children: [
                      "Page ",
                      PageNumber.CURRENT,
                      " of ",
                      PageNumber.TOTAL_PAGES,
                    ],
                  }),
                ],
              }),
            ],
          }),
        },
        children: [],
      },
    ],
  });

  // Used to export the file into a .docx file
  Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync(fname + ".docx", buffer);
  });

  // Done! A file called 'fname' will be in your file system.
}
