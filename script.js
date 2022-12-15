import * as fs from "fs";
import pkg from "docx";
const { Document, Packer, Paragraph, TextRun } = pkg;
import pkg2 from "prompt-sync";
const prompt = pkg2({ sigint: true });

const fname = prompt("Enter file name: ");
generateDoc(fname);

function generateDoc(fname) {
  // Documents contain sections, you can have multiple sections per document, go here to learn more about sections
  // This simple example will only contain one section
  const doc = new Document({
    sections: [
      {
        properties: {},
        children: [
          new Paragraph({
            children: [
              new TextRun("Hello World"),
              new TextRun({
                text: "Foo Bar",
                bold: true,
              }),
              new TextRun({
                text: "\tGithub is the best",
                bold: true,
              }),
            ],
          }),
        ],
      },
    ],
  });

  // Used to export the file into a .docx file
  Packer.toBuffer(doc).then((buffer) => {
    fs.writeFileSync(fname + ".docx", buffer);
  });

  // Done! A file called 'fname' will be in your file system.
}
