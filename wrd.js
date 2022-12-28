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
const prompt = pkg2();

const fname = prompt("Enter file name: ");
generateDoc(fname);

function generateDoc(fname) {
  const marginNum = 550;
  const doc = new Document({
    creator: "Ismail Saleh",
    styles: {
      paragraphStyles: [
        {
          name: "Normal",
          run: {
            size: 20,
            font: "Roboto",
          },
        },
        {
          id: "code",
          name: "Code",
          quickFormat: true,
          run: {
            border: {
              top: {
                color: "black",
                space: 1,
                style: "single",
                size: 6,
              },
              bottom: {
                color: "black",
                space: 1,
                style: "single",
                size: 6,
              },
            },
            size: 20,
            font: "Consolas",
          },
        },
      ],
    },
    sections: [
      {
        // https://docx.js.org/#/usage/page-numbers
        properties: {
          page: {
            margin: {
              header: marginNum,
              footer: marginNum,
              top: marginNum,
              right: marginNum,
              bottom: marginNum,
              left: marginNum,
            },
          },
        },
        headers: {
          default: new Header({
            children: [
              new Paragraph({
                alignment: AlignmentType.END,
                children: [
                  new TextRun({
                    bold: true,
                    text: fname + " ",
                  }),
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

  try {
    Packer.toBuffer(doc)
      .then((buffer) => {
        fs.writeFileSync(fname + ".docx", buffer);
      })
      .then(console.log("Done"));
  } catch (err) {
    console.log("Error occured\n");
    console.log(err);
  }
}
