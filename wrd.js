import * as fs from "fs";
import pkg from "docx";
const {
  AlignmentType,
  Document,
  Packer,
  Paragraph,
  TextRun,
  Header,
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
      characterStyles: [
        {
          id: "code",
          name: "Code",
          basedOn: "Normal",
          quickFormat: true,
          run: {
            // idk why, but border doesnt work, no matter if it is outside or inside run{ }
            border: {
              top: {
                color: "auto",
                space: 1,
                style: "single",
                size: 6,
              },
              bottom: {
                color: "auto",
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
      paragraphStyles: [
        {
          name: "Normal",
          run: {
            size: 20,
            font: "Roboto",
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
