const XlsxTemplate = require("xlsx-template");
const fs = require("fs");
const path = require("path");

fs.readFile(path.join(__dirname, "/templates/read.xlsx"), (err, data) => {
  if (err) {
    console.log(err);
  } else {
    const template = new XlsxTemplate(data);

    // Replacements take place on first sheet
    let sheetNumber = 1;

    // Set up some placeholder values matching the placeholders in the template

    var values = {
      extractDate: new Date(),
      dates: [
        new Date("2013-06-01"),
        new Date("2013-06-02"),
        new Date("2013-06-03"),
      ],
      people: [
        { name: "John Smith", age: 20 },
        { name: "Bob Johnson", age: 22 },
      ],
    };

    // Perform substitution
    template.substitute(sheetNumber, values);

    // Get binary data
    const myData = template.generate({ type: "nodebuffer" });

    fs.writeFile(
      path.join(__dirname, "/templates/write.xlsx"),
      myData,
      function (err) {
        if (err) {
          return console.log(err);
        } else {
          console.log("done");
        }
      }
    );
  }
});
