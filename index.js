console.log("\n\n __________ Word HTML Converter __________\n\n");

let glob = require("glob"),
  mammoth = require("mammoth"),
  fs = require("fs");

let inputDir = "./input/",
  outputDir = "./output/";

glob(inputDir + "*.{doc,docx}", (err, files) => {
  // Convert each file
  files.forEach(filepath => {
    let filename = filepath.substr(inputDir.length);
    // Ignore ~$..docx cache files
    if (filename.substring(0, 2) == "~$") return;
    convertFile({
      filename: filename,
      filepath: filepath,
      outputDir: outputDir
    });
  });
});

const convertFile = options => {
  mammoth
    .convertToHtml({ path: options.filepath })
    .then(function(result) {
      let html = result.value;

      // Replace any weird whitespace characters and trim them to max 1
      html = html.replace(/(\s)+/g, " ");

      // Force adjacent tables to conjoin with an empty row
      // (nb don't use nbsp; instead have td:empty:after { content: '.', transparent, hidden }
      // That way you can still target them with :empty, to hide borders and whatnot
      // html = html.replace(/<\/table><table>/g, '<tr><td></td><td></td></tr>');

      // console.log(html);
      let outputFilepath = options.outputDir + options.filename + ".html";
      console.log(" Writing to file: " + outputFilepath);
      fs.writeFileSync(outputFilepath, html);

      // Escape any quotes
      let json = html.replace(/"/g, '\\"');

      // Remove other gaps while we're at it, actually don't bother because (theoretically)..
      // We could have "<strong>Spaced</strong> <strong>inline</strong> <em>things</em>
      // html = html.replace(/(> <)/g, "><");
      json = json.replace(/(> <\/)/g, "></"); // ..but this handles only closing whitespace

      // JSON wrapper
      json = `{\n  "output": "${json}"\n}`;

      outputFilepath = options.outputDir + options.filename + ".json";
      console.log(" Writing to file: " + outputFilepath);
      fs.writeFileSync(outputFilepath, json);
    })
    .done();
};
