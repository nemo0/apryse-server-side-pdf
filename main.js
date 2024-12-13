const { PDFNet } = require("@pdftron/pdfnet-node");
const fs = require("fs");
const dotenv = require("dotenv");

// Load environment variables
dotenv.config();

// Ensure API key is present
if (!process.env.APRYSE_API_KEY) {
  console.error(
    "Error: APRYSE_API_KEY is not defined in the environment variables."
  );
  process.exit(1);
}

/**
 * Loads JSON data from a file.
 * @param {string} filePath - Path to the JSON file.
 * @returns {object} Parsed JSON object.
 */
function loadJsonData(filePath) {
  try {
    const data = fs.readFileSync(filePath, "utf8");
    return JSON.parse(data);
  } catch (error) {
    console.error(`Error loading JSON data from ${filePath}:`, error.message);
    process.exit(1); // Exit with error
  }
}

/**
 * Generates a PDF from a template and JSON data.
 * @param {string} inputFile - Path to the input Office file.
 * @param {string} outputFile - Path to save the generated PDF.
 * @param {string} jsonFile - Path to the JSON data file.
 */
async function generatePDF(inputFile, outputFile, jsonFile) {
  try {
    console.log("Initializing PDF generation...");

    // Initialize OfficeToPDFOptions
    const options = new PDFNet.Convert.OfficeToPDFOptions();

    // Create a TemplateDocument object from the input Office file
    const templateDoc = await PDFNet.Convert.createOfficeTemplateWithPath(
      inputFile,
      options
    );

    // Load JSON data
    const jsonData = loadJsonData(jsonFile);
    const jsonDataString = JSON.stringify(jsonData);

    // Fill the template with JSON data
    const pdfDoc = await templateDoc.fillTemplateJson(jsonDataString);

    // Save the filled template as a PDF
    await pdfDoc.save(outputFile, PDFNet.SDFDoc.SaveOptions.e_linearized);

    console.log(`PDF successfully saved to ${outputFile}`);
  } catch (error) {
    console.error("An error occurred during PDF generation:", error.message);
  }
}

// Run the script
(async () => {
  const inputFile = process.argv[2] || "./template.docx"; // Default input file
  const outputFile = process.argv[3] || "./output.pdf"; // Default output file
  const jsonFile = process.argv[4] || "data.json"; // Default JSON file

  try {
    await PDFNet.runWithCleanup(
      () => generatePDF(inputFile, outputFile, jsonFile),
      process.env.APRYSE_API_KEY
    );
  } catch (error) {
    console.error("PDFNet initialization error:", error.message);
  } finally {
    await PDFNet.shutdown();
  }
})();
