const { google } = require("googleapis");

const SPREADSHEET_ID = "1pZNwOip3teKUuV2bKQGuKnLMikKl5DVieNGVPBtEzqE"; // FONO_Acquirer_Tracker

const TAB_RANGES = ["Acquirer Performance", "Visit Log"];

const auth = new google.auth.GoogleAuth({
  credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT),
  scopes: ["https://www.googleapis.com/auth/spreadsheets.readonly"],
});

module.exports = async (req, res) => {
  try {
    const sheets = google.sheets({ version: "v4", auth });

    const response = await sheets.spreadsheets.values.batchGet({
      spreadsheetId: SPREADSHEET_ID,
      ranges: TAB_RANGES,
      valueRenderOption: "UNFORMATTED_VALUE",
      dateTimeRenderOption: "SERIAL_NUMBER",
    });

    const [acquirerPerformance, visitLog] = response.data.valueRanges.map(
      (r) => r.values || []
    );

    res.status(200).json({ acquirerPerformance, visitLog });
  } catch (err) {
    console.error(err);
    res.status(500).json({ error: err.message || "Failed to fetch FONO_Acquirer_Tracker sheet data" });
  }
};
