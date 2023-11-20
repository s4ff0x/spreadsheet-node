const { GoogleSpreadsheet } = require("google-spreadsheet");
const { JWT } = require("google-auth-library");
const fs = require("fs");
const { scheduleJob } = require("node-schedule");

async function loadCredentials() {
  const credentialsJSON = fs.readFileSync("credentials.json");
  return JSON.parse(credentialsJSON);
}

async function initializeSpreadsheet() {
  const { private_key, client_email } = await loadCredentials();
  const authClient = new JWT({
    email: client_email,
    key: private_key,
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });

  const spreadsheetId = "12DWub_GDnD9j7fbTkjQJZQxrkC0H18q-3jhZtgT7tBI";
  return new GoogleSpreadsheet(spreadsheetId, authClient);
}

const MAX_PARTICIPANTS = 30;
const SELECTION_POOL_COL_INDEX = 4;
const ALREADY_PARTICIPATED_COL_INDEX = 5;
const NEXT_DATE_CELL = [1, 2];
const WINNER_CELL = [1, 0];
const DAYS_UNTIL_NEXT_SELECTION = 14;

const main = async () => {
  const doc = await initializeSpreadsheet();
  await doc.loadInfo();
  console.log(doc.title);

  const sheet = doc.sheetsByIndex[0];
  console.log(sheet.title);

  await sheet.loadCells(`A1:F${MAX_PARTICIPANTS}`);

  const participants = getParticipants(sheet);

  if (participants.length === 0) {
    returnAll(sheet);
    return await sheet.saveUpdatedCells();
  }

  if (checkDate(sheet)) {
    updateWinner(participants, sheet);
    updateNextDate(sheet);
    await sheet.saveUpdatedCells();
  }
};

function updateNextDate(sheet) {
  const { formattedValue: nextDateString } = sheet.getCell(...NEXT_DATE_CELL);
  const nextDate = new Date(nextDateString);

  nextDate.setDate(nextDate.getDate() + DAYS_UNTIL_NEXT_SELECTION);
  sheet.getCell(...NEXT_DATE_CELL).value = nextDate.toISOString();
}

function updateWinner(participants, sheet) {
  const winner = participants[Math.floor(Math.random() * participants.length)];
  const insertCell = getInsertCell(sheet);
  sheet.getCell(...WINNER_CELL).value = winner.value; // winner cell
  insertCell.value = winner.value;
  winner.value = null;
}

function getInsertCell(sheet) {
  for (let i = 1; i < MAX_PARTICIPANTS; i++) {
    const cell = sheet.getCell(i, ALREADY_PARTICIPATED_COL_INDEX);
    if (cell.formattedValue === null) return cell;
  }
}

function checkDate(sheet) {
  const { formattedValue: nextDateString } = sheet.getCell(...NEXT_DATE_CELL);
  // TODO: compare 2 dates
  const nextDate = new Date(nextDateString);
  const currentDate = new Date();
  return true;
}

function getParticipants(sheet) {
  const participants = [];
  for (let i = 1; i < MAX_PARTICIPANTS; i++) {
    const cell = sheet.getCell(i, SELECTION_POOL_COL_INDEX);
    if (cell.formattedValue) participants.push(cell);
  }
  return participants;
}

function returnAll(sheet) {
  for (let i = 1; i < MAX_PARTICIPANTS; i++) {
    const cell = sheet.getCell(i, ALREADY_PARTICIPATED_COL_INDEX);
    if (cell.formattedValue) {
      sheet.getCell(i, SELECTION_POOL_COL_INDEX).value = cell.value;
      cell.value = null;
    }
  }
  sheet.getCell(...WINNER_CELL).value = null;
}

// main();
const job = scheduleJob("*/5 * * * * *", main);
