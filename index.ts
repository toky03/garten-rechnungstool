import { PDF } from "swissqrbill";
import { mm2pt } from "swissqrbill/utils";

import * as swissqr from "swissqrbill";

import Excel from "exceljs";
import { Languages } from "swissqrbill/types";

interface InvoiceDetails {
  creditor: swissqr.types.Creditor;
  ueberschriftDe: string;
  ueberschriftFr: string;
}

interface VariableData {
  pachtzins: number;
  wasserbezug: number;
  gfAbonement: number;
  strom: number;
  versicherung: number;
  mitgliederbeitrag: number;
  reparaturfonds: number;
  verwaltungskosten: number;
}

interface DebtorData {
  parzelle: string;
  are: number;
  isVorstand: boolean;
  language: swissqr.types.Languages;
  lastName: string;
  debtor: swissqr.types.Debtor;
}

interface CalculatedData extends VariableData {
  total: number;
  are: number;
}

createBills();

async function createBills(): Promise<void> {
  const workbook = new Excel.Workbook();
  await workbook.xlsx.readFile("data/mitgliederliste.xlsx");
  const debtorData = await readDebitorData(workbook);
  const variableData = await readVariableData(workbook);
  const rechnungsdetails = await readRechnungsdetails(workbook);

  debtorData.forEach((debtor: DebtorData) => {
    const calculatedTableData = calculateTableData(variableData, debtor);
    const data: swissqr.types.Data = {
      currency: "CHF",
      amount: calculatedTableData.total,
      creditor: rechnungsdetails.creditor,
      debtor: debtor.debtor,
      message: "Parzelle: " + debtor.parzelle,
    };

    const fileName = `data/bills/${debtor.parzelle}_${debtor.lastName}.pdf`;

    const pdf = new PDF(data, fileName, {
      language: debtor.language,
      autoGenerate: false,
      size: "A4",
    });

    pdf.image("./data/logo.png", mm2pt(20), mm2pt(14), { width: mm2pt(50) });

    addAddressData(pdf, debtor.parzelle as string, debtor.debtor);
    addTitleAndDate(pdf, debtor.language, rechnungsdetails);
    const tableData = calculateTableData(variableData, debtor);
    const table = createTable(tableData, variableData);
    pdf.addTable(table);
    pdf.addQRBill();
    pdf.end();
  });
}

async function readDebitorData(
  workbook: Excel.Workbook
): Promise<DebtorData[]> {
  const mitglieder = await workbook.getWorksheet("Mitgliederliste");

  const debtorDatas = [];

  mitglieder.eachRow((row: Excel.Row, rowNumber: number) => {
    const zip = +row.getCell(5).value;
    const rowLang = row.getCell(11).text;

    if (!zip || !rowLang) {
      return;
    }
    const lastName = row.getCell(2).text;
    const debtor: swissqr.types.Debtor = {
      name: `${row.getCell(3).value} ${lastName}`,
      address: row.getCell(4).text,
      zip,
      city: row.getCell(6).text,
      country: "CH",
    };

    const parzelle = row.getCell(1).value;
    const are = +row.getCell(8).value;
    const isVorstand: boolean = row.getCell(12).value === "J";
    const language: swissqr.types.Languages = rowLang === "FR" ? "FR" : "DE";

    debtorDatas.push({
      parzelle,
      are,
      isVorstand,
      language,
      debtor,
      lastName: lastName.replace(" ", "_").replace("/", "-"),
    });
  });
  return debtorDatas;
}

async function readVariableData(
  workbook: Excel.Workbook
): Promise<VariableData> {
  const worksheet = await workbook.getWorksheet("Betraege");
  const dataRow = worksheet.getRow(2);
  return {
    pachtzins: +dataRow.getCell(1).value,
    wasserbezug: +dataRow.getCell(2).value,
    gfAbonement: +dataRow.getCell(3).value,
    strom: +dataRow.getCell(4).value,
    versicherung: +dataRow.getCell(5).value,
    mitgliederbeitrag: +dataRow.getCell(6).value,
    reparaturfonds: +dataRow.getCell(7).value,
    verwaltungskosten: +dataRow.getCell(8).value,
  };
}

async function readRechnungsdetails(
  workbook: Excel.Workbook
): Promise<InvoiceDetails> {
  const worksheet = await workbook.getWorksheet("Rechnungsdetails");
  return {
    creditor: {
      account: worksheet.getCell(6, 2).text,
      name: worksheet.getCell(1, 2).text,
      address: worksheet.getCell(2, 2).text,
      zip: worksheet.getCell(4, 2).text,
      city: worksheet.getCell(5, 2).text,
      buildingNumber: worksheet.getCell(3, 2).text,
      country: "CH",
    },
    ueberschriftDe: worksheet.getCell(7, 2).text,
    ueberschriftFr: worksheet.getCell(8, 2).text,
  };
}

function roundHalf(num: number): number {
  return Math.round(num * 20) / 20;
}

function calculateTableData(
  variableData: VariableData,
  debtorData: DebtorData
): CalculatedData {
  const pachtzins: number = +variableData.pachtzins * +debtorData.are;
  const wasserbezug: number = +variableData.wasserbezug * +debtorData.are;
  const gfAbonement: number = +variableData.gfAbonement;
  const strom: number = +variableData.strom;
  const versicherung: number = +variableData.versicherung;
  const mitgliederbeitrag: number = +debtorData.isVorstand
    ? 0
    : +variableData.mitgliederbeitrag;
  const reparaturfonds: number = +variableData.reparaturfonds;
  const verwaltungskosten: number = +variableData.verwaltungskosten;

  const total =
    pachtzins +
    wasserbezug +
    gfAbonement +
    strom +
    versicherung +
    mitgliederbeitrag +
    reparaturfonds +
    verwaltungskosten;
  return {
    pachtzins: roundHalf(pachtzins),
    wasserbezug: roundHalf(wasserbezug),
    gfAbonement,
    strom,
    versicherung,
    mitgliederbeitrag,
    reparaturfonds,
    verwaltungskosten,
    are: debtorData.are,
    total: roundHalf(total),
  };
}
function addAddressData(
  pdf: PDF,
  parzelleNumber: string,
  debtor: swissqr.types.Debtor
): void {
  pdf.fillColor("black");
  pdf.fontSize(12);
  pdf.font("Helvetica");
  pdf.text(
    parzelleNumber +
      "\n" +
      debtor.name +
      "\n" +
      debtor.address +
      "\n" +
      debtor.zip +
      " " +
      debtor.city,
    mm2pt(130),
    mm2pt(60),
    {
      width: mm2pt(70),
      height: mm2pt(50),
      align: "left",
    }
  );
}

function addTitleAndDate(
  pdf: PDF,
  language: Languages,
  rechnungsdetails: InvoiceDetails
): void {
  const date = new Date();

  pdf.fontSize(12);
  pdf.font("Helvetica");
  pdf.text(
    "Brügg, " +
      date.getDate() +
      "." +
      (date.getMonth() + 1) +
      "." +
      date.getFullYear(),
    mm2pt(130),
    mm2pt(90),
    {
      width: mm2pt(170),
      align: "left",
    }
  );

  pdf.fontSize(14);
  pdf.font("Helvetica-Bold");
  pdf.text(
    language === "FR"
      ? rechnungsdetails.ueberschriftFr
      : rechnungsdetails.ueberschriftDe,
    mm2pt(20),
    mm2pt(100),
    {
      width: mm2pt(170),
      align: "left",
    }
  );
}

function createTable(
  tableData: CalculatedData,
  variableData: VariableData
): swissqr.types.PDFTable {
  return {
    width: mm2pt(170),
    fontSize: 9,
    rows: [
      {
        height: 30,
        fillColor: "#ECF0F1",
        columns: [
          {
            text: "Anzahl/Nombre",
            width: mm2pt(20),
          },
          {
            text: "Einheit/Unité",
            width: mm2pt(20),
          },
          {
            text: "Bezeichnung/Dénomination",
            width: mm2pt(60),
          },
          {
            text: "Preis/Prix",
            width: mm2pt(30),
          },
          {
            text: "Betrag/Montant",
            width: mm2pt(32),
          },
        ],
      },
      {
        columns: [
          {
            text: tableData.are,
            width: mm2pt(20),
          },
          {
            text: "Aren / Are",
            width: mm2pt(20),
          },
          {
            text: "Pachtzins / Loyer de la parcelle",
            width: mm2pt(60),
          },
          {
            text: "CHF " + variableData.pachtzins.toFixed(2),
            width: mm2pt(30),
            textOptions: { align: "right" },
          },
          {
            text: "CHF " + tableData.pachtzins.toFixed(2),
            width: mm2pt(30),
            textOptions: { align: "right" },
          },
        ],
      },
      {
        columns: [
          {
            text: tableData.are,
            width: mm2pt(20),
          },
          {
            text: "Aren / Are",
            width: mm2pt(20),
          },
          {
            text: "Wasserbezug / Consommation d'eau",
            width: mm2pt(60),
          },
          {
            text: "CHF " + variableData.wasserbezug.toFixed(2),
            width: mm2pt(30),
            textOptions: { align: "right" },
          },
          {
            text: "CHF " + tableData.wasserbezug.toFixed(2),
            width: mm2pt(30),
            textOptions: { align: "right" },
          },
        ],
      },
      {
        columns: [
          {
            text: "1",
            width: mm2pt(20),
          },
          {
            text: "Jahr / Année",
            width: mm2pt(20),
          },
          {
            text: "Abonnement GF/Abonnement du Journal",
            width: mm2pt(60),
          },
          {
            text: "CHF " + variableData.gfAbonement.toFixed(2),
            width: mm2pt(30),
            textOptions: { align: "right" },
          },
          {
            text: "CHF " + tableData.gfAbonement.toFixed(2),
            width: mm2pt(30),
            textOptions: { align: "right" },
          },
        ],
      },
      {
        columns: [
          {
            text: "1",
            width: mm2pt(20),
          },
          {
            text: "Jahr / Année",
            width: mm2pt(20),
          },
          {
            text: "Strom / elctricité",
            width: mm2pt(60),
          },
          {
            text: "CHF " + variableData.strom.toFixed(2),
            width: mm2pt(30),
            textOptions: { align: "right" },
          },
          {
            text: "CHF " + tableData.strom.toFixed(2),
            width: mm2pt(30),
            textOptions: { align: "right" },
          },
        ],
      },
      {
        columns: [
          {
            text: "1",
            width: mm2pt(20),
          },
          {
            text: "Jahr / Année",
            width: mm2pt(20),
          },
          {
            text: "Versicherung / Assurance",
            width: mm2pt(60),
          },
          {
            text: "CHF " + variableData.versicherung.toFixed(2),
            width: mm2pt(30),
            textOptions: { align: "right" },
          },
          {
            text: "CHF " + tableData.versicherung.toFixed(2),
            width: mm2pt(30),
            textOptions: { align: "right" },
          },
        ],
      },
      {
        columns: [
          {
            text: "1",
            width: mm2pt(20),
          },
          {
            text: "Jahr / Année",
            width: mm2pt(20),
          },
          {
            text: "Mitgliederbeitrag / Cotisation",
            width: mm2pt(60),
          },
          {
            text: "CHF " + variableData.mitgliederbeitrag.toFixed(2),
            width: mm2pt(30),
            textOptions: { align: "right" },
          },
          {
            text:
              "CHF " +
              (tableData.mitgliederbeitrag === 0
                ? "-"
                : tableData.mitgliederbeitrag.toFixed(2)),
            width: mm2pt(30),
            textOptions: { align: "right" },
          },
        ],
      },
      {
        columns: [
          {
            text: "1",
            width: mm2pt(20),
          },
          {
            text: "Jahr / Année",
            width: mm2pt(20),
          },
          {
            text: "Reparatur Fonds / Fonds de réparation",
            width: mm2pt(60),
          },
          {
            text: "CHF " + variableData.reparaturfonds.toFixed(2),
            width: mm2pt(30),
            textOptions: { align: "right" },
          },
          {
            text: "CHF " + tableData.reparaturfonds.toFixed(2),
            width: mm2pt(30),
            textOptions: { align: "right" },
          },
        ],
      },
      {
        columns: [
          {
            text: "1",
            width: mm2pt(20),
          },
          {
            text: "Jahr / Année",
            width: mm2pt(20),
          },
          {
            text: "Verwaltungskosten / frais de gestion",
            width: mm2pt(60),
          },
          {
            text: "CHF " + variableData.verwaltungskosten.toFixed(2),
            width: mm2pt(30),
            textOptions: { align: "right" },
          },
          {
            text: "CHF " + tableData.verwaltungskosten.toFixed(2),
            width: mm2pt(30),
            textOptions: { align: "right" },
          },
        ],
      },
      {
        columns: [
          {
            text: "Total",
            width: mm2pt(20),
          },
          {
            text: "",
            width: mm2pt(20),
          },
          {
            text: "",
            width: mm2pt(60),
          },
          {
            text: "",
            width: mm2pt(30),
          },
          {
            text: "CHF " + tableData.total.toFixed(2),
            width: mm2pt(30),
            textOptions: { align: "right" },
          },
        ],
      },
    ],
  };
}
