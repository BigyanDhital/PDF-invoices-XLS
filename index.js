const pdf = require("pdf-creator-node");
const fs = require("fs");
var xlsx = require("node-xlsx");
const fileName = "invoices.xlsx";
let workSheetsFromBuffer;
try {
  console.log(`\n\Loading ${fileName}\n\n`);
  workSheetsFromBuffer = xlsx.parse(fs.readFileSync(`${fileName}`));
} catch (e) {
  console.error("Error loading file\n");
  console.log(e.message);
  setTimeout(() => {
    process.exit();
  }, 2000);
}

if (!workSheetsFromBuffer) return;

const subFolder = "invoices";
const arrayOfData = workSheetsFromBuffer[0].data;
const headers = arrayOfData[0];
const data = arrayOfData.slice(1).filter((row) => {
  let [sn, firstname] = row;
  return !!firstname;
});

console.log(`Printing ${data.length} pdfs`);
// Read HTML Template
const html = fs.readFileSync(`./invoice.html`, "utf8");
const options = {
  timeout: "100000",
  format: "A4",
  orientation: "portrait",
  border: "10mm",
  //   header: {
  //     height: "45mm",
  //     contents: '<div style="text-align: center;">Best Mad Honey</div>',
  //   },
  footer: {
    height: "28mm",
    contents: {
      default:
        '<div style="padding-top:8px; text-align:center; border-top:2px solid #ddd"><span style="color: #444;">Mad Honey Nepal · www.madhoney.net</span><div>', // fallback value
    },
  },
};
console.log("Reading data...");

data.forEach((customer, index) => {
  let [
    SN,
    firstname,
    lastname,
    address,
    city = "",
    state = "",
    zip,
    country,
    phone,
    quantity,
    courier,
    trackingcode,
  ] = customer;

  let price = 3;
  let amount = parseInt(quantity) * price;
  if (!amount) amount = 0;
  var document = {
    html: html,
    headers: { ...headers },
    data: {
      sn: 1,
      date: `${new Date().toDateString()}`,
      firstname,
      lastname,
      address,
      city,
      state,
      zip,
      country,
      phone,
      description: "Sample Honey",
      price,
      amount,
      quantity,
      courier,
      trackingcode,
      empty: "  ",
    },
    path: `./${subFolder}/${index + 1} ${firstname} ${lastname} from ${city}.pdf`,
  };

  try {
    pdf
      .create(document, options)
      .then((res) => {
        let index = res.filename?.lastIndexOf("/");
        const name = res.filename?.slice(index) || "PDF";
        console.log(`Printed ${subFolder}${name}`);
        // console.log(res);
      })
      .catch((error) => {
        console.log(`Error printing`, error);
      });
  } catch (e) {
    console.log(`Error printing: ${e.message}`);
  }
});
