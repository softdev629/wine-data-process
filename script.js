const readXlsxFile = require("read-excel-file/node");
const writeXlsxFile = require("write-excel-file/node");
const fs = require("fs");

// delcares schema of converted json file
const schema = {
  shop_code: {
    prop: "shopCode",
    type: String,
  },
  des_id: {
    prop: "desID",
    type: String,
  },
  description: {
    prop: "description",
    type: String,
  },
  vintage: {
    prop: "vintage",
    type: String,
  },
  rating: {
    prop: "rating",
    type: Number,
  },
  flavour_x: {
    prop: "flavourX",
    type: Number,
  },
  flavour_y: {
    prop: "flavourY",
    type: Number,
  },
  rgb: {
    prop: "rgb",
    type: String,
  },
  wine_type: {
    prop: "wineType",
    type: String,
  },
  foodpair0: {
    prop: "foodpair0",
    type: String,
  },
  foodpair1: {
    prop: "foodpair1",
    type: String,
  },
  foodpair2: {
    prop: "foodpair2",
    type: String,
  },
  foodpair3: {
    prop: "foodpair3",
    type: String,
  },
  foodpair4: {
    prop: "foodpair4",
    type: String,
  },
  foodpair5: {
    prop: "foodpair5",
    type: String,
  },
  foodpair6: {
    prop: "foodpair6",
    type: String,
  },
  collection: {
    prop: "collection",
    type: String,
  },
  subcollection: {
    prop: "subcollection",
    type: String,
  },
  region: {
    prop: "region",
    type: String,
  },
  subregion: {
    prop: "subregion",
    type: String,
  },
  occasion: {
    prop: "occasion",
    type: String,
  },
  ribbon: {
    prop: "ribbon",
    type: String,
  },
  alternative_wine: {
    prop: "alternativeWine",
    type: String,
  },
  vinvalue: {
    prop: "vinValue",
    type: Number,
  },
  color: {
    prop: "color",
    type: String,
  },
  tastingnote: {
    prop: "tastingNote",
    type: String,
  },
  price_sale: {
    prop: "priceSale",
    type: Number,
  },
  vinvaluestore: {
    prop: "vinValueStore",
    type: String,
  },
  pricecompstore: {
    prop: "priceCompStore",
    type: String,
  },
  ratingcompstore: {
    prop: "ratingCompStore",
    type: String,
  },
  flag_onsale: {
    prop: "flagOnSale",
    type: Number,
  },
  flavour_profile: {
    prop: "flavourProfile",
    type: String,
  },
  flavour_description: {
    prop: "flavourDescription",
    type: String,
  },
  flavour_taste1: {
    prop: "flavourTaste1",
    type: String,
  },
  flavour_taste2: {
    prop: "flavourTaste2",
    type: String,
  },
  expand_palate: {
    prop: "expandPalate",
    type: String,
  },
  product_url: {
    prop: "productURL",
    type: String,
  },
};

const HEADER_ROW = [
  {
    value: "wine_name",
    fontWeight: "bold",
  },
  {
    value: "rating",
    fontWeight: "bold",
  },
  {
    value: "flavour_x",
    fontWeight: "bold",
  },
  {
    value: "flavour_y",
    fontWeight: "bold",
  },
  {
    value: "wine_type",
    fontWeight: "bold",
  },
  {
    value: "foodpair",
    fontWeight: "bold",
  },
  {
    value: "region",
    fontWeight: "bold",
  },
  {
    value: "occasion",
    fontWeight: "bold",
  },
  {
    value: "alternative_wine",
    fontWeight: "bold",
  },
  {
    value: "vinvalue",
    fontWeight: "bold",
  },
  {
    value: "colour",
    fontWeight: "bold",
  },
  {
    value: "tastingnote",
    fontWeight: "bold",
  },
  {
    value: "price",
    fontWeight: "bold",
  },
  {
    value: "flag_onsale",
    fontWeight: "bold",
  },
];

const getSweetBoldness = (sweetness, boldness) => {
  if (sweetness === 0 && boldness === 0) return "neutral";
  let degreeStatement = "";

  if (sweetness === -5) degreeStatement += "very sweet";
  if (sweetness > -5 && sweetness < -2) degreeStatement += "sweet";
  if (sweetness > -3 && sweetness < 0) degreeStatement += "semi-sweet";
  if (sweetness > 0 && sweetness < 3) degreeStatement += "semi-dry";
  if (sweetness > 2 && sweetness < 5) degreeStatement += "dry";
  if (sweetness === 5) degreeStatement += "very dry";

  if (boldness === 0) return degreeStatement;
  if (degreeStatement !== "") degreeStatement += " and ";

  if (boldness === -5) degreeStatement += "very light";
  if (boldness > -5 && boldness < -2) degreeStatement += "light";
  if (boldness > -3 && boldness < 0) degreeStatement += "semi-light";
  if (boldness > 0 && boldness < 3) degreeStatement += "semi-bold";
  if (boldness > 2 && boldness < 5) degreeStatement += "bold";
  if (boldness === 5) degreeStatement += "very bold";

  return degreeStatement;
};

readXlsxFile("./given.xlsx", { schema })
  .then(({ rows, errors }) => {
    if (errors.length !== 0) throw errors;
    const docs = rows.map((row) => {
      const text = `About ${row.description} ${row.vintage} wine
    ${row.description} ${row.vintage} has rating of ${row.rating}. ${
        row.description
      } ${row.vintage} is ${getSweetBoldness(row.flavourX, row.flavourY)}. ${
        row.description
      } ${row.vintage} is kind of ${row.wineType}. Food Pair of ${
        row.description
      } ${row.vintage} is : ${row.foodpair0}. High level categorisation of ${
        row.description
      } ${row.vintage} is ${row.collection.split("|")}. ${row.description} ${
        row.vintage
      } is ${row.subcollection}. ${
        row.region !== undefined
          ? `${row.description} ${row.vintage} comes from ${row.subregion} in ${row.region}. `
          : ``
      }${
        row.occasion !== undefined
          ? ` People can drink ${row.description} ${
              row.vintage
            } on the occasion of "${row.occasion
              .split("|")
              .filter((value) => value !== "")}". `
          : ""
      }${
        row.ribbon !== undefined
          ? `${row.description} ${row.vintage} is listed as ${row.ribbon}. `
          : ""
      }${
        row.alternativeWine !== undefined
          ? `${row.alternativeWine} wine can be used instead of ${row.description} ${row.vintage}. `
          : ""
      }${row.description} ${row.vintage} is ${
        row.vinvalue < 0 ? "not " : ""
      }scored as good based on its price. ${row.description} ${
        row.vintage
      } shows ${row.color.toLowerCase()} color. ${row.description} ${
        row.vintage
      } costs $${row.priceSale}. ${row.description} ${
        row.vintage
      } is currently${
        row.flagOnSale === 1 ? "" : " not"
      } on sale. It can be described as ${row.description} ${
        row.vintage
      } is as follows: ${row.tastingNote}. This is production link - <a href="${
        row.productURL
      }">${row.productURL}</a>.`;

      return text;
    });

    // const DATA_ROWS = rows.map((row) => [
    //   { type: String, value: `${row.description} ${row.vintage}` },
    //   { type: Number, value: row.rating },
    //   { type: Number, value: row.flavourX },
    //   { type: Number, value: row.flavourY },
    //   { type: String, value: row.wineType },
    //   { type: String, value: row.foodpair0 },
    //   { type: String, value: `${row.region} ${row.subregion}` },
    //   { type: String, value: row.occasion },
    //   { type: String, value: row.alternativeWine },
    //   { type: Number, value: row.vinValue },
    //   { type: String, value: row.color },
    //   { type: String, value: row.tastingNote },
    //   { type: Number, value: row.priceSale },
    //   { type: Number, value: row.flagOnSale },
    // ]);
    // writeXlsxFile([HEADER_ROW, ...DATA_ROWS], { filePath: "./output.xlsx" });
    fs.writeFileSync("wineinfo.txt", docs.join("\n\n"));
  })
  .catch((err) => console.log(err));
