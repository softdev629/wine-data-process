const readXlsxFile = require("read-excel-file/node");
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
};

readXlsxFile("./given.xlsx", { schema })
  .then(({ rows, errors }) => {
    if (errors.length !== 0) throw errors;
    const docs = rows.map((row) => {
      const text = `About ${row.description} ${row.vintage} wine
${row.description} ${row.vintage} has raiting of ${row.raiting}. ${
        row.description
      } ${row.vintage} is kind of ${row.flavourX > 0 ? "dry" : "sweet"} ${
        row.flavourY > 0 ? "bold" : "light"
      } wine. Type of ${row.description} ${row.vintage} wine is ${
        row.wineType
      }. Foods go well with ${row.description} ${row.vintage} is ${
        row.foodpair1
      }, ${row.foodpair2}, ${row.foodpair3}, ${row.foodpair4} and ${
        row.foodpair5
      }. Foodpair of ${row.description} ${row.vintage} in detail is "${
        row.foodpair0
      }". High level categorisation of ${row.description} ${
        row.vintage
      } wine type is ${row.collection.split("|")}. Simply this belongs to ${
        row.subcollection
      } wine. ${
        row.region !== undefined
          ? `${row.description} ${row.vintage} comes from ${row.subregion} in ${row.region}.`
          : `The region ${row.description} ${row.vintage} comes from is not specified.`
      } ${
        row.occasion !== undefined
          ? ` People can drink ${row.description} ${
              row.vintage
            } on the occasion of "${row.occasion
              .split("|")
              .filter((value) => value !== "")}".`
          : ""
      }${
        row.ribbon !== undefined ? `This wine is listed as ${row.ribbon}.` : " "
      }${
        row.alternativeWine !== undefined
          ? `${row.alternativeWine} wine can be used instead of ${row.description} ${row.vintage}.`
          : ""
      } ${row.description} ${row.vintage} is ${
        row.vinvalue < 0 ? "not " : ""
      }scored as good based on its price. ${row.description} ${
        row.vintage
      } shows ${row.color.toLowerCase()} color. Tasting note of ${
        row.description
      } ${row.vintage} is "${row.tastingNote}". ${row.description} ${
        row.vintage
      } costs $${row.priceSale}. ${row.description} ${
        row.vintage
      } is scored as ${
        row.vinValueStore > 0.5 ? "better" : "worse"
      } than others based on price. ${row.description} ${row.vintage} costs ${
        row.priceCompStore < 0.5 ? "more expensive" : "cheaper"
      } than average price. ${row.description} ${row.vintage} has ${
        row.raitingCompStore > 0.5 ? "better" : "worse"
      } rating than normal ones. ${row.description} ${
        row.vintage
      } is currently${
        row.flagOnSale === 1 ? "" : "not"
      } on sale. Flavour description of ${row.description} ${row.vintage} is "${
        row.flavourDescription
      }. It tastes like ${row.flavourTaste1} and ${row.flavourTaste2}."`;

      return text;
    });
    console.log(docs);
    // fs.writeFileSync("wineinfo.txt", )
  })
  .catch((err) => console.log(err));
