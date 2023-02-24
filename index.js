const reader = require('xlsx');
const _ = require("lodash");

const NoonToNamshiMap = reader.readFile('./Crosslisting ^ Mapping & Product Upload.xlsx');

const Headers = reader.utils.sheet_to_json(NoonToNamshiMap.Sheets['Headers']);
const CategoryMap = reader.utils.sheet_to_json(NoonToNamshiMap.Sheets['Category_Mapping'], { range: 1 });
const catMap = reader.utils.sheet_to_json(NoonToNamshiMap.Sheets['CatMap'], { range: 1 });
const GenderMap = _.keyBy(reader.utils.sheet_to_json(NoonToNamshiMap.Sheets['Gender_Mapping']), "Noon - Gender");

// console.log('GenderMap', GenderMap);
// console.log('CategoryMap', CategoryMap);
console.log('catMap', catMap);

const namshi_to_noon_headers = Headers.reduce((acc, val, index) => {

    const key = val["Namshi"] ? val["Namshi"] : `Empty${index}`;
    const value = val["noon_header"] ? val["noon_header"] : `Empty${index + 1}`;

    acc[key] = value;

    return acc;

}, {});

// console.log('Headers', Headers);
// console.log('namshi_to_noon_headers', namshi_to_noon_headers);

const file2 = reader.readFile('./noon-tested-file.xlsx');

const noonTestedFile = reader.utils.sheet_to_json(file2.Sheets['MP Format'], { range: 1 });

// console.log('Headers', Headers);
// console.log('noonTestedFile', noonTestedFile);

const file3 = reader.readFile('./GMG Product Upload.xlsx');

const noon_format = reader.utils.sheet_to_json(file3.Sheets['noon_format'], { range: 1 });

// console.log('noon_format', noon_format);


const namshiHeaders = ["brand_key",
    "supplier_sku",
    "product_id",
    "color",
    "name_en",
    "name_ar",
    "short_description_en",
    "short_description_ar",
    "specialist",
    "age_group",
    "gender",
    "department",
    "category",
    "sub_category",
    "basic_type",
    "occasion",
    "product_detail",
    "country_of_origin",
    "warranty",
    "dangerous_good_type",
    "images",
    "unit_cost",
    "original_price",
    "sizerun"
];

for (let i = 1; i <= 152; i++) {
    namshiHeaders.push(`S${i}`);
}

// console.log('namshiHeaders', JSON.stringify(namshiHeaders));


const namshi_format = noon_format.map(noon => {

    try {
        const namshiObj = {
            "brand_key": noon[namshi_to_noon_headers["brand_key"]],
            "supplier_sku": noon[namshi_to_noon_headers["supplier_sku"]],
            "product_id": noon[namshi_to_noon_headers["product_id"]],
            "color": noon[namshi_to_noon_headers["color"]],
            "name_en": noon[namshi_to_noon_headers["name_en"]],
            "name_ar": noon[namshi_to_noon_headers["name_en"]],
            "short_description_en": noon[namshi_to_noon_headers["short_description_en"]],
            "short_description_ar": noon[namshi_to_noon_headers["short_description_ar"]],
            "specialist": noon[namshi_to_noon_headers["specialist"]],
            "age_group": GenderMap[noon[namshi_to_noon_headers["gender"]].toLowerCase()].AGE_GROUP,
            "gender": GenderMap[noon[namshi_to_noon_headers["gender"]].toLowerCase()].GENDER,
            "department": noon[namshi_to_noon_headers["department"]],
            "category": noon[namshi_to_noon_headers["category"]],
            "sub_category": noon[namshi_to_noon_headers["sub_category"]],
            "basic_type": noon[namshi_to_noon_headers["basic_type"]],
            "occasion": noon[namshi_to_noon_headers["occasion"]],
            "product_detail": noon[namshi_to_noon_headers["product_detail"]],
            "country_of_origin": noon[namshi_to_noon_headers["country_of_origin"]],
            "warranty": noon[namshi_to_noon_headers["warranty"]],
            "dangerous_good_type": noon[namshi_to_noon_headers["dangerous_good_type"]],
            "images": noon[namshi_to_noon_headers["images"]],
            "unit_cost": noon[namshi_to_noon_headers["unit_cost"]],
            "original_price": noon[namshi_to_noon_headers["original_price"]],
            "sizerun": noon[namshi_to_noon_headers["sizerun"]]
        };

        const categoryDetail = getCategoryData(noon);

        // console.log('categoryDetail', categoryDetail);

        return namshiObj;
    }
    catch (err) {
        console.log('err', err.message);
    }

});

// console.log('namshiHeaders', namshiHeaders);

const colNames = namshiHeaders.reduce((acc, val) => {
    acc[val] = val;
    return acc;
}, {});

namshi_format.unshift(colNames);

console.log('namshi_format', namshi_format);

const ws = reader.utils.json_to_sheet(namshi_format);

const newFile = reader.utils.book_new();

reader.utils.book_append_sheet(newFile, ws, "namshi_format");

reader.writeFile(newFile, './namshi_format.xlsx');

// Functions

function getCategoryData(noon) {

    const categoryDetail = CategoryMap.find(obj => {

        if (obj.product_type == noon.product_type && obj.product_subtype == noon.product_subtype) return obj;

    })

    return categoryDetail;

}


