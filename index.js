const reader = require('xlsx');
const _ = require("lodash");

const NoonToNamshiMap = reader.readFile('./Crosslisting ^ Mapping & Product Upload.xlsx');

const Headers = reader.utils.sheet_to_json(NoonToNamshiMap.Sheets['HeadMap']);
const CategoryMap = reader.utils.sheet_to_json(NoonToNamshiMap.Sheets['CatMap'], { range: 1 });
let CountryMap = reader.utils.sheet_to_json(NoonToNamshiMap.Sheets['Country_Mapping'], { range: 1 });
let ColorMap = reader.utils.sheet_to_json(NoonToNamshiMap.Sheets['Colour_Mapping'],);

CountryMap = CountryMap.reduce((acc, val) => {
    acc[val["country_of_origin_1"]?.toLowerCase().trim()] = val["country_of_origin"]?.toLowerCase().trim();
    return acc;
}, {});

ColorMap = ColorMap.reduce((acc, val) => {
    acc[val["noon"]?.toLowerCase().trim()] = val["namshi"]?.toLowerCase().trim();
    return acc;
}, {});

// const GenderMap = _.keyBy(reader.utils.sheet_to_json(NoonToNamshiMap.Sheets['Gender_Mapping']), "Noon - Gender");

// console.log('Headers', Headers);

// console.log('GenderMap', GenderMap);
// console.log('CategoryMap', CategoryMap);
// console.log('catMap', catMap);
// console.log('ColorMap', ColorMap);

const namshi_to_noon_headers = Headers.reduce((acc, val, index) => {

    const key = val["Namshi"] ? val["Namshi"] : `Empty${index}`;
    const value = val["noon_header"] ? val["noon_header"] : `Empty${index + 1}`;

    acc[key] = value;

    return acc;

}, {});

// console.log('namshi_to_noon_headers', namshi_to_noon_headers);

// const file2 = reader.readFile('./noon-tested-file.xlsx');
// const noonTestedFile = reader.utils.sheet_to_json(file2.Sheets['MP Format'], { range: 1 });
// console.log('noonTestedFile', noonTestedFile);

const file3 = reader.readFile('./noon-data.xlsx');

const noon_format = reader.utils.sheet_to_json(file3.Sheets['noon_format'], { range: 1 });

const file4 = reader.readFile('./size_runs.xlsx');

const size_runs = reader.utils.sheet_to_json(file4.Sheets['size_runs']);


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


let namshi_format = {};

for (let noon of noon_format) {

    try {

        const item = namshi_format[noon[namshi_to_noon_headers["product_id"]]];

        const sizekey = getSize(noon);

        if (item) {
            item[sizekey] = noon.sku;
            continue;
        }

        const categoryDetail = getCategoryData(noon);

        const namshiObj = {
            "brand_key": noon[namshi_to_noon_headers["brand_key"]],
            "supplier_sku": noon[namshi_to_noon_headers["supplier_sku"]],
            "product_id": noon[namshi_to_noon_headers["product_id"]],
            "color": ColorMap[noon[namshi_to_noon_headers["color"]]?.toLowerCase().trim()],
            "name_en": noon[namshi_to_noon_headers["name_en"]],
            "name_ar": noon[namshi_to_noon_headers["name_en"]],
            "short_description_en": noon[namshi_to_noon_headers["short_description_en"]],
            "short_description_ar": noon[namshi_to_noon_headers["short_description_ar"]],
            "specialist": noon[namshi_to_noon_headers["specialist"]],
            "age_group": categoryDetail["age_group"],
            "gender": categoryDetail["gender"],
            "department": categoryDetail["department"],
            "category": categoryDetail["category"],
            "sub_category": categoryDetail["sub_category"],
            "basic_type": categoryDetail["basic_type"],
            "occasion": categoryDetail["occasion"],
            "product_detail": categoryDetail["product_detail"],
            "country_of_origin": CountryMap[noon[namshi_to_noon_headers["country_of_origin"]]?.toLowerCase().trim()],
            "warranty": noon[namshi_to_noon_headers["warranty"]],
            "dangerous_good_type": noon[namshi_to_noon_headers["dangerous_good_type"]],
            "images": noon[namshi_to_noon_headers["images"]],
            "unit_cost": noon[namshi_to_noon_headers["unit_cost"]],
            "original_price": noon[namshi_to_noon_headers["original_price"]],
            "sizerun": categoryDetail["sizerun"]
        };

        namshiObj[sizekey] = noon.sku;

        namshi_format[namshiObj.product_id] = namshiObj;

    }
    catch (err) {
        console.log('err', err.message);
    }

};

// console.log('namshi_format', namshi_format);
const colNames = namshiHeaders.reduce((acc, val) => {
    acc[val] = val;
    return acc;
}, {});

namshi_format = Object.values(namshi_format);
namshi_format.unshift(colNames);

const ws = reader.utils.json_to_sheet(namshi_format);

const newFile = reader.utils.book_new();

reader.utils.book_append_sheet(newFile, ws, "namshi_format");

reader.writeFile(newFile, './n2n-code-converted.xlsx');

// Functions

function getCategoryData(noon) {

    const categoryDetail = CategoryMap.find(obj => {

        if (obj.department.toLowerCase() == noon.department.toLowerCase() &&
            obj.gender.toLowerCase() == noon.gender.toLowerCase() &&
            obj.family.toLowerCase() == noon.family.toLowerCase() &&
            obj.product_type.toLowerCase() == noon.product_type.toLowerCase() &&
            obj.product_subtype.toLowerCase() == noon.product_subtype.toLowerCase() &&
            obj.basic_type.toLowerCase() == noon.basic_type.toLowerCase() &&
            obj.occasion.toLowerCase() == noon.occasion.toLowerCase() &&
            obj.pattern_type.toLowerCase() == noon.pattern_type.toLowerCase()
        ) return obj;

    });

    if (!categoryDetail) return {};

    return {
        "age_group": categoryDetail["age_group"],
        "gender": categoryDetail["gender_1"],
        "department": categoryDetail["department_1"],
        "category": categoryDetail["category"],
        "sub_category": categoryDetail["sub_category"],
        "basic_type": categoryDetail["basic_type_1"],
        "occasion": categoryDetail["occasion_1"],
        "product_detail": categoryDetail["product_detail"],
        "sizerun": categoryDetail["size_run"]
    };

}

function getSize(noon) {

    const sizeRun = size_runs.find(run => run.size_run.toLowerCase() == "MEN Tops XS-4XL".toLowerCase());

    for (let size in sizeRun) {

        if (sizeRun[size].toLowerCase() == noon.size.toLowerCase()) return size;

    }

};